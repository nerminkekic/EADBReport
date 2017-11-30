"""
This script will obtain four peaces of information for all ASP sites.
    1. Total number of Archives
    2. Total number or Retrieves
    3. Total number of Archives in GB
    4. Total number of Retrieves in GB
"""
import pyodbc
import datetime
import smtplib
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders


def ea_monthly_report():

    """
    This function will do 4 tasks:
    1.  Connect to SQL Server.
    2.  Run query and retrieve data from SQL server Data Base.
    3.  Write the data to Excel worksheet.
    4.  Email the worksheet to end user.
    """

    # File name for Excel Worksheet
    excel_filename = ""

    # Get all Virtual Archive names from SQL Server
    virtual_archives = get_archives()

    # Set up Excel Worksheet.
    work_book = Workbook()
    work_sheet = work_book.active
    work_sheet.title = "EA Monthly Report"
    work_sheet.append(['Archive Name', 'Year', 'Month', 'Total Exams', 'Total Size in MB', 'Average Exam Size in MB', 'Total Size in GB', 'Average Exam Size in GB'])

    # Assign font and background color properties for Column Title cells
    f = Font(name="Arial", size=14, bold=True, color="FF000000")
    fill = PatternFill(fill_type="solid", start_color="00FFFF00")

    # Format Worksheet columns
    for L in "ABCDEFGH":
        work_sheet[L + "1"].fill = fill
        work_sheet[L + "1"].font = f
        if L in "AD":
            work_sheet.column_dimensions[L].width = 25.0
        if L in "BC":
            work_sheet.column_dimensions[L].width = 10.0
        if L in "EFGH":
            work_sheet.column_dimensions[L].width = 35.0

    # Obtain Exam Volume for all virtual archives and write the data to excel sheet.
    for archive in virtual_archives:

        # Variable used to sum up totals for Exam and storage volume
        sum_of_exams = 0
        sum_of_mb = 0
        sum_of_gb = 0

        # Obtains archive volume from SQL server.
        rows = archive_volume(archive)
        for row in rows:
            average_exam_volume_in_gb = round((average_exam_size(row[2], row[3]) / 1024), 6)
            # Adds archive name and exam volume to Worksheet.
            work_sheet.append([archive,                             # Archive Name
                               row[0],                              # Study By Year
                               row[1],                              # Study By Month
                               int(row[2]),                         # Total Exams
                               round(row[3], 2),                    # Total Size in MB
                               average_exam_size(row[2], row[3]),   # Average Exam Size in MB
                               exam_size_in_gb(row[3]),             # Total Size in GB
                               average_exam_volume_in_gb            # Average Exam Size in GB
                               ])
            # Format cells to use 1000 comma separator.
            work_sheet['C{}'.format(work_sheet.max_row)].style = 'Comma [0]'
            work_sheet['D{}'.format(work_sheet.max_row)].style = 'Comma'
            work_sheet['E{}'.format(work_sheet.max_row)].style = 'Comma'
            work_sheet['F{}'.format(work_sheet.max_row)].style = 'Comma'

            # Add total Exam Volume, Total Size in MB and GB
            sum_of_exams += row[2]
            sum_of_mb += row[3]
            sum_of_gb += exam_size_in_gb(row[3])


        # Add sum volumes at the end of the report
        work_sheet.append(["", "", "Total", sum_of_exams, sum_of_mb,"", sum_of_gb])

        # Format cells to use 1000 comma separator.
        work_sheet['D{}'.format(work_sheet.max_row)].style = 'Comma'
        work_sheet['E{}'.format(work_sheet.max_row)].style = 'Comma'
        work_sheet['G{}'.format(work_sheet.max_row)].style = 'Comma'

        print("Added {} Exam Volume to Workbook!".format(archive))
        # Add blank line
        work_sheet.append([])

    # Saves Excel worksheet.
    excel_filename = "ASP_Monthly_Report_{}.xlsx".format(datetime.datetime.now().strftime("%Y-%m-%d"))
    work_book.save(excel_filename)

    # Send email with attachment.
    # send_email(excel_filename)
    #


# Obtain Archive Names from SQL Server.
def get_archives():
    """
    This function will return list of Virtual Archive Names from SQL Server.
    """
    # Create list to store archive names.
    archive_list = []

    # Read file for credentials
    with open("data.txt", "r") as f:
        read_data = f.readline().split()

    # Define data base connection parameters.
    sqlserver = 'SQL1'
    database = 'RSAdmin'
    username = read_data[0]
    password = read_data[1]

    # Establish DB connections.
    conn = pyodbc.connect(
        r'DRIVER={SQL Server};'
        r'SERVER=' + sqlserver + ';'
        r'DATABASE=' + database + ';'
        r'UID=' + username + ';'
        r'PWD=' + password + ''
    )
    cur = conn.cursor()
    # Execute query on Data Base.
    cur.execute("""
                SELECT DBName from tblArchive
                ORDER BY DBName
                """)
    rows = cur.fetchall()

    # Add Archive names to archive list.
    for row in rows:
        archive_list.append(row[0])
    # Close SQL Connection.
    cur.close()
    conn.close()

    return archive_list


# Obtain archive volume.
def archive_volume(db_name):
    """
    This function will obtain archive volume form SQL server.
    """

    # Read file for credentials
    with open("data.txt", "r") as f:
        read_data = f.readline().split()

    # Define data base connection parameters.
    sqlserver = 'SQL1'
    database = db_name
    username = read_data[0]
    password = read_data[1]

    # Establish DB connections.
    conn = pyodbc.connect(
        r'DRIVER={SQL Server};'
        r'SERVER='+sqlserver+';'
        r'DATABASE='+database+';'
        r'UID='+username+';'
        r'PWD='+password+''
        )
    cur = conn.cursor()
    # Execute query on Data Base.
    cur.execute("""
                set transaction isolation level read uncommitted
                select year(firstarchivedate) as studyyear, month(firstarchivedate) as studymonth,
                count(distinct id1) as StudyCount, sum(bytesize/1024/1024) as sumofMB
                from ((tbldicomstudy left join tbldicomseries on tbldicomstudy.id1=tbldicomseries._id1)left join  tblfile on tbldicomseries.id2=tblfile._id2file)
                where firstarchivedate > '2000-01-01' and firstarchivedate < getdate()
                group by year(firstarchivedate), month(firstarchivedate)
                order by year(firstarchivedate), Month(firstarchivedate)
                """)
    rows = cur.fetchall()
    # Close SQL Server Connections.
    cur.close()
    conn.close()
    return rows


# Send email with Report
def send_email(file_attachment):
    """This function will send email with the attachment.
    It takes attachment file name as argument.
    """

    # Define email body
    body = "This is EA Monthly report. See attached file for Total Exam Volume for each customer."
    content = MIMEText(body, 'plain')

    # Open file attachment
    filename = file_attachment
    infile = open(filename, "rb")

    # Set up attachment to be send in email
    part = MIMEBase("application", "octet-stream")
    part.set_payload(infile.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment", filename=filename)

    msg = MIMEMultipart("alternative")

    # Define email recipients
    to_email = ["na@na.com"
        ]
    # Define From email
    from_email = "na@na.com"

    # Create email content
    msg["Subject"] = "ASP Monthly Report {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    msg["From"] = from_email
    msg["To"] = ",".format(to_email)
    msg.attach(part)
    msg.attach(content)
    # Send email to SMTP server
    s = smtplib.SMTP("10.4.1.1", 25)
    s.sendmail(from_email, to_email, msg.as_string())
    s.close()


# Perform calculation for Average Exam size in MB
def average_exam_size(exam_volume, total_storage_size):
    """Calculate average exam size in MB"""
    return round((total_storage_size / exam_volume), 2)


# Convert from MB to GB
def exam_size_in_gb(size_in_mb):
    """Convert Average exam size in MB to GB"""
    return round((size_in_mb / 1024), 2)


# Format Data Base time to use format MM-YYY
def format_date(db_date):
    """Format Data Base time to use format MM-YYY
    """
    try:
        dt = datetime.date(int(db_date[0:4]), int(db_date[4:6]), 1)
    except TypeError:
        print("There is no valid Study Date in Data base!")
    except Exception as e:
        print(e)
    else:
        return dt.strftime("%b{}%y".format("-"))


# Run script
ea_monthly_report()
