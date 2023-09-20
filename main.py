import re
import csv
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
import pdfkit


# Define the file name for the CSV for the logging process
csv_filename = "email_results.csv"



# Read data from CSV file
with open('Questionnaire_Distance_Instructor_course_evaluation_REST_REST__ (2).csv', 'r', encoding='utf-8') as file:
    reader = csv.DictReader(file)
    data = list(reader)

# Run Different Function based on the File name.
# The Distance are divided by 46 for the Total and the Conv is divided by 17

def Dist():
    for idx, row in enumerate(data):
        shortname = row['shortname']
        if shortname not in grouped_data:
            grouped_data[shortname] = []
        grouped_data[shortname].append(row)
        avg = float(row['average_rankvalue'].replace(',', '.'))

        if idx % 46 == 0:
            avg_sum = 0
        avg_sum = avg_sum + avg
        str_avg_values[shortname] = round(avg_sum / 46, 2)  # Store str_avg value for each shortname

def Conv():
    for idx, row in enumerate(data):
        shortname = row['shortname']
        if shortname not in grouped_data:
            grouped_data[shortname] = []
        grouped_data[shortname].append(row)
        avg = float(row['average_rankvalue'].replace(',', '.'))

        if idx % 17 == 0:
            avg_sum = 0
        avg_sum = avg_sum + avg
        str_avg_values[shortname] = round(avg_sum / 17, 2)  # Store str_avg value for each shortname




# Group data by shortname
grouped_data = {}
i = 0
avg_sum = 0
str_avg_values = {}  # Dictionary to store str_avg values for each shortname

if "Dist" in file.name:
    Dist()
elif "Conv" in file.name:
    Conv()


# Generate and save HTML files
for shortname, rows in grouped_data.items():
    # Generate HTML content
    html_content = f'''
   <!DOCTYPE html>
    <html>
    <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
         <link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Lato&display=swap" rel="stylesheet">
        <style>
            /* Global styles */
            body {{
                font-family: 'Lato', sans-serif;
                margin: 0;
                padding: 3% 15% 0 15%;
                
            }}

            /* Responsive styles */
            @media (max-width: 600px) {{
                body {{
                    padding: 10px;
                }}
            }}

            /* Table styles */
            table {{
                width: 100%;
                border-collapse: collapse;
            }}

            th, td {{
                padding: 8px;
                text-align: left;
                border: 1px solid #000;
            }}

            td:empty {{
                border: 0px solid #ddd;
                border-top: 1px solid #ddd;
            }}

            th {{
                background-color: #f2f2f2;
            }}

            /* Responsive table styles */
            @media (max-width: 600px) {{
                th, td {{
                    font-size: 14px;
                }}
            }}

            #results td {{
                border: 1px solid #000;
            }}
        </style>
    </head>
    <body>
            <img src="https://www.nup.ac.cy/wp-content/uploads/2021/07/Neapolis-Logo-EN.png" alt="" decoding="async" data-lazy-src="nuplogo.png" data-ll-status="loaded" style="width:40%;margin-bottom: 2%;">
        <table>
        <tbody>
        <tr>
        <td>Όνομα Μαθήματος: <b>{rows[0]['fullname']}</b></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        </tr>
        <tr>
        <td>Όνομα Διδάσκοντος: <b>{rows[0]['teacherName']}</b></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        </tr>
        <tr>
        <td><b>{rows[0]['Department']}</b></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        </tr>
        <tr>
        <td><b>{rows[0]['Semester']}</b></td>
        <td></td>
        <td></td>
        <td></td>
        <td>Ημερομηνία παράδοσης αξιολόγησης στο ΓΔΠ:29/06/2023</td>
        </tr>
        <tr>
        <td></td>
        <td></td>
        <td></td>
        <td>Εγγεγραμμένοι στο μάθημα:</td>
        <td>Απάντησαν στο ερωτηματολόγιο:</td>
        </tr>
        <tr>
        <td></td>
        <td></td>
        <td></td>
        <td><b>{rows[0]['totalinclass']}</b></td>
        <td><b>{rows[0]['tookpart']}</b></td>
        </tr>
        </tbody>
        </table>

        <p style="padding: 1%;text-align: center;font-weight: 600;">
        <span style="">Περιγραφή αξιολόγησης:</span><br><br>
        5 = Strongly Agree, 4 = Agree, 3 = Neither Agree nor Disagree, 2 = Disagree, 1 = Strongly Disagree
        </p>
        <table id="results">
        <tbody>
        {''.join(f"<tr><td>{row['question']}</td><td>{row['content']}</td><td>{row['average_rankvalue']}</td></tr>" for row in rows)}
        <tr><td></td><td><b>TOTAL</b></td><td>{str_avg_values[shortname]}</td></tr>
        </tbody>
        </table>
    </body>
    </html>
    '''

    # Replace invalid characters in filename
    filename = re.sub(r'[\/:*?"<>|]', '_', f"{shortname}.html")
    # if((filename == "ACCN101_S23.html") or (filename == "CPSYC514_S23.html") or (filename == "PSYC410-EN_S23 .html") or (filename == "MBA574-EN_S23 .html")) :
    print(filename)
    # Save HTML file
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(html_content)


        # Generate PDF from HTML
        pdf_filename = re.sub(r'.html', '', f"{filename}")
        pdf_filename = f"{pdf_filename}.pdf"

        # Define options for PDF generation (including page width)
        options = {
            'page-size': 'A4',
            'encoding': 'UTF-8',
            'margin-top': '0',
            'margin-right': '0',
            'margin-bottom': '0',
            'margin-left': '0',
            'zoom': '1.0',
            'viewport-size': '1920x1080'  # Adjust the width as needed
        }
        config = pdfkit.configuration(wkhtmltopdf='C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')

        pdfkit.from_file(filename, pdf_filename, options=options,configuration=config)

        # Send email
        smtp_server = 'smtp.office365.com'
        smtp_port = 587

        sender_email = 'it_support@nup.ac.cy'
        sender_password = 'Giannakis345#$!'

        receiver_email = rows[0]['email']
        # receiver_email= 'p.papagiannis@nup.ac.cy' #Test
        subject = 'Αποτελέσματα αξιολόγησης από τους φοιτητές – Εαρινό εξ. 2023 (update)'
        body = 'Παρακαλώ όπως βρείτε συνημμένα τα αποτελέσματα αξιολόγησης από τους φοιτητές για το εαρινό εξάμηνο 2023. '

        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject
        message['Body'] = body

        # Attach PDF file
        with open(pdf_filename, 'rb') as file:
            attachment = MIMEBase('application', 'pdf')
            attachment.set_payload(file.read())
            encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_filename))
        message.attach(attachment)
        # Attach the body content to the message
        message.attach(MIMEText(body, 'plain'))

        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(message)
            status = "Success"
        except Exception as e:
            status = f"Failed: {str(e)}"

        # Write the results to the CSV file
        with open(csv_filename, mode='a', newline='') as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerow([receiver_email,rows[0]['teacherName'], rows[0]['shortname'], status])
    # Clean up - Delete the generated HTML OR PDF files
    os.remove(filename)
    # os.remove(pdf_filename)







