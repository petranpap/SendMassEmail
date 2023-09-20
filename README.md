# Python -Mass Emails

Send Mass Email From Python with Data from CSV

## Installation

Use the package manager [pip](https://pip.pypa.io/en/stable/) to install all the libraries.

```bash
pip install re
pip install csv
pip install smtplib
pip install os
pip install MIMEMultipart
pip install MIMEBase
pip install encoders
pip install MIMEText
pip install pdfkit

```

## Usage

```python
1) Create a csv file to store the results
# Define the file name for the CSV for the logging process
csv_filename = "email_results.csv"

2) Define the CSV file
# Read data from CSV file
with open('Questionnaire_Distance_Instructor_course_evaluation_DPSYC).csv', 'r', encoding='utf-8') as file:
    reader = csv.DictReader(file)
    data = list(reader)
#Check the file names, based on different file name it runs different functions
if "Dist" in file.name:
    Dist()
elif "Conv" in file.name:
    Conv()
# The Distance are divided by 46 for the Total and the Conv is divided by 17

3) install wkhtmltopdf and add the exe location
config = pdfkit.configuration(wkhtmltopdf='C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')

4) add the correct email data
smtp_server = 'smtp.office365.com'
smtp_port = 587
sender_email = 'it_support@nup.ac.cy'
sender_password = ''

# If it crashes it is propably because the teacher name is written with Greek Chars and it cannot write to the csv
writer.writerow([receiver_email,rows[0]['teacherName'], rows[0]['shortname'], status])


## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.

Please make sure to test before running with your personal email.

## License

[MIT](https://choosealicense.com/licenses/mit/)