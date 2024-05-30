from flask import Flask, request, render_template
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email_validator import validate_email, EmailNotValidError
import os
import time

app = Flask(__name__)

# Configuration variables
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SMTP_USER = 'automationpython799@gmail.com'
SMTP_PASSWORD = 'xqtf udtf qozt synt'
COLUMN_HEADERS = [
    "Instructor", "Semester", "Term", "Course", "Subplatform", "Days",
    "Time", "Location", "Sum of Course Count", "Count of Off Load", "Email"
]
EMAIL_SENT_FOLDER = 'email_sent'

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            df = pd.read_excel(file, skiprows=2)
            df = preprocess_dataframe(df)
            if df.empty:
                return render_template('upload.html', error="No valid email found in the uploaded file or 'Grand Total' is at the top.")
            grouped_data = get_schedule_by_professor(df)
            server = setup_smtp_server()
            results, success_emails, failed_emails = send_emails(grouped_data, server)
            server.quit()
            save_emails_as_html(grouped_data)  # Save emails after sending
            return render_template('results.html', results=results)
    return render_template('upload.html')

def preprocess_dataframe(dataframe):
    dataframe.columns = COLUMN_HEADERS
    dataframe['Email'] = dataframe['Email'].ffill()
    dataframe['Valid Email'] = dataframe['Email'].apply(validate_email_address)
    dataframe = dataframe[dataframe['Valid Email']]  # Keep only rows with valid emails
    dataframe.drop(columns='Valid Email', inplace=True)

    # Stop processing at the 'Grand Total' marker in the Instructor column
    if 'Grand Total' in dataframe['Instructor'].values:
        end_index = dataframe[dataframe['Instructor'] == 'Grand Total'].index[0]
        dataframe = dataframe.loc[:end_index - 1]

    return dataframe

def validate_email_address(email):
    try:
        validate_email(email)
        return True
    except EmailNotValidError:
        return False

def get_schedule_by_professor(dataframe):
    grouped = []
    temp_group = []
    current_email = None
    for index, row in dataframe.iterrows():
        if pd.notnull(row['Email']):
            if row['Email'] != current_email:
                if current_email is not None:
                    grouped.append((current_email, pd.DataFrame(temp_group)))
                    temp_group = []
                current_email = row['Email']
            temp_group.append(row)
    if temp_group:
        grouped.append((current_email, pd.DataFrame(temp_group)))
    return grouped

def compose_email(recipient_email, group):
    msg = MIMEMultipart()
    msg['From'] = SMTP_USER
    msg['To'] = recipient_email
    msg['Subject'] = 'Your Teaching Schedule for Next Year'
    instructor_name = group['Instructor'].iloc[0]
    group = group.drop(columns='Email')  # Drop the email column from the group
    html = f"""
    <html><head><style>
    table {{width: 100%; border-collapse: collapse;}}
    th, td {{border: 1px solid #dddddd; padding: 8px; text-align: left;}}
    th {{background-color: #f2f2f2;}}
    p {{font-size: 16px;}}
    </style></head><body>
    <p>Dear {instructor_name},</p>
    <p>This is your teaching schedule for the upcoming academic year. Please review the details below:</p>
    {group.to_html(index=False, na_rep=' ')}
    </body></html>
    """
    msg.attach(MIMEText(html, 'html'))
    return msg

def send_emails(grouped_data, server):
    success_emails = []
    failed_emails = []
    results = []
    for email, group in grouped_data:
        msg = compose_email(email, group)
        try:
            server.send_message(msg)
            success_emails.append(email)
            results.append({'name': group['Instructor'].iloc[0], 'status': 'Success'})
        except Exception as e:
            failed_emails.append(email)
            results.append({'name': group['Instructor'].iloc[0], 'status': 'Failed', 'error': str(e)})
    return results, success_emails, failed_emails

def setup_smtp_server():
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(SMTP_USER, SMTP_PASSWORD)
    return server

def save_emails_as_html(grouped_data):
    timestamp = int(time.time())
    folder_path = os.path.join(EMAIL_SENT_FOLDER, str(timestamp))
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    for idx, (email, group) in enumerate(grouped_data):
        html_content = compose_email(email, group).as_string()
        file_path = os.path.join(folder_path, f'email_{idx}_{email}.html')
        with open(file_path, 'w') as file:
            file.write(html_content)

if __name__ == '__main__':
    app.run(debug=True)
