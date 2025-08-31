from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

app = Flask(__name__)

# Serve index.html
@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

# Serve CSS
@app.route('/style.css')
def style():
    return send_from_directory('.', 'style.css')

# Serve JS
@app.route('/script.js')
def script():
    return send_from_directory('.', 'script.js')

@app.route('/send_emails', methods=['POST'])
def send_emails():
    try:
        excel_file = request.files['excel_file']
        sender_email = request.form['sender_email'].strip().replace(u'\xa0', ' ')
        sender_password = request.form['sender_password'].strip().replace(u'\xa0', ' ')
        email_subject = request.form['subject'].strip()
        email_content = request.form['content']

        # Save Excel temporarily
        temp_path = 'temp_emails.xlsx'
        excel_file.save(temp_path)

        # Read Excel file
        df = pd.read_excel(temp_path)
        if df.empty:
            return jsonify({"success": False, "error": "Excel file is empty"}), 400

        email_column = df.columns[0]  # First column assumed as email list
        total_emails = len(df)
        results = []

        # SMTP setup
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()

        try:
            server.login(sender_email, sender_password)
        except Exception as e:
            return jsonify({"success": False, "error": f"Login failed: {str(e)}"}), 401

        for _, row in df.iterrows():
            recipient_email = str(row[email_column]).strip()
            try:
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient_email
                # Ensure UTF-8 for subject
                msg['Subject'] = str(email_subject).encode('utf-8').decode('utf-8')

                # Attach message body with UTF-8 encoding
                msg.attach(MIMEText(email_content, 'plain', 'utf-8'))

                server.send_message(msg)
                results.append({'email': recipient_email, 'status': 'Success', 'message': 'Email sent successfully'})
                del msg
            except Exception as e:
                results.append({'email': recipient_email, 'status': 'Failed', 'message': str(e)})

        server.quit()

        if os.path.exists(temp_path):
            os.remove(temp_path)

        return jsonify({
            'success': True,
            'results': results,
            'total': total_emails,
            'successful': len([r for r in results if r['status'] == 'Success'])
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True)
