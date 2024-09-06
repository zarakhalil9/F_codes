import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from flask import Flask, request, render_template, redirect, url_for, flash
import os

app = Flask(__name__)
app.secret_key = os.urandom(24)  # Securely generated secret key

# Email setup
smtp_server = "smtp.gmail.com"
smtp_port = 587
smtp_user = ""  # Your Gmail address
smtp_password = ""  # Your Gmail app password

def send_bulk_emails(excel_file_path):
    df = pd.read_excel(excel_file_path)
    errors = []

    for _, row in df.iterrows():
        email = row['email']
        try:
            # Create email
            msg = MIMEMultipart()
            msg['From'] = formataddr(('Your Name', smtp_user))
            msg['To'] = email
            msg['Subject'] = "Your Payment Link"

            body = f"""Dear {row['customer_name']},

Here is your payment link for order {row['order_id']}.

Amount: {row['amount']} {row['currency']}
Transaction Type: {row['transaction_type']}
Expiry Date: {row['payment_expiry_date']}
Payment Link: {row['payment_url']}

Thank you!
"""
            msg.attach(MIMEText(body, 'plain'))

            # Send email
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(smtp_user, smtp_password)
                server.send_message(msg)
            print(f"Email sent to {email}")

        except Exception as e:
            errors.append(f"Failed to send email to {email}: {e}")

    return errors

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if file and file.filename.endswith('.xlsx'):
            file_path = os.path.join("uploads", file.filename)
            file.save(file_path)
            
            errors = send_bulk_emails(file_path)

            if errors:
                for error in errors:
                    flash(error, 'error')
                return redirect(url_for('index'))

            flash('Emails sent successfully!', 'success')
            return redirect(url_for('index'))
        else:
            flash('Invalid file format. Please upload an Excel file.', 'error')
    
    return render_template("index.html")

if __name__ == "__main__":
    if not os.path.exists("uploads"):
        os.makedirs("uploads")
    app.run(debug=True)

