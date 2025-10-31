from flask import Flask, render_template, request, flash
import smtplib
from email.mime.text import MIMEText
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Change to something secret

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/skills')
def skills():
    return render_template('skills.html')

@app.route('/education')
def education():
    return render_template('education.html')

@app.route('/resume')
def resume():
    return render_template('resume.html')

@app.route('/internships')
def internships():
    return render_template('internships.html')

@app.route('/projects')
def projects():
    return render_template('projects.html')

@app.route('/certifications')
def certifications():
    return render_template('certifications.html')

@app.route('/contact', methods=['GET', 'POST'])
def contact():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        feedback = request.form['feedback'].replace('\r\n', '\n').replace('\r', '\n')

        # Try saving to Excel
        try:
            excel_path = "feedback.xlsx"
            if os.path.exists(excel_path):
                wb = load_workbook(excel_path)
                ws = wb.active
            else:
                wb = Workbook()
                ws = wb.active
                ws.append(['Name', 'Email', 'Feedback'])  # header

            ws.append([name, email, feedback])
            for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True)
            wb.save(excel_path)
        except Exception as excel_e:
            flash(f"Failed to save feedback in Excel: {excel_e}", "danger")
            print(f"Excel error: {excel_e}")
            return render_template('contact.html')

        # Try sending email
        try:
            msg = MIMEText(f'Name: {name}\nEmail: {email}\nFeedback:\n{feedback}')
            msg['Subject'] = 'Portfolio Feedback'
            msg['From'] = email
            msg['To'] = 'iaviswanadhveera@gmail.com'
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                server.login('iamviswanadhveera@gmail.com', 'rzue roit eymm wbfo')
                server.sendmail(email, 'iaviswanadhveera@gmail.com', msg.as_string())
            flash('Feedback sent and stored successfully!', 'success')
        except Exception as e:
            flash(f'Feedback saved, but email failed: {e}', 'warning')
            print(f"Email error: {e}")

    return render_template('contact.html')


# Only ONE main block at the bottom!
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
