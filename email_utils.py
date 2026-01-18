import yagmail
import os

def send_analysis_email(recipient_email, excel_data, filename):
    # Replace with your Gmail and an "App Password" 
    # (Generated in Google Account > Security > 2-Step Verification > App Passwords)
    SENDER_EMAIL = "atharvaujoshi@gmail.com"
    SENDER_PASSWORD = "nybl zsnx zvdw edqr" 

    try:
        yag = yagmail.SMTP(SENDER_EMAIL, SENDER_PASSWORD)
        
        # Save BytesIO to a temp file for attachment
        temp_path = f"temp_{filename}"
        with open(temp_path, "wb") as f:
            f.write(excel_data)

        contents = [
            f"Hello,\n\nPlease find the attached Property Analysis Report: {filename}",
            temp_path
        ]
        
        yag.send(to=recipient_email, subject="Property Analysis Professional Report", contents=contents)
        
        # Cleanup
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return True, "Email sent successfully!"
    except Exception as e:
        return False, f"Email failed: {str(e)}"
