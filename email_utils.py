import yagmail
import os

def send_analysis_email(recipient_email, excel_data, filename):
    # Setup your credentials here or use environment variables
    SENDER_EMAIL = "your-email@gmail.com"
    SENDER_PASSWORD = "your-app-password" 

    try:
        yag = yagmail.SMTP(SENDER_EMAIL, SENDER_PASSWORD)
        
        # We write the BytesIO content to a temporary file to send as attachment
        temp_path = f"temp_{filename}"
        with open(temp_path, "wb") as f:
            f.write(excel_data)

        contents = [
            f"Please find the attached Real Estate Analysis Report: {filename}",
            temp_path
        ]
        
        yag.send(to=recipient_email, subject="Property Analysis Report", contents=contents)
        
        # Clean up
        os.remove(temp_path)
        return True, "Email sent successfully!"
    except Exception as e:
        return False, f"Error: {str(e)}"
