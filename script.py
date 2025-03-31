import smtplib
import pandas as pd
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def get_ms_forms_responses(sheet_link):
    """Fetch responses from an Excel/Google Sheet link."""
    responses = pd.read_excel(sheet_link)  # Fetch data from MS Forms response sheet
    return responses

def check_availability(user_slot, booked_slots):
    """Check if the requested slot is available."""
    return user_slot not in booked_slots

def send_email(to_email, subject, body):
    """Send an email notification."""
    sender_email = "gs.barak@iitg.ac.in"  # Update with actual email
    sender_password = "Mks-573100"  # Update with actual credentials
    cc_emails = ["gs.barak@iitg.ac.in"]
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['CC'] = ", ".join(cc_emails)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, [to_email] + cc_emails, msg.as_string())

def process_bookings(sheet_link):
    """Process all form responses and send emails accordingly."""
    responses = get_ms_forms_responses(sheet_link)
    booked_slots = set(responses.iloc[:, 14].dropna())  # Get existing booked slots from the 15th column
    
    for index, row in responses.iterrows():
        user_email = row.iloc[3]  # 4th column (zero-based index 3) contains email
        user_slot = row.iloc[14]  # 15th column (zero-based index 14) contains slot data
        
        if check_availability(user_slot, booked_slots):
            # Confirm booking
            booked_slots.add(user_slot)
            subject = "Pool Table Booking Confirmed"
            body = f"Your pool table slot ({user_slot}) has been successfully booked."
        else:
            # Inform user that the slot is taken
            subject = "Pool Table Slot Unavailable"
            body = f"The slot ({user_slot}) you requested is already occupied. Please choose another slot."
        
        send_email(user_email, subject, body)

if __name__ == "__main__":
    sheet_link = "https://forms.office.com/Pages/RedirectToExcelPage.aspx?id=jacKheGUxkuc84wRtTBwHKBMNLlkXLtFva55P_XrQdRUQzRJVzZTTEhXUFJXMlZBTkVMWjNUSDRNNC4u&forceReExport=false&isSdxDataSync=true"  # Updated sheet link
    while True:
        process_bookings(sheet_link)
        time.sleep(600)  # Runs every 10 minutes
