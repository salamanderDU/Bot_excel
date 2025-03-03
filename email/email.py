import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import datetime

def create_ics(event_details):
    """Create an ICS file for the calendar event."""
    ics_content = f"""BEGIN:VCALENDAR
VERSION:2.0
CALSCALE:GREGORIAN
BEGIN:VEVENT
UID:{event_details['uid']}
SUMMARY:{event_details['summary']}
LOCATION:{event_details['location']}
DESCRIPTION:{event_details['description']}
DTSTART:{event_details['start'].strftime('%Y%m%dT%H%M%S')}
DTEND:{event_details['end'].strftime('%Y%m%dT%H%M%S')}
STATUS:CONFIRMED
SEQUENCE:0
BEGIN:VALARM
TRIGGER:-PT15M
DESCRIPTION:Reminder for {event_details['summary']}
ACTION:DISPLAY
END:VALARM
END:VEVENT
END:VCALENDAR
"""
    return ics_content

def send_email_with_attachment(smtp_server, port, sender_email,recipient_email, subject, body, attachment):
    """Send an email with an ICS attachment."""
    try:
        # Create a multipart message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        # Attach the email body

        msg.attach(MIMEText(body, 'plain'))

        # Attach the ICS file
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        f'attachment; filename="event.ics"')
        msg.attach(part)

        # Create an SMTP session and send the email
        with smtplib.SMTP(smtp_server, port) as server:
            server.set_debuglevel(1)
            response = server.sendmail(sender_email, recipient_email, msg.as_string())
            print("Response: "+str(response))
            print("Email sent successfully with calendar event!")

    except Exception as e:
        print(f"Error occurred: {e}")

if __name__ == "__main__":
    # SMTP server configuration
    #SMTP_SERVER = 'botmail.bot.or.th' #EWS ต้องใข้ตัวนี้
    SMTP_SERVER = 'mailgateway1.bot.or.th'  # Replace with your SMTP server
    PORT = 25                         # For TLS
    SENDER_EMAIL = 'services@bot.or.th'  # Replace with your email
    RECIPIENT_EMAIL = 'suphakit_k@mfec.co.th'
    SUBJECT = 'Upcoming Event Notification'    # Subject of the email
    BODY = 'Please find the attached calendar event.\n\nBest regards,'  # Body of the email

    # Define the event details
    event_details = {
        'uid': 'suphakit_k@mfec.co.th',  # Unique identifier for the event
        'summary': 'Team Meeting',
        'location': 'Online',
        'description': 'Monthly team meeting to discuss project updates.',
        'start': datetime.datetime(2025, 1, 31, 10, 0),  # Start time
        'end': datetime.datetime(2025, 1, 31, 11, 0)     # End time
    }

    # Create the ICS content
    ics_content = create_ics(event_details)

    # Send the email with the ICS attachment
    send_email_with_attachment(
        SMTP_SERVER, PORT, SENDER_EMAIL,RECIPIENT_EMAIL, SUBJECT, BODY, ics_content)
