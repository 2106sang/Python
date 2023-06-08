import imaplib
import email
import os
from email.header import decode_header

# Set up your Gmail account information
username = "sangxyz1997@gmail.com"
app_password = "app_password"

# Connect to the Gmail IMAP server
server = imaplib.IMAP4_SSL("imap.gmail.com")
server.login(username, app_password)

# Select the "[Gmail]/Sent Mail" folder
server.select('"[Gmail]/Sent Mail"')

# Search for all sent emails
status, sent_messages = server.search(None, "ALL")

# Create a directory to store the exported emails
export_directory = "sent_emails"
if not os.path.exists(export_directory):
    os.makedirs(export_directory)

# Export sent emails
if status == "OK":
    sent_message_ids = sent_messages[0].split()
    for sent_message_id in sent_message_ids:
        _, sent_msg_data = server.fetch(sent_message_id, "(RFC822)")
        sent_msg = email.message_from_bytes(sent_msg_data[0][1])

        # Get the subject of the email and decode it
        subject = sent_msg["Subject"]
        subject_decoded = decode_header(subject)[0][0]
        if isinstance(subject_decoded, bytes):
            subject_decoded = subject_decoded.decode("utf-8")

        # Sanitize the file name by removing invalid characters
        sanitized_subject = "".join(c for c in subject_decoded if c.isalnum() or c in "._- ")
        sanitized_subject = sanitized_subject.strip()

        # Save the email as a text file
        file_path = os.path.join(export_directory, f"{sanitized_subject}.txt")
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(f"Subject: {subject_decoded}\n")
            file.write(f"From: {sent_msg['From']}\n\n")

            # Iterate over the message parts
            for part in sent_msg.walk():
                content_type = part.get_content_type()
                disposition = part.get("Content-Disposition")

                # Check if the part is text/plain
                if content_type == "text/plain" and not disposition:
                    payload = part.get_payload(decode=True)
                    if payload:
                        file.write(payload.decode("utf-8"))
                
                # Check if the part is an attachment
                elif disposition:
                    filename = part.get_filename()
                    if filename:
                        # Sanitize the file name by removing invalid characters
                        sanitized_filename = "".join(c for c in filename if c.isalnum() or c in "._- ")
                        sanitized_filename = sanitized_filename.strip()
                        attachment_path = os.path.join(export_directory, sanitized_filename)
                        with open(attachment_path, "wb") as attachment_file:
                            attachment_file.write(part.get_payload(decode=True))
        
        print(f"Exported email: {subject_decoded}")
else:
    print("Failed to retrieve sent emails.")

# Close the connection to the Gmail IMAP server
server.close()
