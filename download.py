import os
import email
import imapclient
from imapclient.exceptions import IMAPClientError
from email.header import decode_header
from datetime import datetime 
import ssl
import socket
import time
import traceback

def download_pdf_attachments(imap_server, email_address, password):
    # Connect to the IMAP server
    print(f"Connecting to the IMAP server: {imap_server}")
    ssl_context = ssl.create_default_context()
    mail = imapclient.IMAPClient(imap_server, ssl=True, ssl_context=ssl_context)
    
    processed_messages = set()

    # Initialize processed_messages_file and last_processed_uid_file
    processed_messages_file = "processed_messages.txt"
    last_processed_uid_file = "last_processed_uid.txt"

    try:
        mail.login(email_address, password)
        print(f"Logging in as {email_address}")

        # Select the mailbox (e.g., 'inbox')
        mail.select_folder("inbox")

        # Load the highest UID processed from a file
        last_processed_seq = 0
        if os.path.exists(last_processed_uid_file):
            with open(last_processed_uid_file, "r") as file:
                last_processed_seq = int(file.read().strip())


        # Load the list of processed email message IDs from a file
        if os.path.exists(processed_messages_file):
            with open(processed_messages_file, "r") as file:
                processed_messages = set(file.read().strip().split('\n'))
                

        
        while True:
            # Search for all emails
            messages = mail.search(['UID', f'{last_processed_seq + 1}:*'])
            print(f"IMAP Search Result: {messages}")
            
            # messages = mail.search(['SUBJECT', 'Your Subject Here'])
            if not messages:
                print("No new emails. Everything is up-to-date.")
                return

            # Fetch and download attachments for each new message
            for msg_id in messages:
                print(f"Processing new message with ID: {msg_id}")
                    
                try:
                    fetch_response = mail.fetch(msg_id, ["RFC822","INTERNALDATE","UID"])
                    message_data = fetch_response.get(msg_id, {}).get(b"RFC822")
                    current_uid = fetch_response[msg_id].get(b"UID", None)

                    if current_uid is not None:
                        last_processed_seq = max(last_processed_seq, current_uid)

                    
                    if message_data is None:
                        print(f"Skipping message with ID {msg_id}. Could not retrieve email message.")
                        continue

                    
                    # msg_data = fetch_response[msg_id].get(b"RFC822") 
                    msg = email.message_from_bytes(message_data)
                    received_timestamp = datetime.fromtimestamp(
                        email.utils.mktime_tz(email.utils.parsedate_tz(
                            fetch_response.get(msg_id, {}).get(b"INTERNALDATE", b"").strftime("%a, %d %b %Y %H:%M:%S %z")
                        ))
                    )

                    print(f"Processing email with subject: {msg['Subject']}")

                    for part in msg.walk():
                        print("Checking attachment...")
                        print(f"Content Type: {part.get_content_type()}")
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                    
                        # Modified line to handle NoneType from decode_header
                        filename_info = part.get_filename()
                        if filename_info is not None:
                            filename_info = decode_header(filename_info)
                            filename = filename_info[0][0] if filename_info and filename_info[0] and filename_info[0][0] else "Unknown_File"


                            if isinstance(filename, bytes):
                                filename = filename.decode(errors='replace')

                            # Initialize the file extension
                            _, extension = os.path.splitext(filename)
                            extension = extension.lower()
                        
                        
                            # Check if the file already exists in the download path
                            base_filename, _ = os.path.splitext(filename)
                            formatted_timestamp = received_timestamp.strftime('%Y%m%d%H%M%S')
                            safe_base_filename = ''.join(c for c in base_filename if c.isalnum() or c in ('_', '-'))
                            file_path = os.path.join(r"D:\project\resumes", f"{safe_base_filename}_{formatted_timestamp}{extension}")
                        
                            if extension == '.pdf':
                                if os.path.exists(file_path):
                                    print(f"File {filename} already exists. Skipping.")
                                else:
                                # file_path = os.path.join(r"D:\project\py", f"{filename}")
                                    with open(file_path, "wb") as f:
                                        f.write(part.get_payload(decode=True))
                                    print(f"File {filename} downloaded successfully.")
                    
                            else:
                                print("No filename information found for the attachment.")

                            # Mark the message ID as processed and append it to the file
                            processed_messages.add(str(msg_id))

                            # Save the last processed UID to the file
                            with open(last_processed_uid_file, "w") as file:
                                file.write(str(last_processed_seq))
                            print(f"Last processed UID saved: {last_processed_seq}")

                            # Save the processed message IDs to the file
                            with open(processed_messages_file, "a") as file:
                                file.writelines(str(msg_id) + "\n" )  
                
                except (IMAPClientError, ValueError, socket.error) as e:
                    print(f"Error processing message with Seq Num {msg_id}: {e}")
                    traceback.print_exc()
                    print("Attempting to reconnect...")

                    # mail.logout()
                    try:
                        mail.logout()
                    except Exception as e:
                        print(f"Error during logout: {e}")
                    time.sleep(5)  # Wait for a few seconds before reconnecting
                    mail = imapclient.IMAPClient(imap_server, ssl=True, ssl_context=ssl_context)
                    mail.login(email_address, password)
                    mail.select_folder("inbox")

                    continue

            with open(last_processed_uid_file, "w") as file:
                file.write(str(last_processed_seq))
            print(f"Last processed UID saved: {last_processed_seq}")

            # Save the processed message IDs to the file
            with open(processed_messages_file, "a") as file:
                file.writelines(str(msg_id) + "\n" for msg_id in messages)
                                
    except IMAPClientError as e:
        print(f"IMAP Error: {e}")
    
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        traceback.print_exc()

    finally:
    
        # Logout from the server
        print("Logging out from the IMAP server")
        try:
            mail.logout()
        except Exception as e:
            print(f"Error during logout: {e}")
        print("Before logout. Closing connection.")

if __name__ == "__main__":
    imap_server = "mail.demo.com"
    email_address = "ak.demo.com"
    password = "ADfeS@DmI!!398"
    print("Starting script...")
    download_pdf_attachments(imap_server, email_address, password)
    print("Script completed.")

