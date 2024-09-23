import os
import re
import uuid
import logging
import sys
import shutil
import win32com.client as win32
from email.utils import parseaddr
from win32com.client import Dispatch, pywintypes
from concurrent.futures import ThreadPoolExecutor

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def create_folder(output_dir):

    # Import necessary libraries
    shell = win32.Dispatch("WScript.Shell")

    # Create an instance of the Windows Script Host Shell object
    documentPath = shell.SpecialFolders("MyDocuments")

    # Retrieve the path to the "My Documents" folder using the Shell object
    folderPath = os.path.join(documentPath, output_dir)

    return os.path.join(folderPath, '')

def is_valid_email(email: str) -> bool:
    """
    Validates an email address using email.utils.parseaddr.
    Args:
        email (str): Email address to validate.
    Returns:
        bool: True if valid, False otherwise.
    """
    return '@' in parseaddr(email)[1]

def is_outlook_installed() -> bool:
    """
    Check if Outlook is installed by attempting to dispatch an Outlook instance.
    Returns:
        bool: True if Outlook is installed, False otherwise.
    """
    # Clear any cached COM modules to avoid conflicts
    MODULE_LIST = [m.__name__ for m in sys.modules.values()]
    for module in MODULE_LIST:
        if re.match(r'win32com\.gen_py\..+', module):
            del sys.modules[module]
    shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'), ignore_errors=True)

    try:
        win32.gencache.EnsureDispatch('Outlook.Application')
        return True
    except pywintypes.com_error as e:
        logging.error(f"COM error when accessing Outlook: {e}")
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
    return False

def send_email_via_outlook(
    to: str,
    subject: str,
    body: str,
    attachments: list = None,
    html_body: bool = False,
    cc: list = None,
    bcc: list = None,
    embedded_images: list = None
) -> str:
    """
    Send an email via Outlook.

    Args:
        to (str): Recipient's email address.
        subject (str): Email subject.
        body (str): Email body.
        attachments (list, optional): List of file paths to attach.
        cc (list, optional): List of CC recipients' email addresses.
        bcc (list, optional): List of BCC recipients' email addresses.
        html_body (bool, optional): Set to True to send body as HTML content.
        embedded_images (list, optional): List of image paths to embed in the email body.

    Returns:
        str: Status message.
    """
    if not is_outlook_installed():
        return "Outlook is not installed or accessible."

    # Validate email addresses
    if not is_valid_email(to):
        return f"Invalid email addresses provided: {to}"

    try:
        # Create a new instance of Outlook
        outlook = win32.Dispatch('Outlook.Application')

        # Create a new email item
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject

        # Add CC and BCC recipients, if any
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc

        # Set email body (HTML or plain text)
        if html_body:
            mail.HTMLBody = body
        else:
            mail.Body = body

        # Attach embedded images, if provided
        if embedded_images:
            for image_path in embedded_images:
                attachment = mail.Attachments.Add(image_path)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001E",
                    os.path.basename(image_path)
                )

        # Add file attachments, if any
        if attachments:
            for attachment in attachments:
                attch_path = os.path.abspath(attachment["path"])
                if os.path.exists(attch_path):
                    mail.Attachments.Add(attch_path)
                else:
                    logging.warning(f"Attachment not found: {attch_path}")

        # Send the email
        mail.Send()
        logging.info("Email sent successfully.")
        return "Email sent successfully."

    except Exception as e:
        logging.error(f"Error sending email: {e}")
        return f"Error sending email: {e}"

def get_emails(
    email: str,
    subject: str,
    folder_name: str,
    output_dir: str,
    allowed_file_types: list = None,
    include_read: bool = False,
    since=None
) -> list:
    """
    Retrieve emails from an Outlook folder based on specific criteria.

    Args:
        email (str): Email account address.
        subject (str): Subject filter.
        folder_name (str): Folder name to search emails in (e.g., Inbox).
        output_dir (str): Directory to save email attachments.
        allowed_file_types (list, optional): Allowed attachment file extensions.
        include_read (bool, optional): Whether to include read emails.
        since (datetime, optional): Only retrieve emails received after this date.

    Returns:
        list: A list of email details dictionaries.
    """
    emails = []

    # Create the folder for saving attachments
    folderPath = create_folder(output_dir)

    if not os.path.isdir(folderPath):
        os.makedirs(folderPath, exist_ok=True)
        logging.info(f"Created directory: {folderPath}")

    # Validate the provided email address
    if not is_valid_email(email):
        logging.error(f"Invalid email addresses provided: {email}")
        return emails

    try:
        outlook = win32.Dispatch('Outlook.Application')
        mapi = outlook.GetNamespace("MAPI")

        # Get the folder (e.g., Inbox)
        found_folder = next(
            (
                folder for folder in mapi.GetDefaultFolder(6).Folders
                if folder.Name == folder_name
            ),
            None
        )
        if not found_folder:
            logging.error(f"Folder '{folder_name}' not found.")
            return emails

        # Retrieve messages filtered by subject
        messages = found_folder.Items.Restrict(f"@SQL=(urn:schemas:httpmail:subject LIKE '%{subject}%')")

        # Add date filter if provided
        if since:
            date_filter = since.strftime('%Y-%m-%d %H:%M:%S')
            messages = messages.Restrict(f"[ReceivedTime] >= '{date_filter}'")

        # Sort messages by received time (newest first)
        messages.Sort("[ReceivedTime]", True)

        # Process messages
        for msg in messages:
            if not include_read and msg.UnRead or include_read:
                try:
                    process_email(msg, emails, folderPath, allowed_file_types)
                except Exception as e:
                    logging.error(f"Error processing individual email: {e}")

    except Exception as e:
        logging.error(f"Error when processing email messages: {e}")

    logging.info(f"Processed {len(emails)} emails.")
    return emails

def process_email(
    msg,
    emails_list: list,
    output_dir: str,
    allowed_file_types: list = None
) -> None:
    """
    Process a single email and save its attachments.

    Args:
        msg: The Outlook email message object.
        emails_list (list): List to append email details to.
        output_dir (str): Directory to save email attachments.
        allowed_file_types (list, optional): List of allowed file extensions.
    """
    try:
        # List to store attachment details
        attachment_details = []

        # Process and save email attachments
        allowed_file_types = allowed_file_types or []
        with ThreadPoolExecutor() as executor:
            futures = [
                executor.submit(download_attachment, att, output_dir, allowed_file_types)
                for att in msg.Attachments
            ]
            for future in futures:
                try:
                    attachment_detail = future.result()
                    if attachment_detail:
                        attachment_details.append(attachment_detail)
                except Exception as e:
                    logging.error(f"Error saving attachment: {e}")

        # Collect email details
        email_detail = {
            'name': msg.Subject,
            'subject': msg.Subject,
            'conversation_id': msg.ConversationID,
            'from': msg.Sender.GetExchangeUser().PrimarySmtpAddress,
            'sender': msg.Sender.Name,
            'body': msg.Body,
            'files': attachment_details,
            'status': "pending"
        }

        # Append the email details to the list
        emails_list.append(email_detail)

    except Exception as e:
        logging.error(f"Error processing email: {e}")

def download_attachment(attachment, output_dir: str, allowed_file_types: list) -> dict:
    """
    Downloads an attachment and returns its details.

    Args:
        attachment: The Outlook attachment object.
        output_dir (str): Directory to save the attachment.
        allowed_file_types (list): List of allowed file extensions.

    Returns:
        dict: Details of the saved attachment, or None if it was not saved.
    """
    file_extension = os.path.splitext(attachment.FileName)[1]
    if not allowed_file_types or file_extension in allowed_file_types:
        try:
            # Save the attachment
            attachment_path = os.path.join(output_dir, attachment.FileName)
            attachment.SaveAsFile(attachment_path)
            logging.info(f"Attachment {attachment.FileName} saved")

            # Return attachment details
            return {
                'path': attachment_path,
                'name': attachment.FileName,
                'id': str(uuid.uuid4())
            }
        except Exception as e:
            logging.error(f"Error saving attachment {attachment.FileName}: {e}")
            return None
    return None
