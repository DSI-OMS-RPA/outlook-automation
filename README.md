# Email Automation with Outlook

This Python project automates sending and retrieving emails from Microsoft Outlook using `pywin32`.

## Features

- **Send Emails**: Automatically send emails through Outlook with support for attachments, CC/BCC, HTML body, and embedded images.
- **Retrieve Emails**: Fetch emails from a specific Outlook folder, filter by subject, and download any attachments.
- **Save Attachments**: Automatically saves email attachments to a specified directory.

## Requirements

- **Microsoft Outlook**: The script requires Outlook to be installed and configured on your system.
- **Python 3.x**
- `pywin32`: Python library to interact with Outlook via COM.

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/Kowts/outlook-automation.git
   ```
2. **Create a virtual environment** (optional but recommended):

   ```bash
   python -m venv venv
   source venv/bin/activate      # On Linux/MacOS
   venv\Scripts\activate         # On Windows
   ```
3. **Install the required dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

The script can be used for both sending emails and retrieving emails with attachments.

### Sending Emails

To send an email, use the `send_email_via_outlook` function from `main.py`. Here's an example of how to call the function:

```python
from main import send_email_via_outlook

send_email_via_outlook(
    to="recipient@example.com",
    subject="Test Email",
    body="This is a test email.",
    attachments=[{'path': 'file.pdf'}],
    cc=["cc@example.com"],
    bcc=["bcc@example.com"],
    html_body=True,
    embedded_images=["image1.png"]
)
```

### Retrieving Emails

To retrieve emails, use the get_emails function from main.py:

```python
from main import get_emails

emails = get_emails(
    email="your-email@example.com",
    subject="Project Update",
    folder_name="Inbox",
    output_dir="attachments",
    allowed_file_types=[".pdf", ".docx"],
    include_read=True
)
```
This will save any attachments from emails matching the subject and store them in the attachments folder.

## Configuration

- Ensure that Outlook is properly installed and configured on your system before running the script.
- The script relies on `pywin32` to interface with Outlook, so be sure to install the dependency via `requirements.txt`.

## Contributing

Feel free to fork the repository, make changes, and submit a pull request. Any contributions are welcome!

## License

This project is licensed under the MIT License.
