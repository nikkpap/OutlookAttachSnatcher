# OutlookAttachSnatcher (Outlook Bulk Attachment Downloader)

A simple VBA macro for **Classic Microsoft Outlook for Windows** that allows you to download attachments from multiple selected emails at once.

Instead of opening each email and saving its attachments manually, select multiple emails, run the macro, choose how you want the files organized, and let the macro do the rest.

## Features

- Download attachments from multiple selected Outlook emails
- Select any number of emails at once
- Choose between:
  - **Single Folder** – save all attachments together
  - **Separate Folders** – create a folder for each email
- Separate folders are automatically named using:
  - Received date and time
  - Sender name
  - Email subject
- Automatically handles duplicate filenames
- Existing files are never overwritten
- Invalid Windows filename characters are automatically cleaned
- Displays the total number of downloaded attachments
- No external software or libraries required

## Requirements

- Windows
- Classic Microsoft Outlook
- VBA / Macros enabled

> **Important:** This macro does not work with the New Outlook for Windows.  
> You must switch to **Classic Outlook** to use VBA macros.

## Installation

1. Open **Classic Microsoft Outlook**.
2. Press `Alt + F11` to open the VBA Editor.
3. Select:

   `Insert → Module`

4. Paste the VBA code into the new module.
5. Press `Ctrl + S` to save.
6. Close the VBA Editor.

## Usage

1. Open Outlook.
2. Select multiple emails using `Ctrl + Click` or `Shift + Click`.
3. Press `Alt + F8`.
4. Select:

   `SaveAttachmentsFromSelectedEmails`

5. Click **Run**.

The macro will ask:

```text
Theleis ola ta attachments se ENAN fakelo?

YES = Ola mazi se enan fakelo
NO = Xehoristos fakelos gia kathe email
CANCEL = Akyro
