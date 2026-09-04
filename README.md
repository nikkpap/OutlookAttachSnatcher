# 📎 OutlookAttachSnatcher

**Bulk-download and organize Outlook attachments without opening emails one by one.**

OutlookAttachSnatcher is a VBA macro for **Classic Microsoft Outlook for Windows** that can download attachments from multiple emails in a single operation.

You can either select emails manually or search the current Outlook folder using filters such as sender, email address, subject, and date range.

The downloader includes progress tracking, duplicate protection, temporary-file safety, cancellation support, disk-space checking, and handling designed for large batches of emails and attachments.

---

## ✨ Features

- 📥 Bulk-download attachments from multiple Outlook emails
- 🔎 Search the current Outlook folder automatically
- 👤 Filter by **sender name OR sender email address**
- 📝 Filter by text contained in the **Subject**
- 📅 Optional **From / To** date filtering
- 📂 Save all attachments into one folder
- 🗂️ Or create a separate folder for each email
- 🔢 Automatically handles duplicate filenames
- 🛡️ Never intentionally overwrites an existing file
- ⏳ Live download progress
- 📊 Shows current attachment / total attachments
- 📈 Shows percentage completed
- 💾 Calculates total attachment size before starting
- 🚨 Checks available disk space before downloading
- 🧹 Cleans invalid Windows filename characters
- 📏 Protects against excessively long Windows paths
- 🗃️ Uses temporary `.downloading` files while saving
- ❌ Failed attachments do not stop the entire batch
- ⌨️ Press `ESC` to cancel a running operation
- ⚡ Uses cooperative UI yielding to keep Outlook responsive
- 🧠 Avoids keeping large numbers of Outlook `MailItem` objects in memory
- 📦 Everything is contained in **one VBA module**
- 🔧 No external libraries or additional software required

---

# 🖥️ Requirements

- Windows
- **Classic Microsoft Outlook for Windows**
- VBA / Macros enabled

> [!IMPORTANT]
> OutlookAttachSnatcher does **not** work with the **New Outlook for Windows** because New Outlook does not support traditional Outlook VBA macros.
>
> Switch to **Classic Outlook** before installing or running the macro.

---

# 📦 Installation

## 1. Open Classic Outlook

If you are currently using New Outlook, switch back to:

**Classic Outlook**

---

## 2. Open the VBA Editor

Press:

```text
ALT + F11
```

The Microsoft Visual Basic for Applications editor will open.

---

## 3. Create a Module

In the VBA Editor select:

```text
Insert → Module
```

A new standard VBA module will appear.

---

## 4. Add OutlookAttachSnatcher

Open:

```text
OutlookAttachSnatcher.bas
```

Copy the complete source code and paste it into the new Outlook VBA module.

Alternatively, if your VBA editor supports importing modules:

```text
File → Import File...
```

and select:

```text
OutlookAttachSnatcher.bas
```

---

## 5. Save

Press:

```text
CTRL + S
```

Close the VBA Editor.

---

# ▶️ Running OutlookAttachSnatcher

There are two convenient ways to run it.

## Method 1 — Keyboard

Press:

```text
ALT + F8
```

Select:

```text
OutlookBulkAttachmentDownloader
```

and click:

```text
Run
```

---

# 🔘 Add OutlookAttachSnatcher as a Button

For regular use, adding the macro as an Outlook button is much faster than using `ALT + F8`.

## Option A — Quick Access Toolbar

In **Classic Outlook**:

1. Open:

   ```text
   File → Options
   ```

2. Select:

   ```text
   Quick Access Toolbar
   ```

3. In:

   ```text
   Choose commands from:
   ```

   select:

   ```text
   Macros
   ```

4. Find the macro:

   ```text
   OutlookBulkAttachmentDownloader
   ```

   Depending on your VBA project/module name, Outlook may display a longer name such as:

   ```text
   Project1.Module1.OutlookBulkAttachmentDownloader
   ```

5. Select the macro.

6. Click:

   ```text
   Add >>
   ```

7. Select the newly added macro on the right.

8. Click:

   ```text
   Modify...
   ```

9. Choose an icon.

10. Change the display name to something friendly, for example:

   ```text
   OutlookAttachSnatcher
   ```

11. Click:

   ```text
   OK
   ```

12. Click:

   ```text
   OK
   ```

You now have a dedicated OutlookAttachSnatcher button in the Quick Access Toolbar.

---

## Option B — Add it to the Outlook Ribbon

You can also create a dedicated group in the Outlook Ribbon.

Open:

```text
File → Options → Customize Ribbon
```

On the right side:

1. Select the **Home** tab.
2. Click **New Group**.
3. Rename the group to:

   ```text
   Attachments
   ```

On the left side:

4. Change:

   ```text
   Choose commands from:
   ```

   to:

   ```text
   Macros
   ```

5. Select:

   ```text
   OutlookBulkAttachmentDownloader
   ```

6. Click:

   ```text
   Add >>
   ```

7. Use **Rename** to change its displayed name to:

   ```text
   OutlookAttachSnatcher
   ```

8. Choose an icon if desired.
9. Click **OK**.

You can now launch the downloader directly from:

```text
Home → Attachments → OutlookAttachSnatcher
```

---

# 📥 Download Modes

When OutlookAttachSnatcher starts, it asks:

```text
Choose how you want to process emails:

YES = Process currently selected emails
NO = Search current Outlook folder
CANCEL = Exit
```

---

## YES — Process Selected Emails

Select multiple emails in Outlook using:

```text
CTRL + Click
```

or:

```text
SHIFT + Click
```

Then run OutlookAttachSnatcher.

Only the selected emails will be processed.

This is useful when you already know exactly which messages contain the attachments you need.

---

## NO — Search Current Outlook Folder

OutlookAttachSnatcher can search the currently open Outlook folder automatically.

You can filter using:

- Sender name
- Sender email address
- Subject
- Start date
- End date

---

# 🔎 Sender Search

The Sender field accepts either a person's display name or their email address.

For example:

```text
Name Surname
```

or:

```text
namesurname@example.com
```

Partial searches are also supported:

```text
surname
```

The comparison is case-insensitive.

The macro checks both:

```text
SenderName
```

and:

```text
SenderEmailAddress
```

For Exchange accounts, it also attempts to resolve the sender's primary SMTP address.

---

# 📝 Subject Search

You can enter part of an email subject.

For example:

```text
Daily Report
```

This can match subjects such as:

```text
Daily Report 20260624
Daily Report Project A
Site Daily Report
DAILY REPORT 2026-06-24
```

Subject matching is case-insensitive.

Leave the Subject field blank if you do not want to filter by subject.

---

# 📅 Date Filtering

The date range is optional.

Example:

```text
From Date:
01/06/2026

To Date:
30/06/2026
```

The entire final day is included.

Leave either field blank if that limit is not required.

---

# 📂 Attachment Organization

After choosing the emails, OutlookAttachSnatcher asks:

```text
How should attachments be organized?

YES = Save all attachments in one folder
NO = Create a separate folder for each email
CANCEL = Exit
```

---

## 📁 Single Folder

Selecting **YES** saves every attachment into the selected destination.

Example:

```text
Attachments/
├── Drawing.pdf
├── Report.xlsx
├── Photo.jpg
├── Drawing_1.pdf
└── Report_1.xlsx
```

---

## 🗂️ Separate Folder for Each Email

Selecting **NO** creates a separate directory for each email.

Example:

```text
Attachments/
│
├── 2026-06-24_09-15 - Name Surname - Daily Report/
│   ├── Daily_Report.pdf
│   └── Drawing.pdf
│
├── 2026-06-25_09-20 - Name Surname - Daily Report/
│   ├── Daily_Report.pdf
│   └── Site_Photo.jpg
│
└── 2026-06-26_09-18 - Name Surname - Daily Report/
    └── Daily_Report.pdf
```

The folder naming format is:

```text
YYYY-MM-DD_HH-MM - Sender - Subject
```

Invalid Windows filename characters are automatically replaced.

---

# 🔢 Duplicate Filename Protection

OutlookAttachSnatcher does not intentionally overwrite existing files.

If the destination already contains:

```text
Report.pdf
```

additional files are automatically renamed:

```text
Report.pdf
Report_1.pdf
Report_2.pdf
Report_3.pdf
```

This is especially useful when downloading attachments from recurring reports where every attachment uses the same filename.

---

# ⏳ Progress Tracking

Before downloading, OutlookAttachSnatcher scans the matching messages.

For example:

```text
Matched emails: 184
Total attachments: 300
Total size: 4.82 GB
```

You can confirm the operation before any files are downloaded.

During processing, live progress is displayed.

Example:

```text
[##########--------------] 127 / 300 - 42%
```

The current attachment can also be displayed:

```text
Downloading: Daily_Report_20260624.pdf (184.70 MB)
```

---

# ⌨️ Cancel a Download

Press:

```text
ESC
```

during processing to request cancellation.

The downloader stops at a safe point rather than deliberately terminating Outlook.

> [!NOTE]
> `Attachment.SaveAsFile` is a synchronous Outlook operation.
>
> If Outlook is currently writing one very large attachment, cancellation takes effect after that individual save operation returns.

---

# 🛡️ Large File Safety

Attachments are first written using a temporary filename.

For example:

```text
Daily_Report.pdf.downloading
```

After Outlook successfully saves the attachment, the file is renamed to:

```text
Daily_Report.pdf
```

This reduces the chance of an interrupted or failed operation leaving a partial file that appears to be a successfully downloaded attachment.

Failed temporary files are cleaned up where possible.

---

# 💾 Disk Space Check

Before downloading, OutlookAttachSnatcher calculates the combined reported size of the matching attachments.

A safety margin is added when checking available local disk space.

For example:

```text
Total attachments: 300
Total size: 4.82 GB
Safety margin: 10%
```

If there does not appear to be enough free space, the operation is stopped before downloading.

For UNC/network locations, free-space reporting is not always reliable, so the downloader does not block the operation solely because it cannot determine network free space.

---

# 📏 Long Filename Protection

Windows and Outlook can encounter problems when a combination of:

```text
Destination Path
+ Sender
+ Subject
+ Attachment Filename
```

becomes excessively long.

OutlookAttachSnatcher automatically limits generated filename lengths to reduce these failures.

---

# 🚀 Large Mailboxes and Large Batches

OutlookAttachSnatcher is designed to handle larger operations without unnecessarily keeping hundreds or thousands of Outlook message objects alive.

During the initial scan it stores lightweight identifiers for matching messages and retrieves the messages again when needed for processing.

It also periodically calls:

```text
DoEvents
```

to allow Windows and Outlook to process UI events during long operations.

This helps keep Outlook more responsive when processing large batches.

---

# 🧵 Why Isn't It Multithreaded?

OutlookAttachSnatcher intentionally processes Outlook attachments sequentially.

Classic Outlook VBA uses the Outlook Object Model through COM. Outlook objects such as:

```text
MailItem
Attachment
AddressEntry
```

should not be treated as safely parallelizable VBA objects.

Attempting to force multithreaded Outlook VBA processing could make the application less reliable rather than faster.

Instead, OutlookAttachSnatcher uses:

- Sequential attachment processing
- Cooperative UI yielding
- Lightweight message references
- Per-file error handling
- Temporary-file protection
- Cancellation support
- Pre-scanning
- Disk-space validation

The priority is **reliability and data safety**.

For genuinely parallel, server-side attachment processing, a separate implementation using technologies such as Microsoft Graph or another appropriate API would be a different architecture.

---

# 📊 Completion Summary

At the end of a job, OutlookAttachSnatcher displays a summary.

Example:

```text
Operation finished.

Matched emails: 184
Total attachments: 300
Total size: 4.82 GB
Downloaded: 298
Failed: 2
Processed: 300 / 300

Status: Completed
```

If the user cancels:

```text
Status: Cancelled by user
```

---

# ⚠️ Macro Security

Microsoft Outlook may prevent VBA macros from running depending on your Office security configuration or your organization's policies.

Only enable and run macros from sources you trust.

You can inspect the complete source code in:

```text
OutlookAttachSnatcher.bas
```

before running it.

> [!WARNING]
> Do not globally weaken Office security settings simply to run a macro, especially on a managed work computer.
>
> If macros are disabled by your organization, contact your IT administrator.

---

# 🌐 Unicode / International Filenames

The macro does not intentionally convert attachment filenames, email subjects, or sender names to ASCII/Greeklish.

Windows-compatible Unicode filenames can therefore remain in their original language where supported by Outlook/VBA and the filesystem.

The user-interface text contained directly in the VBA source is English for maximum compatibility across different Windows and Office language configurations.

---

# 🧰 Troubleshooting

### The macro does not appear in `ALT + F8`

Make sure the main procedure is declared as:

```vb
Public Sub OutlookBulkAttachmentDownloader()
```

and that the code is inside a **standard Module**, not `ThisOutlookSession`.

---

### The button does nothing

Check that:

- You are using Classic Outlook
- VBA macros are permitted
- The macro still exists in the VBA project
- The button is assigned to `OutlookBulkAttachmentDownloader`

---

### Search returns no emails

Check:

- You are inside the correct Outlook folder
- Sender spelling
- Email address
- Subject text
- Date range

Remember that automatic searching currently operates on the **current Outlook folder**.

---

### Does it search subfolders?

No.

The current version searches the folder that is currently open in Outlook.

Subfolder/recursive mailbox searching is not currently enabled.

---

### A file failed to download

The downloader will count the attachment as failed and continue with the remaining files where possible.

This prevents one problematic attachment from stopping an entire large batch.

---

# 📄 Project Structure

The repository can remain very simple:

```text
OutlookAttachSnatcher/
│
├── OutlookAttachSnatcher.bas
├── README.md
└── LICENSE
```

No external dependencies are required.

---

# 🔧 Compatibility

| Feature | Supported |
|---|---|
| Classic Outlook for Windows | ✅ |
| New Outlook for Windows | ❌ |
| Multiple selected emails | ✅ |
| Search by sender name | ✅ |
| Search by sender email | ✅ |
| Subject filtering | ✅ |
| Date filtering | ✅ |
| Duplicate protection | ✅ |
| Separate folders | ✅ |
| Progress counter | ✅ |
| ESC cancellation | ✅ |
| Large batch handling | ✅ |
| True multithreading | ❌ |
| Recursive subfolder search | ❌ |

---

# 🤝 Contributing

Bug reports, improvements, suggestions, and pull requests are welcome.

Potential future improvements include:

- Recursive subfolder searching
- Download log / CSV report
- Attachment extension filters
- Attachment size filters
- Saved search presets
- Improved progress UI
- Resume/retry failed downloads

---

# 📜 License

This project is licensed under the **GNU General Public License v3.0 (GPL-3.0)**.

See the `LICENSE` file for details.

---

# ⭐ OutlookAttachSnatcher

If OutlookAttachSnatcher saves you from manually opening hundreds of emails and clicking **Save As** all day, consider giving the repository a ⭐.

Happy attachment snatching. 📎
