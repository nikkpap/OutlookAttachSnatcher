# 📎 OutlookAttachSnatcher

**Bulk-download and organize Outlook attachments without opening emails one by one.**

OutlookAttachSnatcher is a VBA macro for **Classic Microsoft Outlook for Windows** that downloads attachments from multiple emails in one operation. Select messages manually, or search the current Outlook folder by sender name/email, subject and date range.

It includes progress tracking, duplicate protection, temporary-file safety, cancellation support, disk-space checking and large-batch-friendly processing.

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
- ⏳ Live download progress with attachment counter and percentage
- 💾 Calculates total attachment size before starting
- 🚨 Checks available local disk space before downloading
- 🧹 Cleans invalid Windows filename characters
- 📏 Protects against excessively long Windows paths
- 🗃️ Uses temporary `.downloading` files while saving
- ❌ Failed attachments do not stop the entire batch
- ⌨️ Press `ESC` to cancel at a safe point
- ⚡ Uses cooperative UI yielding to keep Outlook responsive
- 🧠 Stores lightweight message IDs instead of keeping large numbers of `MailItem` objects in memory
- 📦 Everything is contained in **one VBA module**
- 🔧 No external libraries or additional software required

---

## 🖥️ Requirements

- Windows
- **Classic Microsoft Outlook for Windows**
- VBA / Macros enabled

> [!IMPORTANT]
> OutlookAttachSnatcher does **not** work with **New Outlook for Windows** because New Outlook does not support traditional Outlook VBA macros. Switch to **Classic Outlook** before installing or running it.

---

## 📦 Installation

### Option A — Import the module (recommended)

1. Download `OutlookAttachSnatcher.bas` from this repository.
2. Open **Classic Outlook**.
3. Press `ALT + F11` to open the VBA Editor.
4. Choose **File → Import File...**.
5. Select `OutlookAttachSnatcher.bas`.
6. Press `CTRL + S` to save the Outlook VBA project.
7. Close the VBA Editor.

### Option B — Copy and paste

1. Open `OutlookAttachSnatcher.bas` on GitHub and copy its contents.
2. In Classic Outlook press `ALT + F11`.
3. Choose **Insert → Module**.
4. Paste the code into the new standard module.
5. Press `CTRL + S`.

---

## ▶️ Run OutlookAttachSnatcher

Press:

```text
ALT + F8
```

Select:

```text
OutlookAttachSnatcher
```

and click **Run**.

---

## 🔘 Add OutlookAttachSnatcher as an Outlook Button

For regular use, add the macro to the **Quick Access Toolbar** or the Outlook **Ribbon**.

### Option A — Quick Access Toolbar

In Classic Outlook:

1. Open **File → Options → Quick Access Toolbar**.
2. Set **Choose commands from:** to **Macros**.
3. Find `OutlookAttachSnatcher`.
4. Outlook may show a qualified name such as `Project1.Module1.OutlookAttachSnatcher`.
5. Select it and click **Add >>**.
6. Select the newly added macro on the right and click **Modify...**.
7. Choose an icon and set the display name to `OutlookAttachSnatcher`.
8. Click **OK**, then **OK** again.

You can now launch the downloader from the Quick Access Toolbar.

### Option B — Outlook Ribbon

1. Open **File → Options → Customize Ribbon**.
2. On the right, select the **Home** tab.
3. Click **New Group** and rename it `Attachments`.
4. On the left, set **Choose commands from:** to **Macros**.
5. Select `OutlookAttachSnatcher`.
6. Click **Add >>**.
7. Rename the displayed command to `OutlookAttachSnatcher` and choose an icon if desired.
8. Click **OK**.

You can then run it from:

```text
Home → Attachments → OutlookAttachSnatcher
```

---

## 📥 Download Modes

When the macro starts:

```text
Choose how you want to process emails:

YES = Process currently selected emails
NO = Search current Outlook folder
CANCEL = Exit
```

### Selected emails

Select multiple messages with `CTRL + Click` or `SHIFT + Click`, then run `OutlookAttachSnatcher`.

### Search current folder

Choose **NO** and optionally filter by:

- Sender name
- Sender email address
- Subject text
- Start date
- End date

The search operates on the **currently open Outlook folder** and does not recursively search subfolders.

---

## 🔎 Sender Search

The Sender field accepts a display name, email address or partial text.

```text
Name Surname
```

```text
namesurname@example.com
```

```text
surname
```

Matching is case-insensitive. The macro checks both the sender display name and email address. For Exchange accounts it also attempts to resolve the sender's primary SMTP address.

---

## 📝 Subject Search

Enter any text the subject must contain, for example:

```text
Daily Report
```

This can match:

```text
Daily Report 20260624
Daily Report Project A
Site Daily Report
DAILY REPORT 2026-06-24
```

Leave the field blank to ignore subject filtering.

---

## 📅 Date Filtering

The date range is optional.

```text
From Date: 01/06/2026
To Date:   30/06/2026
```

The complete final day is included. Leave either field blank if that limit is not required.

---

## 📂 Attachment Organization

The macro asks:

```text
How should attachments be organized?

YES = Save all attachments in one folder
NO = Create a separate folder for each email
CANCEL = Exit
```

### Single folder

```text
Attachments/
├── Drawing.pdf
├── Report.xlsx
├── Photo.jpg
├── Drawing_1.pdf
└── Report_1.xlsx
```

### Separate folder for each email

```text
Attachments/
├── 2026-06-24_09-15 - Name Surname - Daily Report/
│   ├── Daily_Report.pdf
│   └── Drawing.pdf
└── 2026-06-25_09-20 - Name Surname - Daily Report/
    ├── Daily_Report.pdf
    └── Site_Photo.jpg
```

Generated email folders use the format:

```text
YYYY-MM-DD_HH-MM - Sender - Subject
```

---

## 🔢 Duplicate Filename Protection

Existing files are not intentionally overwritten. If `Report.pdf` already exists, subsequent attachments are named:

```text
Report.pdf
Report_1.pdf
Report_2.pdf
Report_3.pdf
```

---

## ⏳ Progress Tracking

Before downloading, OutlookAttachSnatcher scans the matching messages and shows totals such as:

```text
Matched emails: 184
Total attachments: 300
Total size: 4.82 GB
```

During processing, a temporary Outlook progress toolbar shows the current count and percentage, for example:

```text
[##########--------------] 127 / 300 - 42%
```

It also displays the current attachment/status where supported by the Classic Outlook CommandBars UI.

---

## ⌨️ Cancel a Download

Press `ESC` during processing to request cancellation. The macro stops at a safe point.

> [!NOTE]
> `Attachment.SaveAsFile` is synchronous. If Outlook is currently writing one large attachment, cancellation takes effect after that individual save operation returns.

---

## 🛡️ Large File Safety

Attachments are first written using a temporary filename:

```text
Daily_Report.pdf.downloading
```

After Outlook successfully saves the attachment, it is renamed to:

```text
Daily_Report.pdf
```

If the save fails, the macro attempts to clean up the temporary file and continues with the remaining attachments.

---

## 💾 Disk Space Check

Before downloading, OutlookAttachSnatcher calculates the combined reported attachment size and checks local free disk space with a 10% safety margin.

For UNC/network paths, free-space reporting is not always reliable, so inability to determine network free space does not by itself block the job.

---

## 📏 Path and Filename Protection

The macro sanitizes characters that are invalid in Windows filenames and limits generated attachment filename lengths to reduce failures caused by long paths.

---

## 🚀 Large Mailboxes and Large Batches

During the initial scan, OutlookAttachSnatcher stores lightweight `EntryID` and `StoreID` references for matching messages. It retrieves each message again only when needed for processing instead of deliberately retaining hundreds or thousands of Outlook `MailItem` COM objects.

It also calls `DoEvents` periodically so Outlook and Windows can process UI events during long jobs.

---

## 🧵 Why Isn't It Multithreaded?

Classic Outlook VBA uses the Outlook Object Model through COM. Outlook objects such as `MailItem`, `Attachment` and `AddressEntry` should not be treated as safely parallelizable VBA objects.

OutlookAttachSnatcher therefore prioritizes reliability with:

- Sequential attachment processing
- Cooperative UI yielding
- Lightweight message references
- Per-file error handling
- Temporary-file protection
- Cancellation support
- Pre-scanning
- Disk-space validation

A truly parallel/server-side implementation would require a different architecture, such as Microsoft Graph or another appropriate API.

---

## 📊 Completion Summary

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

A cancelled job reports `Status: Cancelled by user`.

---

## ⚠️ Macro Security

Outlook may prevent VBA macros from running depending on Office security settings or organizational policy. Only enable and run macros from sources you trust, and inspect `OutlookAttachSnatcher.bas` before running it.

> [!WARNING]
> Do not globally weaken Office security settings simply to run a macro, especially on a managed work computer. If macros are disabled by your organization, contact your IT administrator.

---

## 🌐 Unicode / International Filenames

The macro does not intentionally convert attachment filenames, email subjects or sender names to ASCII/Greeklish. Windows-compatible Unicode filenames can remain in their original language where supported by Outlook/VBA and the filesystem.

The macro's own UI messages are written in English for broad compatibility across Office installations.

---

## 🧰 Troubleshooting

### The macro does not appear in `ALT + F8`

Make sure the main procedure is:

```vb
Public Sub OutlookAttachSnatcher()
```

and that the code is in a **standard Module**, not `ThisOutlookSession`.

### The button does nothing

Check that:

- You are using Classic Outlook
- VBA macros are permitted
- `OutlookAttachSnatcher` still exists in the VBA project
- The button is assigned to `OutlookAttachSnatcher`

### Search returns no emails

Check the current Outlook folder, sender spelling/email, subject filter and date range.

### Does it search subfolders?

No. The current version searches only the folder currently open in Outlook.

### A file failed to download

The macro counts the attachment as failed and continues with the remaining files where possible.

---

## 📸 Screenshots

Screenshots will be added under `screenshots/` as the UI documentation is completed.

Recommended repository layout:

```text
OutlookAttachSnatcher/
├── OutlookAttachSnatcher.bas
├── README.md
├── LICENSE
└── screenshots/
    ├── 01-outlook-button.png
    ├── 02-mode-selection.png
    ├── 03-search-filter.png
    ├── 04-pre-scan-summary.png
    ├── 05-progress.png
    └── 06-completed.png
```

Before publishing screenshots, remove or blur real names, email addresses, project names and other private/company information.

---

## 🔧 Compatibility

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
| Large-batch handling | ✅ |
| True multithreading | ❌ |
| Recursive subfolder search | ❌ |

---

## 🗺️ Roadmap

Potential future improvements:

- CSV download/error log
- Full destination-path-aware folder truncation
- Recursive subfolder searching
- Attachment extension filters
- Attachment size filters
- Saved search presets
- Improved progress UI
- Resume/retry failed downloads

---

## 🤝 Contributing

Bug reports, improvements, suggestions and pull requests are welcome.

---

## 📜 License

This project is licensed under the **GNU General Public License v3.0 (GPL-3.0)**. See `LICENSE` for details.

---

## ⭐ OutlookAttachSnatcher

If OutlookAttachSnatcher saves you from manually opening hundreds of emails and clicking **Save As** all day, consider giving the repository a ⭐.

**Happy attachment snatching. 📎**
