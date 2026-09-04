Option Explicit

' ============================================================
' OUTLOOK BULK ATTACHMENT DOWNLOADER
' Single-module version
'
' Features:
' - Selected emails OR search current Outlook folder
' - Sender name OR sender email
' - Subject contains
' - Optional date range
' - One folder OR separate folder per email
' - Pre-scan and attachment counting
' - Total download size
' - Live progress toolbar
' - ESC to cancel
' - Duplicate filename protection
' - Temporary .downloading files
' - Long filename/path protection
' - Disk space check
' - Large mailbox / many-record friendly
' - Individual file error handling
'
' Classic Outlook for Windows only.
' ============================================================


' ============================================================
' WINDOWS API
' Used only to detect ESC for cancellation.
' ============================================================

#If VBA7 Then

    Private Declare PtrSafe Function GetAsyncKeyState _
        Lib "user32" _
        (ByVal vKey As Long) As Integer

#Else

    Private Declare Function GetAsyncKeyState _
        Lib "user32" _
        (ByVal vKey As Long) As Integer

#End If


' ============================================================
' CONSTANTS
' ============================================================

Private Const MAX_PATH_LENGTH As Long = 235
Private Const UI_YIELD_INTERVAL As Long = 5
Private Const FREE_SPACE_MARGIN As Double = 1.1
Private Const PROGRESS_BAR_LENGTH As Long = 24
Private Const VK_ESCAPE As Long = 27
Private Const msoBarFloating As Long = 4
Private Const msoControlButton As Long = 1


' ============================================================
' GLOBAL PROGRESS TOOLBAR OBJECTS
' ============================================================

Private gProgressBar As Object
Private gProgressStatus As Object
Private gProgressCount As Object
Private gProgressCancel As Object
Private gCancelled As Boolean


' ============================================================
' MAIN PROCEDURE
' ============================================================

Public Sub OutlookAttachSnatcher()

    On Error GoTo FatalError

    Dim mode As VbMsgBoxResult
    Dim saveMode As VbMsgBoxResult
    Dim senderFilter As String
    Dim subjectFilter As String
    Dim fromDateText As String
    Dim toDateText As String
    Dim fromDate As Date
    Dim toDate As Date
    Dim useFromDate As Boolean
    Dim useToDate As Boolean
    Dim baseFolder As String
    Dim shellApp As Object
    Dim folderObj As Object
    Dim ns As Outlook.NameSpace
    Dim objSelection As Outlook.Selection
    Dim objFolder As Outlook.Folder
    Dim objItems As Outlook.Items
    Dim objItem As Object
    Dim mail As Outlook.MailItem
    Dim att As Outlook.Attachment
    Dim matchedEmails As Long
    Dim totalAttachments As Long
    Dim totalBytes As Double
    Dim processedFiles As Long
    Dim savedFiles As Long
    Dim failedFiles As Long
    Dim scanCounter As Long
    Dim continueAnswer As VbMsgBoxResult
    Dim entryIDs As Collection
    Dim storeIDs As Collection
    Dim entryID As String
    Dim storeID As String
    Dim i As Long
    Dim resultText As String

    gCancelled = False
    Set ns = Application.Session
    Set entryIDs = New Collection
    Set storeIDs = New Collection

    mode = MsgBox( _
        "Choose how you want to process emails:" & vbCrLf & vbCrLf & _
        "YES = Process currently selected emails" & vbCrLf & _
        "NO = Search current Outlook folder" & vbCrLf & _
        "CANCEL = Exit", _
        vbYesNoCancel + vbQuestion, _
        "OutlookAttachSnatcher")

    If mode = vbCancel Then Exit Sub

    If mode = vbNo Then

        senderFilter = InputBox( _
            "Enter sender name OR sender email." & vbCrLf & vbCrLf & _
            "Examples:" & vbCrLf & _
            "Name Surname" & vbCrLf & _
            "namesurname@example.com" & vbCrLf & _
            "surname" & vbCrLf & vbCrLf & _
            "Leave blank to ignore sender filtering.", _
            "Sender Filter")

        subjectFilter = InputBox( _
            "Enter text that the email subject must contain." & vbCrLf & vbCrLf & _
            "Example:" & vbCrLf & _
            "Daily Report" & vbCrLf & vbCrLf & _
            "Leave blank to ignore subject filtering.", _
            "Subject Filter")

        fromDateText = InputBox( _
            "Optional start date." & vbCrLf & vbCrLf & _
            "Example:" & vbCrLf & _
            "01/06/2026" & vbCrLf & vbCrLf & _
            "Leave blank for no start date.", _
            "From Date")

        If Trim(fromDateText) <> "" Then
            If Not IsDate(fromDateText) Then
                MsgBox "Invalid start date.", vbExclamation, "Error"
                Exit Sub
            End If
            fromDate = DateValue(CDate(fromDateText))
            useFromDate = True
        End If

        toDateText = InputBox( _
            "Optional end date." & vbCrLf & vbCrLf & _
            "Example:" & vbCrLf & _
            "30/06/2026" & vbCrLf & vbCrLf & _
            "Leave blank for no end date.", _
            "To Date")

        If Trim(toDateText) <> "" Then
            If Not IsDate(toDateText) Then
                MsgBox "Invalid end date.", vbExclamation, "Error"
                Exit Sub
            End If
            toDate = DateValue(CDate(toDateText)) + 1
            useToDate = True
        End If

        If useFromDate And useToDate Then
            If fromDate >= toDate Then
                MsgBox _
                    "The start date must be before or equal to the end date.", _
                    vbExclamation, _
                    "Invalid Date Range"
                Exit Sub
            End If
        End If
    End If

    saveMode = MsgBox( _
        "How should attachments be organized?" & vbCrLf & vbCrLf & _
        "YES = Save all attachments in one folder" & vbCrLf & _
        "NO = Create a separate folder for each email" & vbCrLf & _
        "CANCEL = Exit", _
        vbYesNoCancel + vbQuestion, _
        "Save Mode")

    If saveMode = vbCancel Then Exit Sub

    Set shellApp = CreateObject("Shell.Application")
    Set folderObj = shellApp.BrowseForFolder(0, "Select the destination folder:", 0)
    If folderObj Is Nothing Then Exit Sub

    baseFolder = folderObj.Self.Path
    If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"

    If mode = vbYes Then

        Set objSelection = Application.ActiveExplorer.Selection

        If objSelection.Count = 0 Then
            MsgBox "No emails are selected.", vbExclamation, "No Selection"
            Exit Sub
        End If

        For Each objItem In objSelection
            If TypeOf objItem Is Outlook.MailItem Then
                Set mail = objItem
                entryID = ""
                storeID = ""

                On Error Resume Next
                entryID = mail.EntryID
                storeID = mail.Parent.StoreID
                On Error GoTo FatalError

                If entryID <> "" Then
                    entryIDs.Add entryID
                    storeIDs.Add storeID
                    matchedEmails = matchedEmails + 1

                    For Each att In mail.Attachments
                        totalAttachments = totalAttachments + 1
                        totalBytes = totalBytes + CDbl(att.Size)
                    Next att
                End If
            End If

            scanCounter = scanCounter + 1
            If scanCounter Mod 25 = 0 Then DoEvents
        Next objItem

    Else

        Set objFolder = Application.ActiveExplorer.CurrentFolder
        Set objItems = objFolder.Items

        For Each objItem In objItems
            If TypeOf objItem Is Outlook.MailItem Then
                Set mail = objItem

                If MailMatchesFilters( _
                    mail, senderFilter, subjectFilter, fromDate, toDate, _
                    useFromDate, useToDate) Then

                    entryID = ""
                    storeID = ""

                    On Error Resume Next
                    entryID = mail.EntryID
                    storeID = mail.Parent.StoreID
                    On Error GoTo FatalError

                    If entryID <> "" Then
                        entryIDs.Add entryID
                        storeIDs.Add storeID
                        matchedEmails = matchedEmails + 1

                        For Each att In mail.Attachments
                            totalAttachments = totalAttachments + 1
                            totalBytes = totalBytes + CDbl(att.Size)
                        Next att
                    End If
                End If
            End If

            scanCounter = scanCounter + 1
            If scanCounter Mod 25 = 0 Then DoEvents
        Next objItem
    End If

    If matchedEmails = 0 Then
        MsgBox "No matching emails were found.", vbInformation, "No Results"
        Exit Sub
    End If

    If totalAttachments = 0 Then
        MsgBox _
            Format(matchedEmails, "#,##0") & " matching email(s) were found," & vbCrLf & _
            "but none of them contain attachments.", _
            vbInformation, _
            "No Attachments"
        Exit Sub
    End If

    If Not HasEnoughDiskSpace(baseFolder, totalBytes) Then
        MsgBox _
            "There may not be enough free disk space." & vbCrLf & vbCrLf & _
            "Attachment size: " & FormatBytes(totalBytes) & vbCrLf & _
            "Safety margin: 10%" & vbCrLf & vbCrLf & _
            "Please select another drive or free some disk space.", _
            vbCritical, _
            "Insufficient Disk Space"
        Exit Sub
    End If

    continueAnswer = MsgBox( _
        "Ready to download attachments." & vbCrLf & vbCrLf & _
        "Matched emails: " & Format(matchedEmails, "#,##0") & vbCrLf & _
        "Total attachments: " & Format(totalAttachments, "#,##0") & vbCrLf & _
        "Total size: " & FormatBytes(totalBytes) & vbCrLf & vbCrLf & _
        "Press ESC during processing to cancel." & vbCrLf & vbCrLf & _
        "Continue?", _
        vbYesNo + vbQuestion, _
        "Ready")

    If continueAnswer <> vbYes Then Exit Sub

    CreateProgressToolbar
    UpdateProgressToolbar 0, totalAttachments, "Starting download..."
    Call GetAsyncKeyState(VK_ESCAPE)

    For i = 1 To entryIDs.Count

        If EscapePressed() Then
            gCancelled = True
            Exit For
        End If

        Set mail = Nothing

        On Error Resume Next
        Set mail = ns.GetItemFromID(CStr(entryIDs(i)), CStr(storeIDs(i)))
        On Error GoTo FatalError

        If Not mail Is Nothing Then
            SaveMailAttachments _
                mail, baseFolder, saveMode, processedFiles, savedFiles, failedFiles, totalAttachments
        End If

        If gCancelled Then Exit For
        If i Mod 10 = 0 Then DoEvents
    Next i

    If gCancelled Then
        UpdateProgressToolbar processedFiles, totalAttachments, "Cancelled by user"
    Else
        UpdateProgressToolbar totalAttachments, totalAttachments, "Completed"
    End If

    DoEvents

    resultText = _
        "Operation finished." & vbCrLf & vbCrLf & _
        "Matched emails: " & Format(matchedEmails, "#,##0") & vbCrLf & _
        "Total attachments: " & Format(totalAttachments, "#,##0") & vbCrLf & _
        "Total size: " & FormatBytes(totalBytes) & vbCrLf & _
        "Downloaded: " & Format(savedFiles, "#,##0") & vbCrLf & _
        "Failed: " & Format(failedFiles, "#,##0") & vbCrLf & _
        "Processed: " & Format(processedFiles, "#,##0") & _
        " / " & Format(totalAttachments, "#,##0")

    If gCancelled Then
        resultText = resultText & vbCrLf & vbCrLf & "Status: Cancelled by user"
    Else
        resultText = resultText & vbCrLf & vbCrLf & "Status: Completed"
    End If

    DestroyProgressToolbar

    MsgBox _
        resultText & vbCrLf & vbCrLf & _
        "Destination:" & vbCrLf & baseFolder, _
        vbInformation, _
        "OutlookAttachSnatcher"

    Exit Sub

FatalError:

    DestroyProgressToolbar

    MsgBox _
        "The operation stopped because of an unexpected error." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & vbCrLf & Err.Description, _
        vbCritical, _
        "OutlookAttachSnatcher"

End Sub


Private Function MailMatchesFilters( _
    ByVal mail As Outlook.MailItem, _
    ByVal senderFilter As String, _
    ByVal subjectFilter As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    ByVal useFromDate As Boolean, _
    ByVal useToDate As Boolean) As Boolean

    Dim senderName As String
    Dim senderEmail As String
    Dim senderMatched As Boolean
    Dim subjectMatched As Boolean

    senderName = ""
    senderEmail = ""

    On Error Resume Next
    senderName = mail.SenderName
    senderEmail = GetSenderEmail(mail)
    On Error GoTo 0

    If Trim(senderFilter) = "" Then
        senderMatched = True
    Else
        senderMatched = _
            InStr(1, senderName, senderFilter, vbTextCompare) > 0 Or _
            InStr(1, senderEmail, senderFilter, vbTextCompare) > 0
    End If

    If Trim(subjectFilter) = "" Then
        subjectMatched = True
    Else
        subjectMatched = InStr(1, mail.Subject, subjectFilter, vbTextCompare) > 0
    End If

    If Not senderMatched Then Exit Function
    If Not subjectMatched Then Exit Function

    If useFromDate Then
        If mail.ReceivedTime < fromDate Then Exit Function
    End If

    If useToDate Then
        If mail.ReceivedTime >= toDate Then Exit Function
    End If

    MailMatchesFilters = True

End Function


Private Sub SaveMailAttachments( _
    ByVal mail As Outlook.MailItem, _
    ByVal baseFolder As String, _
    ByVal saveMode As VbMsgBoxResult, _
    ByRef processedFiles As Long, _
    ByRef savedFiles As Long, _
    ByRef failedFiles As Long, _
    ByVal totalAttachments As Long)

    Dim objAttachment As Outlook.Attachment
    Dim targetFolder As String
    Dim emailFolderName As String
    Dim fileName As String
    Dim finalPath As String
    Dim tempPath As String
    Dim counter As Long
    Dim saveSucceeded As Boolean

    If mail.Attachments.Count = 0 Then Exit Sub

    If saveMode = vbNo Then

        emailFolderName = _
            Format(mail.ReceivedTime, "yyyy-mm-dd_hh-nn") & " - " & _
            CleanFileName(GetSenderDisplayName(mail)) & " - " & _
            CleanFileName(mail.Subject)

        If Len(emailFolderName) > 120 Then emailFolderName = Left(emailFolderName, 120)

        targetFolder = GetUniqueFolderPath(baseFolder, emailFolderName)

        If Not FolderExists(targetFolder) Then
            On Error Resume Next
            MkDir targetFolder

            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0

                failedFiles = failedFiles + mail.Attachments.Count
                processedFiles = processedFiles + mail.Attachments.Count

                UpdateProgressToolbar _
                    processedFiles, totalAttachments, "Could not create destination folder"
                Exit Sub
            End If

            On Error GoTo 0
        End If

        targetFolder = targetFolder & "\"
    Else
        targetFolder = baseFolder
    End If

    For Each objAttachment In mail.Attachments

        If EscapePressed() Then
            gCancelled = True
            Exit Sub
        End If

        fileName = CleanFileName(objAttachment.FileName)
        fileName = LimitFileNameLength(targetFolder, fileName, MAX_PATH_LENGTH)
        finalPath = targetFolder & fileName
        counter = 1

        Do While FileExists(finalPath)
            finalPath = _
                targetFolder & GetFileNameWithoutExtension(fileName) & _
                "_" & counter & GetFileExtension(fileName)
            counter = counter + 1
        Loop

        tempPath = GetTemporaryAttachmentPath(finalPath)

        UpdateProgressToolbar _
            processedFiles, totalAttachments, _
            "Downloading: " & ShortDisplayName(fileName, 55) & _
            " (" & FormatBytes(CDbl(objAttachment.Size)) & ")"

        saveSucceeded = False

        On Error Resume Next
        Err.Clear
        objAttachment.SaveAsFile tempPath

        If Err.Number = 0 Then
            If FileExists(tempPath) Then
                Err.Clear
                Name tempPath As finalPath
                If Err.Number = 0 Then saveSucceeded = True
            End If
        End If

        If saveSucceeded Then
            savedFiles = savedFiles + 1
        Else
            failedFiles = failedFiles + 1
            DeleteFileSafe tempPath
        End If

        Err.Clear
        On Error GoTo 0

        processedFiles = processedFiles + 1

        UpdateProgressToolbar _
            processedFiles, totalAttachments, _
            "Downloaded: " & Format(savedFiles, "#,##0") & _
            "   Failed: " & Format(failedFiles, "#,##0")

        If processedFiles Mod UI_YIELD_INTERVAL = 0 Then DoEvents

        If EscapePressed() Then
            gCancelled = True
            Exit Sub
        End If

    Next objAttachment

End Sub


Private Sub CreateProgressToolbar()

    On Error Resume Next

    DestroyProgressToolbar

    Set gProgressBar = Application.ActiveExplorer.CommandBars.Add( _
        Name:="AttachmentDownloadProgress", _
        Position:=msoBarFloating, _
        Temporary:=True)

    If gProgressBar Is Nothing Then Exit Sub

    Set gProgressStatus = gProgressBar.Controls.Add(Type:=msoControlButton)
    Set gProgressCount = gProgressBar.Controls.Add(Type:=msoControlButton)
    Set gProgressCancel = gProgressBar.Controls.Add(Type:=msoControlButton)

    gProgressStatus.Caption = "Preparing..."
    gProgressCount.Caption = "[------------------------] 0 / 0 - 0%"
    gProgressCancel.Caption = "Press ESC to cancel"

    gProgressStatus.Enabled = False
    gProgressCount.Enabled = False
    gProgressCancel.Enabled = False

    gProgressBar.Visible = True

    On Error GoTo 0

End Sub


Private Sub UpdateProgressToolbar( _
    ByVal currentValue As Long, _
    ByVal totalValue As Long, _
    ByVal statusText As String)

    Dim pct As Double
    Dim filled As Long
    Dim emptyCount As Long
    Dim barText As String

    If totalValue <= 0 Then
        pct = 0
    Else
        pct = currentValue / totalValue
    End If

    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1

    filled = CLng(pct * PROGRESS_BAR_LENGTH)
    If filled > PROGRESS_BAR_LENGTH Then filled = PROGRESS_BAR_LENGTH
    emptyCount = PROGRESS_BAR_LENGTH - filled

    barText = "[" & String(filled, "#") & String(emptyCount, "-") & "]"

    On Error Resume Next

    If Not gProgressStatus Is Nothing Then gProgressStatus.Caption = statusText

    If Not gProgressCount Is Nothing Then
        gProgressCount.Caption = _
            barText & " " & Format(currentValue, "#,##0") & _
            " / " & Format(totalValue, "#,##0") & _
            " - " & Format(pct, "0%")
    End If

    If Not gProgressBar Is Nothing Then gProgressBar.Visible = True

    On Error GoTo 0
    DoEvents

End Sub


Private Sub DestroyProgressToolbar()

    On Error Resume Next

    If Not gProgressBar Is Nothing Then gProgressBar.Delete

    Set gProgressCancel = Nothing
    Set gProgressCount = Nothing
    Set gProgressStatus = Nothing
    Set gProgressBar = Nothing

    On Error GoTo 0

End Sub


Private Function EscapePressed() As Boolean

    DoEvents
    EscapePressed = ((GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0)

End Function


Private Function GetSenderEmail(ByVal mail As Outlook.MailItem) As String

    Dim exchUser As Outlook.ExchangeUser
    Dim senderObj As Outlook.AddressEntry

    On Error Resume Next

    If mail.SenderEmailType = "EX" Then

        Set senderObj = mail.Sender

        If Not senderObj Is Nothing Then
            Set exchUser = senderObj.GetExchangeUser

            If Not exchUser Is Nothing Then
                GetSenderEmail = exchUser.PrimarySmtpAddress
            Else
                GetSenderEmail = mail.SenderEmailAddress
            End If
        Else
            GetSenderEmail = mail.SenderEmailAddress
        End If

    Else
        GetSenderEmail = mail.SenderEmailAddress
    End If

    On Error GoTo 0

End Function


Private Function GetSenderDisplayName(ByVal mail As Outlook.MailItem) As String

    On Error Resume Next

    If Trim(mail.SenderName) <> "" Then
        GetSenderDisplayName = mail.SenderName
    Else
        GetSenderDisplayName = GetSenderEmail(mail)
    End If

    If Trim(GetSenderDisplayName) = "" Then GetSenderDisplayName = "Unknown Sender"

    On Error GoTo 0

End Function


Private Function CleanFileName(ByVal fileName As String) As String

    Dim invalidChars As Variant
    Dim ch As Variant

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each ch In invalidChars
        fileName = Replace(fileName, ch, "_")
    Next ch

    fileName = Trim(fileName)

    Do While Len(fileName) > 0
        If Right(fileName, 1) = "." Or Right(fileName, 1) = " " Then
            fileName = Left(fileName, Len(fileName) - 1)
        Else
            Exit Do
        End If
    Loop

    If fileName = "" Then fileName = "Unnamed"
    CleanFileName = fileName

End Function


Private Function GetFileNameWithoutExtension(ByVal fileName As String) As String

    Dim p As Long
    p = InStrRev(fileName, ".")

    If p > 1 Then
        GetFileNameWithoutExtension = Left(fileName, p - 1)
    Else
        GetFileNameWithoutExtension = fileName
    End If

End Function


Private Function GetFileExtension(ByVal fileName As String) As String

    Dim p As Long
    p = InStrRev(fileName, ".")

    If p > 1 Then
        GetFileExtension = Mid(fileName, p)
    Else
        GetFileExtension = ""
    End If

End Function


Private Function FileExists(ByVal filePath As String) As Boolean

    On Error Resume Next
    FileExists = _
        (Len(Dir(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly Or vbArchive)) > 0)
    On Error GoTo 0

End Function


Private Function FolderExists(ByVal folderPath As String) As Boolean

    Dim attr As Long

    On Error Resume Next
    Err.Clear
    attr = GetAttr(folderPath)

    If Err.Number = 0 Then
        FolderExists = ((attr And vbDirectory) = vbDirectory)
    Else
        FolderExists = False
    End If

    Err.Clear
    On Error GoTo 0

End Function


Private Function GetUniqueFolderPath( _
    ByVal baseFolder As String, _
    ByVal folderName As String) As String

    Dim folderPath As String
    Dim counter As Long

    folderPath = baseFolder & folderName
    counter = 1

    Do While FolderExists(folderPath)
        folderPath = baseFolder & folderName & "_" & counter
        counter = counter + 1
    Loop

    GetUniqueFolderPath = folderPath

End Function


Private Function GetTemporaryAttachmentPath(ByVal finalPath As String) As String

    Dim tempPath As String
    Dim counter As Long

    tempPath = finalPath & ".downloading"
    counter = 1

    Do While FileExists(tempPath)
        tempPath = finalPath & ".downloading_" & counter
        counter = counter + 1
    Loop

    GetTemporaryAttachmentPath = tempPath

End Function


Private Sub DeleteFileSafe(ByVal filePath As String)

    On Error Resume Next
    If FileExists(filePath) Then Kill filePath
    Err.Clear
    On Error GoTo 0

End Sub


Private Function LimitFileNameLength( _
    ByVal folderPath As String, _
    ByVal fileName As String, _
    ByVal maxFullPath As Long) As String

    Dim extension As String
    Dim baseName As String
    Dim availableLength As Long

    extension = GetFileExtension(fileName)
    baseName = GetFileNameWithoutExtension(fileName)

    availableLength = maxFullPath - Len(folderPath) - Len(extension)
    If availableLength < 10 Then availableLength = 10

    If Len(baseName) > availableLength Then baseName = Left(baseName, availableLength)

    LimitFileNameLength = baseName & extension

End Function


Private Function ShortDisplayName(ByVal text As String, ByVal maxLength As Long) As String

    If Len(text) <= maxLength Then
        ShortDisplayName = text
    Else
        ShortDisplayName = Left(text, maxLength - 3) & "..."
    End If

End Function


Private Function FormatBytes(ByVal bytes As Double) As String

    If bytes >= 1099511627776# Then
        FormatBytes = Format(bytes / 1099511627776#, "0.00") & " TB"
    ElseIf bytes >= 1073741824# Then
        FormatBytes = Format(bytes / 1073741824#, "0.00") & " GB"
    ElseIf bytes >= 1048576# Then
        FormatBytes = Format(bytes / 1048576#, "0.00") & " MB"
    ElseIf bytes >= 1024# Then
        FormatBytes = Format(bytes / 1024#, "0.0") & " KB"
    Else
        FormatBytes = Format(bytes, "0") & " bytes"
    End If

End Function


Private Function HasEnoughDiskSpace( _
    ByVal folderPath As String, _
    ByVal requiredBytes As Double) As Boolean

    Dim fso As Object
    Dim drv As Object
    Dim driveLetter As String

    On Error GoTo UnableToCheck

    If Left(folderPath, 2) = "\\" Then
        HasEnoughDiskSpace = True
        Exit Function
    End If

    driveLetter = Left(folderPath, 2)

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set drv = fso.GetDrive(driveLetter)

    HasEnoughDiskSpace = _
        CDbl(drv.FreeSpace) >= (requiredBytes * FREE_SPACE_MARGIN)

    Exit Function

UnableToCheck:
    HasEnoughDiskSpace = True

End Function
