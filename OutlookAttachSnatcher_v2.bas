Option Explicit

' ============================================================
' OutlookAttachSnatcher v2 - EXPERIMENTAL
' Classic Microsoft Outlook for Windows / VBA
'
' Stable version remains: OutlookAttachSnatcher.bas
'
' v2 additions:
' - Optional recursive subfolder search
' - Attachment extension filter
' - Optional attachment min/max size filters
' - Unicode CSV audit log (UTF-16LE, Excel-friendly)
' - Full-path-aware email folder truncation
' - Same safe temp-file, duplicate, disk-space and ESC handling
'
' IMPORTANT:
' Outlook VBA / Outlook Object Model is intentionally processed
' sequentially. Do not try to parallelize MailItem/Attachment COM
' objects from VBA.
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState _
        Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState _
        Lib "user32" (ByVal vKey As Long) As Integer
#End If

Private Const MAX_PATH_LENGTH As Long = 235
Private Const EMAIL_FOLDER_PATH_LIMIT As Long = 180
Private Const TEMP_SUFFIX_RESERVE As Long = 24
Private Const UI_YIELD_INTERVAL As Long = 5
Private Const FREE_SPACE_MARGIN As Double = 1.1
Private Const PROGRESS_BAR_LENGTH As Long = 24
Private Const VK_ESCAPE As Long = 27
Private Const msoBarFloating As Long = 4
Private Const msoControlButton As Long = 1

Private gProgressBar As Object
Private gProgressStatus As Object
Private gProgressCount As Object
Private gProgressCancel As Object
Private gCancelled As Boolean


Public Sub OutlookAttachSnatcherV2()

    On Error GoTo FatalError

    Dim mode As VbMsgBoxResult
    Dim saveMode As VbMsgBoxResult
    Dim recursiveAnswer As VbMsgBoxResult
    Dim senderFilter As String
    Dim subjectFilter As String
    Dim extensionFilter As String
    Dim minSizeText As String
    Dim maxSizeText As String
    Dim minSizeBytes As Double
    Dim maxSizeBytes As Double
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
    Dim objItem As Object
    Dim mail As Outlook.MailItem
    Dim matchedEmails As Long
    Dim totalAttachments As Long
    Dim totalBytes As Double
    Dim processedFiles As Long
    Dim savedFiles As Long
    Dim failedFiles As Long
    Dim continueAnswer As VbMsgBoxResult
    Dim entryIDs As Collection
    Dim storeIDs As Collection
    Dim i As Long
    Dim resultText As String
    Dim logPath As String
    Dim logFSO As Object
    Dim logStream As Object

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
        "OutlookAttachSnatcher v2")

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
            "Example: Daily Report" & vbCrLf & vbCrLf & _
            "Leave blank to ignore subject filtering.", _
            "Subject Filter")

        fromDateText = InputBox( _
            "Optional start date." & vbCrLf & vbCrLf & _
            "Example: 01/06/2026" & vbCrLf & vbCrLf & _
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
            "Example: 30/06/2026" & vbCrLf & vbCrLf & _
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

        recursiveAnswer = MsgBox( _
            "Search subfolders too?" & vbCrLf & vbCrLf & _
            "YES = Current folder + all subfolders" & vbCrLf & _
            "NO = Current folder only", _
            vbYesNo + vbQuestion, _
            "Recursive Search")

    End If

    extensionFilter = NormalizeExtensionFilter(InputBox( _
        "Optional attachment extension filter." & vbCrLf & vbCrLf & _
        "Examples:" & vbCrLf & _
        "pdf" & vbCrLf & _
        "pdf,xlsx,jpg" & vbCrLf & _
        ".pdf; .xlsx; .docx" & vbCrLf & vbCrLf & _
        "Leave blank to include all attachment types.", _
        "Attachment Extensions"))

    minSizeText = InputBox( _
        "Optional minimum attachment size in MB." & vbCrLf & vbCrLf & _
        "Example: 1.5" & vbCrLf & _
        "Leave blank for no minimum.", _
        "Minimum Attachment Size")

    If Trim(minSizeText) <> "" Then
        If Not IsNumeric(minSizeText) Or CDbl(minSizeText) < 0 Then
            MsgBox "Invalid minimum attachment size.", vbExclamation, "Error"
            Exit Sub
        End If
        minSizeBytes = CDbl(minSizeText) * 1048576#
    End If

    maxSizeText = InputBox( _
        "Optional maximum attachment size in MB." & vbCrLf & vbCrLf & _
        "Example: 100" & vbCrLf & _
        "Leave blank for no maximum.", _
        "Maximum Attachment Size")

    If Trim(maxSizeText) <> "" Then
        If Not IsNumeric(maxSizeText) Or CDbl(maxSizeText) < 0 Then
            MsgBox "Invalid maximum attachment size.", vbExclamation, "Error"
            Exit Sub
        End If
        maxSizeBytes = CDbl(maxSizeText) * 1048576#
    End If

    If minSizeBytes > 0 And maxSizeBytes > 0 Then
        If minSizeBytes > maxSizeBytes Then
            MsgBox "Minimum size cannot be greater than maximum size.", vbExclamation, "Error"
            Exit Sub
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

    ' --------------------------------------------------------
    ' PRE-SCAN
    ' --------------------------------------------------------

    If mode = vbYes Then

        Set objSelection = Application.ActiveExplorer.Selection

        If objSelection.Count = 0 Then
            MsgBox "No emails are selected.", vbExclamation, "No Selection"
            Exit Sub
        End If

        For Each objItem In objSelection
            If TypeOf objItem Is Outlook.MailItem Then
                Set mail = objItem
                AddMailIfEligible _
                    mail, entryIDs, storeIDs, matchedEmails, _
                    totalAttachments, totalBytes, _
                    extensionFilter, minSizeBytes, maxSizeBytes
            End If
            If (matchedEmails Mod 25) = 0 Then DoEvents
        Next objItem

    Else

        Set objFolder = Application.ActiveExplorer.CurrentFolder

        ScanFolderForMatches _
            objFolder, (recursiveAnswer = vbYes), _
            senderFilter, subjectFilter, fromDate, toDate, _
            useFromDate, useToDate, _
            extensionFilter, minSizeBytes, maxSizeBytes, _
            entryIDs, storeIDs, matchedEmails, totalAttachments, totalBytes

    End If

    If matchedEmails = 0 Then
        MsgBox "No matching emails were found.", vbInformation, "No Results"
        Exit Sub
    End If

    If totalAttachments = 0 Then
        MsgBox _
            Format(matchedEmails, "#,##0") & " matching email(s) were found," & vbCrLf & _
            "but no attachments matched the attachment filters.", _
            vbInformation, _
            "No Matching Attachments"
        Exit Sub
    End If

    If Not HasEnoughDiskSpaceV2(baseFolder, totalBytes) Then
        MsgBox _
            "There may not be enough free disk space." & vbCrLf & vbCrLf & _
            "Matching attachment size: " & FormatBytesV2(totalBytes) & vbCrLf & _
            "Safety margin: 10%", _
            vbCritical, _
            "Insufficient Disk Space"
        Exit Sub
    End If

    continueAnswer = MsgBox( _
        "Ready to download attachments." & vbCrLf & vbCrLf & _
        "Matched emails: " & Format(matchedEmails, "#,##0") & vbCrLf & _
        "Matching attachments: " & Format(totalAttachments, "#,##0") & vbCrLf & _
        "Total size: " & FormatBytesV2(totalBytes) & vbCrLf & _
        "Extensions: " & DisplayExtensionFilter(extensionFilter) & vbCrLf & _
        "Recursive search: " & IIf(mode = vbNo And recursiveAnswer = vbYes, "Yes", "No") & vbCrLf & vbCrLf & _
        "A CSV log will be created in the destination folder." & vbCrLf & _
        "Press ESC during processing to cancel." & vbCrLf & vbCrLf & _
        "Continue?", _
        vbYesNo + vbQuestion, _
        "OutlookAttachSnatcher v2")

    If continueAnswer <> vbYes Then Exit Sub

    ' --------------------------------------------------------
    ' OPEN UNICODE CSV LOG
    ' --------------------------------------------------------

    logPath = baseFolder & _
        "OutlookAttachSnatcher_v2_Log_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"

    Set logFSO = CreateObject("Scripting.FileSystemObject")
    Set logStream = logFSO.CreateTextFile(logPath, True, True)

    logStream.WriteLine _
        "Timestamp,Status,Sender,SenderEmail,Received,Subject,Attachment,SizeBytes,SavedPath,Error"

    CreateProgressToolbarV2
    UpdateProgressToolbarV2 0, totalAttachments, "Starting download..."
    Call GetAsyncKeyState(VK_ESCAPE)

    For i = 1 To entryIDs.Count

        If EscapePressedV2() Then
            gCancelled = True
            Exit For
        End If

        Set mail = Nothing

        On Error Resume Next
        Set mail = ns.GetItemFromID(CStr(entryIDs(i)), CStr(storeIDs(i)))
        On Error GoTo FatalError

        If Not mail Is Nothing Then
            SaveMailAttachmentsV2 _
                mail, baseFolder, saveMode, _
                extensionFilter, minSizeBytes, maxSizeBytes, _
                processedFiles, savedFiles, failedFiles, totalAttachments, _
                logStream
        End If

        If gCancelled Then Exit For
        If i Mod 10 = 0 Then DoEvents

    Next i

    If gCancelled Then
        UpdateProgressToolbarV2 processedFiles, totalAttachments, "Cancelled by user"
    Else
        UpdateProgressToolbarV2 totalAttachments, totalAttachments, "Completed"
    End If

    If Not logStream Is Nothing Then
        logStream.WriteLine CsvLine(Array( _
            Format(Now, "yyyy-mm-dd hh:nn:ss"), _
            IIf(gCancelled, "CANCELLED", "SUMMARY"), _
            "", "", "", "", "", "", "", _
            "MatchedEmails=" & matchedEmails & _
            "; TotalAttachments=" & totalAttachments & _
            "; Downloaded=" & savedFiles & _
            "; Failed=" & failedFiles & _
            "; Processed=" & processedFiles))
        logStream.Close
    End If

    DestroyProgressToolbarV2

    resultText = _
        "Operation finished." & vbCrLf & vbCrLf & _
        "Matched emails: " & Format(matchedEmails, "#,##0") & vbCrLf & _
        "Matching attachments: " & Format(totalAttachments, "#,##0") & vbCrLf & _
        "Total size: " & FormatBytesV2(totalBytes) & vbCrLf & _
        "Downloaded: " & Format(savedFiles, "#,##0") & vbCrLf & _
        "Failed: " & Format(failedFiles, "#,##0") & vbCrLf & _
        "Processed: " & Format(processedFiles, "#,##0") & _
        " / " & Format(totalAttachments, "#,##0") & vbCrLf & vbCrLf & _
        "Status: " & IIf(gCancelled, "Cancelled by user", "Completed") & vbCrLf & vbCrLf & _
        "CSV log:" & vbCrLf & logPath

    MsgBox resultText, vbInformation, "OutlookAttachSnatcher v2"
    Exit Sub

FatalError:

    On Error Resume Next
    If Not logStream Is Nothing Then
        logStream.WriteLine CsvLine(Array( _
            Format(Now, "yyyy-mm-dd hh:nn:ss"), _
            "FATAL", "", "", "", "", "", "", "", _
            "Error " & Err.Number & ": " & Err.Description))
        logStream.Close
    End If
    DestroyProgressToolbarV2
    On Error GoTo 0

    MsgBox _
        "The operation stopped because of an unexpected error." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & vbCrLf & Err.Description, _
        vbCritical, _
        "OutlookAttachSnatcher v2"

End Sub


Private Sub ScanFolderForMatches( _
    ByVal folder As Outlook.Folder, _
    ByVal recursive As Boolean, _
    ByVal senderFilter As String, _
    ByVal subjectFilter As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    ByVal useFromDate As Boolean, _
    ByVal useToDate As Boolean, _
    ByVal extensionFilter As String, _
    ByVal minSizeBytes As Double, _
    ByVal maxSizeBytes As Double, _
    ByRef entryIDs As Collection, _
    ByRef storeIDs As Collection, _
    ByRef matchedEmails As Long, _
    ByRef totalAttachments As Long, _
    ByRef totalBytes As Double)

    Dim items As Outlook.Items
    Dim objItem As Object
    Dim mail As Outlook.MailItem
    Dim subFolder As Outlook.Folder
    Dim counter As Long

    On Error GoTo FolderDone

    Set items = folder.Items

    For Each objItem In items

        If EscapePressedV2() Then
            gCancelled = True
            Exit Sub
        End If

        If TypeOf objItem Is Outlook.MailItem Then
            Set mail = objItem

            If MailMatchesFiltersV2( _
                mail, senderFilter, subjectFilter, fromDate, toDate, _
                useFromDate, useToDate) Then

                AddMailIfEligible _
                    mail, entryIDs, storeIDs, matchedEmails, _
                    totalAttachments, totalBytes, _
                    extensionFilter, minSizeBytes, maxSizeBytes
            End If
        End If

        counter = counter + 1
        If counter Mod 25 = 0 Then DoEvents

    Next objItem

    If recursive Then
        For Each subFolder In folder.Folders
            If gCancelled Then Exit Sub
            ScanFolderForMatches _
                subFolder, True, _
                senderFilter, subjectFilter, fromDate, toDate, _
                useFromDate, useToDate, _
                extensionFilter, minSizeBytes, maxSizeBytes, _
                entryIDs, storeIDs, matchedEmails, totalAttachments, totalBytes
        Next subFolder
    End If

FolderDone:
    On Error GoTo 0

End Sub


Private Sub AddMailIfEligible( _
    ByVal mail As Outlook.MailItem, _
    ByRef entryIDs As Collection, _
    ByRef storeIDs As Collection, _
    ByRef matchedEmails As Long, _
    ByRef totalAttachments As Long, _
    ByRef totalBytes As Double, _
    ByVal extensionFilter As String, _
    ByVal minSizeBytes As Double, _
    ByVal maxSizeBytes As Double)

    Dim att As Outlook.Attachment
    Dim matchingInMail As Long
    Dim matchingBytes As Double
    Dim entryID As String
    Dim storeID As String

    For Each att In mail.Attachments
        If AttachmentMatchesFilters( _
            att, extensionFilter, minSizeBytes, maxSizeBytes) Then

            matchingInMail = matchingInMail + 1
            matchingBytes = matchingBytes + CDbl(att.Size)
        End If
    Next att

    If matchingInMail = 0 Then Exit Sub

    entryID = ""
    storeID = ""

    On Error Resume Next
    entryID = mail.EntryID
    storeID = mail.Parent.StoreID
    On Error GoTo 0

    If entryID = "" Then Exit Sub

    entryIDs.Add entryID
    storeIDs.Add storeID
    matchedEmails = matchedEmails + 1
    totalAttachments = totalAttachments + matchingInMail
    totalBytes = totalBytes + matchingBytes

End Sub


Private Function MailMatchesFiltersV2( _
    ByVal mail As Outlook.MailItem, _
    ByVal senderFilter As String, _
    ByVal subjectFilter As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    ByVal useFromDate As Boolean, _
    ByVal useToDate As Boolean) As Boolean

    Dim senderName As String
    Dim senderEmail As String

    On Error Resume Next
    senderName = mail.SenderName
    senderEmail = GetSenderEmailV2(mail)
    On Error GoTo 0

    If Trim(senderFilter) <> "" Then
        If InStr(1, senderName, senderFilter, vbTextCompare) = 0 And _
           InStr(1, senderEmail, senderFilter, vbTextCompare) = 0 Then
            Exit Function
        End If
    End If

    If Trim(subjectFilter) <> "" Then
        If InStr(1, mail.Subject, subjectFilter, vbTextCompare) = 0 Then
            Exit Function
        End If
    End If

    If useFromDate Then
        If mail.ReceivedTime < fromDate Then Exit Function
    End If

    If useToDate Then
        If mail.ReceivedTime >= toDate Then Exit Function
    End If

    MailMatchesFiltersV2 = True

End Function


Private Function AttachmentMatchesFilters( _
    ByVal att As Outlook.Attachment, _
    ByVal extensionFilter As String, _
    ByVal minSizeBytes As Double, _
    ByVal maxSizeBytes As Double) As Boolean

    Dim ext As String

    ext = LCase$(GetFileExtensionV2(att.FileName))
    If Left$(ext, 1) = "." Then ext = Mid$(ext, 2)

    If extensionFilter <> "" Then
        If InStr(1, "," & extensionFilter & ",", "," & ext & ",", vbTextCompare) = 0 Then
            Exit Function
        End If
    End If

    If minSizeBytes > 0 Then
        If CDbl(att.Size) < minSizeBytes Then Exit Function
    End If

    If maxSizeBytes > 0 Then
        If CDbl(att.Size) > maxSizeBytes Then Exit Function
    End If

    AttachmentMatchesFilters = True

End Function


Private Sub SaveMailAttachmentsV2( _
    ByVal mail As Outlook.MailItem, _
    ByVal baseFolder As String, _
    ByVal saveMode As VbMsgBoxResult, _
    ByVal extensionFilter As String, _
    ByVal minSizeBytes As Double, _
    ByVal maxSizeBytes As Double, _
    ByRef processedFiles As Long, _
    ByRef savedFiles As Long, _
    ByRef failedFiles As Long, _
    ByVal totalAttachments As Long, _
    ByVal logStream As Object)

    Dim att As Outlook.Attachment
    Dim targetFolder As String
    Dim emailFolderName As String
    Dim fileName As String
    Dim finalPath As String
    Dim tempPath As String
    Dim counter As Long
    Dim saveSucceeded As Boolean
    Dim errorText As String

    If saveMode = vbNo Then

        emailFolderName = _
            Format(mail.ReceivedTime, "yyyy-mm-dd_hh-nn") & " - " & _
            CleanFileNameV2(GetSenderDisplayNameV2(mail)) & " - " & _
            CleanFileNameV2(mail.Subject)

        emailFolderName = LimitFolderNameForPath(baseFolder, emailFolderName)
        targetFolder = GetUniqueFolderPathV2(baseFolder, emailFolderName)

        If Not FolderExistsV2(targetFolder) Then
            On Error Resume Next
            Err.Clear
            MkDir targetFolder

            If Err.Number <> 0 Then
                errorText = "Could not create destination folder: " & Err.Description
                On Error GoTo 0

                For Each att In mail.Attachments
                    If AttachmentMatchesFilters( _
                        att, extensionFilter, minSizeBytes, maxSizeBytes) Then

                        failedFiles = failedFiles + 1
                        processedFiles = processedFiles + 1
                        WriteLogLine logStream, mail, att, "FAILED", "", errorText
                    End If
                Next att

                UpdateProgressToolbarV2 _
                    processedFiles, totalAttachments, "Could not create destination folder"
                Exit Sub
            End If
            On Error GoTo 0
        End If

        targetFolder = targetFolder & "\"

    Else
        targetFolder = baseFolder
    End If

    For Each att In mail.Attachments

        If Not AttachmentMatchesFilters( _
            att, extensionFilter, minSizeBytes, maxSizeBytes) Then
            GoTo NextAttachment
        End If

        If EscapePressedV2() Then
            gCancelled = True
            Exit Sub
        End If

        fileName = CleanFileNameV2(att.FileName)
        fileName = LimitFileNameLengthV2( _
            targetFolder, fileName, MAX_PATH_LENGTH - TEMP_SUFFIX_RESERVE)

        finalPath = targetFolder & fileName
        counter = 1

        Do While FileExistsV2(finalPath)
            finalPath = _
                targetFolder & GetFileNameWithoutExtensionV2(fileName) & _
                "_" & counter & GetFileExtensionV2(fileName)
            counter = counter + 1
        Loop

        tempPath = GetTemporaryAttachmentPathV2(finalPath)

        UpdateProgressToolbarV2 _
            processedFiles, totalAttachments, _
            "Downloading: " & ShortDisplayNameV2(fileName, 55) & _
            " (" & FormatBytesV2(CDbl(att.Size)) & ")"

        saveSucceeded = False
        errorText = ""

        On Error Resume Next
        Err.Clear
        att.SaveAsFile tempPath

        If Err.Number <> 0 Then
            errorText = "SaveAsFile: " & Err.Number & " - " & Err.Description
        ElseIf Not FileExistsV2(tempPath) Then
            errorText = "Temporary file was not created."
        Else
            Err.Clear
            Name tempPath As finalPath

            If Err.Number = 0 Then
                saveSucceeded = True
            Else
                errorText = "Rename temp file: " & Err.Number & " - " & Err.Description
            End If
        End If
        On Error GoTo 0

        processedFiles = processedFiles + 1

        If saveSucceeded Then
            savedFiles = savedFiles + 1
            WriteLogLine logStream, mail, att, "SUCCESS", finalPath, ""
        Else
            failedFiles = failedFiles + 1
            DeleteFileSafeV2 tempPath
            WriteLogLine logStream, mail, att, "FAILED", finalPath, errorText
        End If

        UpdateProgressToolbarV2 _
            processedFiles, totalAttachments, _
            "Downloaded: " & Format(savedFiles, "#,##0") & _
            "   Failed: " & Format(failedFiles, "#,##0")

        If processedFiles Mod UI_YIELD_INTERVAL = 0 Then DoEvents

        If EscapePressedV2() Then
            gCancelled = True
            Exit Sub
        End If

NextAttachment:
    Next att

End Sub


Private Sub WriteLogLine( _
    ByVal logStream As Object, _
    ByVal mail As Outlook.MailItem, _
    ByVal att As Outlook.Attachment, _
    ByVal status As String, _
    ByVal savedPath As String, _
    ByVal errorText As String)

    If logStream Is Nothing Then Exit Sub

    logStream.WriteLine CsvLine(Array( _
        Format(Now, "yyyy-mm-dd hh:nn:ss"), _
        status, _
        GetSenderDisplayNameV2(mail), _
        GetSenderEmailV2(mail), _
        Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn:ss"), _
        mail.Subject, _
        att.FileName, _
        CStr(att.Size), _
        savedPath, _
        errorText))

End Sub


Private Function CsvLine(ByVal values As Variant) As String

    Dim i As Long
    Dim s As String
    Dim item As String

    For i = LBound(values) To UBound(values)
        item = CStr(values(i))
        item = Replace(item, """", """"")
        item = """" & item & """"

        If i > LBound(values) Then s = s & ","
        s = s & item
    Next i

    CsvLine = s

End Function


Private Function NormalizeExtensionFilter(ByVal rawFilter As String) As String

    Dim s As String
    Dim parts As Variant
    Dim p As Variant
    Dim result As String
    Dim token As String

    s = LCase$(Trim$(rawFilter))
    s = Replace(s, ";", ",")
    s = Replace(s, " ", "")

    If s = "" Then
        NormalizeExtensionFilter = ""
        Exit Function
    End If

    parts = Split(s, ",")

    For Each p In parts
        token = Trim$(CStr(p))
        If Left$(token, 1) = "." Then token = Mid$(token, 2)

        If token <> "" Then
            If InStr(1, "," & result & ",", "," & token & ",", vbTextCompare) = 0 Then
                If result <> "" Then result = result & ","
                result = result & token
            End If
        End If
    Next p

    NormalizeExtensionFilter = result

End Function


Private Function DisplayExtensionFilter(ByVal extensionFilter As String) As String
    If extensionFilter = "" Then
        DisplayExtensionFilter = "All"
    Else
        DisplayExtensionFilter = extensionFilter
    End If
End Function


Private Function LimitFolderNameForPath( _
    ByVal baseFolder As String, _
    ByVal folderName As String) As String

    Dim available As Long

    available = EMAIL_FOLDER_PATH_LIMIT - Len(baseFolder)
    If available < 20 Then available = 20

    If Len(folderName) > available Then
        folderName = Left$(folderName, available)
    End If

    Do While Len(folderName) > 0 And _
        (Right$(folderName, 1) = "." Or Right$(folderName, 1) = " ")
        folderName = Left$(folderName, Len(folderName) - 1)
    Loop

    If folderName = "" Then folderName = "Email"
    LimitFolderNameForPath = folderName

End Function


Private Function GetSenderEmailV2(ByVal mail As Outlook.MailItem) As String

    Dim exchUser As Outlook.ExchangeUser
    Dim senderObj As Outlook.AddressEntry

    On Error Resume Next

    If mail.SenderEmailType = "EX" Then
        Set senderObj = mail.Sender

        If Not senderObj Is Nothing Then
            Set exchUser = senderObj.GetExchangeUser

            If Not exchUser Is Nothing Then
                GetSenderEmailV2 = exchUser.PrimarySmtpAddress
            Else
                GetSenderEmailV2 = mail.SenderEmailAddress
            End If
        Else
            GetSenderEmailV2 = mail.SenderEmailAddress
        End If
    Else
        GetSenderEmailV2 = mail.SenderEmailAddress
    End If

    On Error GoTo 0

End Function


Private Function GetSenderDisplayNameV2(ByVal mail As Outlook.MailItem) As String

    On Error Resume Next

    If Trim$(mail.SenderName) <> "" Then
        GetSenderDisplayNameV2 = mail.SenderName
    Else
        GetSenderDisplayNameV2 = GetSenderEmailV2(mail)
    End If

    If Trim$(GetSenderDisplayNameV2) = "" Then
        GetSenderDisplayNameV2 = "Unknown Sender"
    End If

    On Error GoTo 0

End Function


Private Function CleanFileNameV2(ByVal fileName As String) As String

    Dim invalidChars As Variant
    Dim ch As Variant

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each ch In invalidChars
        fileName = Replace(fileName, ch, "_")
    Next ch

    fileName = Trim$(fileName)

    Do While Len(fileName) > 0
        If Right$(fileName, 1) = "." Or Right$(fileName, 1) = " " Then
            fileName = Left$(fileName, Len(fileName) - 1)
        Else
            Exit Do
        End If
    Loop

    If fileName = "" Then fileName = "Unnamed"
    CleanFileNameV2 = fileName

End Function


Private Function GetFileNameWithoutExtensionV2(ByVal fileName As String) As String

    Dim p As Long
    p = InStrRev(fileName, ".")

    If p > 1 Then
        GetFileNameWithoutExtensionV2 = Left$(fileName, p - 1)
    Else
        GetFileNameWithoutExtensionV2 = fileName
    End If

End Function


Private Function GetFileExtensionV2(ByVal fileName As String) As String

    Dim p As Long
    p = InStrRev(fileName, ".")

    If p > 1 Then
        GetFileExtensionV2 = Mid$(fileName, p)
    Else
        GetFileExtensionV2 = ""
    End If

End Function


Private Function LimitFileNameLengthV2( _
    ByVal folderPath As String, _
    ByVal fileName As String, _
    ByVal maxFullPath As Long) As String

    Dim ext As String
    Dim baseName As String
    Dim available As Long

    ext = GetFileExtensionV2(fileName)
    baseName = GetFileNameWithoutExtensionV2(fileName)

    available = maxFullPath - Len(folderPath) - Len(ext)
    If available < 10 Then available = 10

    If Len(baseName) > available Then
        baseName = Left$(baseName, available)
    End If

    LimitFileNameLengthV2 = baseName & ext

End Function


Private Function FileExistsV2(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExistsV2 = _
        (Len(Dir(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly Or vbArchive)) > 0)
    On Error GoTo 0
End Function


Private Function FolderExistsV2(ByVal folderPath As String) As Boolean

    Dim attr As Long

    On Error Resume Next
    Err.Clear
    attr = GetAttr(folderPath)

    If Err.Number = 0 Then
        FolderExistsV2 = ((attr And vbDirectory) = vbDirectory)
    Else
        FolderExistsV2 = False
    End If

    Err.Clear
    On Error GoTo 0

End Function


Private Function GetUniqueFolderPathV2( _
    ByVal baseFolder As String, _
    ByVal folderName As String) As String

    Dim folderPath As String
    Dim counter As Long
    Dim suffix As String
    Dim safeName As String

    safeName = folderName
    folderPath = baseFolder & safeName
    counter = 1

    Do While FolderExistsV2(folderPath)
        suffix = "_" & counter
        safeName = folderName

        If Len(baseFolder & safeName & suffix) > EMAIL_FOLDER_PATH_LIMIT Then
            safeName = Left$(safeName, _
                EMAIL_FOLDER_PATH_LIMIT - Len(baseFolder) - Len(suffix))
        End If

        folderPath = baseFolder & safeName & suffix
        counter = counter + 1
    Loop

    GetUniqueFolderPathV2 = folderPath

End Function


Private Function GetTemporaryAttachmentPathV2(ByVal finalPath As String) As String

    Dim tempPath As String
    Dim counter As Long

    tempPath = finalPath & ".downloading"
    counter = 1

    Do While FileExistsV2(tempPath)
        tempPath = finalPath & ".downloading_" & counter
        counter = counter + 1
    Loop

    GetTemporaryAttachmentPathV2 = tempPath

End Function


Private Sub DeleteFileSafeV2(ByVal filePath As String)
    On Error Resume Next
    If FileExistsV2(filePath) Then Kill filePath
    Err.Clear
    On Error GoTo 0
End Sub


Private Function ShortDisplayNameV2(ByVal text As String, ByVal maxLength As Long) As String
    If Len(text) <= maxLength Then
        ShortDisplayNameV2 = text
    Else
        ShortDisplayNameV2 = Left$(text, maxLength - 3) & "..."
    End If
End Function


Private Function FormatBytesV2(ByVal bytes As Double) As String

    If bytes >= 1099511627776# Then
        FormatBytesV2 = Format(bytes / 1099511627776#, "0.00") & " TB"
    ElseIf bytes >= 1073741824# Then
        FormatBytesV2 = Format(bytes / 1073741824#, "0.00") & " GB"
    ElseIf bytes >= 1048576# Then
        FormatBytesV2 = Format(bytes / 1048576#, "0.00") & " MB"
    ElseIf bytes >= 1024# Then
        FormatBytesV2 = Format(bytes / 1024#, "0.0") & " KB"
    Else
        FormatBytesV2 = Format(bytes, "0") & " bytes"
    End If

End Function


Private Function HasEnoughDiskSpaceV2( _
    ByVal folderPath As String, _
    ByVal requiredBytes As Double) As Boolean

    Dim fso As Object
    Dim drv As Object
    Dim driveLetter As String

    On Error GoTo UnableToCheck

    If Left$(folderPath, 2) = "\\" Then
        HasEnoughDiskSpaceV2 = True
        Exit Function
    End If

    driveLetter = Left$(folderPath, 2)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set drv = fso.GetDrive(driveLetter)

    HasEnoughDiskSpaceV2 = _
        CDbl(drv.FreeSpace) >= (requiredBytes * FREE_SPACE_MARGIN)
    Exit Function

UnableToCheck:
    HasEnoughDiskSpaceV2 = True

End Function


Private Sub CreateProgressToolbarV2()

    On Error Resume Next

    DestroyProgressToolbarV2

    Set gProgressBar = Application.ActiveExplorer.CommandBars.Add( _
        Name:="OutlookAttachSnatcherV2Progress", _
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


Private Sub UpdateProgressToolbarV2( _
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

    If Not gProgressStatus Is Nothing Then
        gProgressStatus.Caption = statusText
    End If

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


Private Sub DestroyProgressToolbarV2()

    On Error Resume Next

    If Not gProgressBar Is Nothing Then gProgressBar.Delete

    Set gProgressCancel = Nothing
    Set gProgressCount = Nothing
    Set gProgressStatus = Nothing
    Set gProgressBar = Nothing

    On Error GoTo 0

End Sub


Private Function EscapePressedV2() As Boolean
    DoEvents
    EscapePressedV2 = ((GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0)
End Function
