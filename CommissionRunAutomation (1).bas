'=============================================================================
' COMMISSION RUN AUTOMATION - NAB Commercial Broker
' Master automation: Report 1 (SharePoint) + Report 2 (BI) + Distribution List + Email
'=============================================================================
' HOW TO USE:
' 1. Open Excel, press Alt+F11 to open VBA Editor
' 2. Insert > Module, paste this code
' 3. Close VBA Editor, press Alt+F8, select "RunCommissionAutomation" and click Run
' 4. Follow the prompts — it collects ALL inputs upfront, then processes everything
'=============================================================================

Option Explicit

' ---- Global variables set by the input prompts ----
Private gRunMonthDate As Date        ' 1st day of the run month (e.g. 1/Aug/2026)
Private gRunMonthName As String      ' e.g. "August"
Private gRunMonthYear As Long        ' e.g. 2026
Private gRunMonthMM As String        ' e.g. "08"
Private gFiscalYear As String        ' e.g. "FY26"
Private gTodayStr As String          ' e.g. "01.09.2026"
Private gSaveFolderPath As String    ' full path to save folder

' ---- File paths collected upfront ----
Private gSharePointCSVPath As String       ' user-selected SharePoint CSV
Private gFullLoanSummaryPath As String     ' Full Loan Summary (auto-found or user-selected)
Private gVerificationReportPath As String  ' Verification Report (auto-found or user-selected)
Private gMBLPath As String                 ' Master BUID List (user-selected)
Private gSharedMailboxRoot As Object        ' Root folder of nabcommercialbroker shared mailbox
Private gSharedDriveLetter As String         ' Detected drive letter for mapped share drive (e.g. "Z", "S", "X")

' ---- DEBUG MODE: Set to True during development, False for production ----
' When True: skips MsgBox confirmations, auto-confirms dates, shows extra debug info
' When False: normal operation with all prompts
#Const DEBUG_MODE = False

' ---- Logging helper ----
Private Sub WriteToLog(msg As String)
    #If DEBUG_MODE Then
        Debug.Print Format(Now, "hh:nn:ss") & " - " & msg
    #End If
End Sub

'=============================================================================
' MASTER ENTRY POINT - One button does everything
'=============================================================================
Public Sub RunCommissionAutomation()

    ' ---- Phase 0: Detect shared drive letter ----
    ' Scans mapped drives for "NAB COMMERCIAL BROKER" folder so the code
    ' works regardless of which letter the user chose (Z:, S:, X:, etc.)
    gSharedDriveLetter = DetectSharedDriveLetter()

    If gSharedDriveLetter = "" Then
        MsgBox "Could not determine the shared network drive." & vbNewLine & vbNewLine & _
               "The macro needs the mapped drive containing 'NAB COMMERCIAL BROKER'." & vbNewLine & _
               "Please map the shared drive and try again.", vbCritical, "Drive Not Found"
        Exit Sub
    End If

    ' ---- Phase 1: Collect ALL inputs upfront ----
    ' This way the user provides everything at the start,
    ' then the macro runs uninterrupted.

    ' Step 1: Get run month
    If Not GetRunMonthInput() Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If

    ' Step 2: Collect all file paths
    If Not CollectAllFilePaths() Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If

    ' Step 3: Show final confirmation with everything
    If Not ShowFinalConfirmation() Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If

    ' ---- Phase 2: Process everything ----
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Report 1: SharePoint CSV
    Application.StatusBar = "Processing Report 1 of 3 — SharePoint CSV..."
    Dim report1Path As String
    report1Path = ProcessReport1_SharePoint()

    If report1Path = "" Then
        MsgBox "Report 1 failed. Please check the error and try again.", vbExclamation
        GoTo CleanUp
    End If

    ' Report 2: BI Assist (Full Loan Summary + Verification Report)
    Application.StatusBar = "Processing Report 2 of 3 — BI Assist..."
    Dim report2Path As String
    report2Path = ProcessReport2_BIAssist()

    If report2Path = "" Then
        MsgBox "Report 2 failed. Please check the error and try again.", vbExclamation
        GoTo CleanUp
    End If

    ' Report 3: Distribution List
    Application.StatusBar = "Processing Report 3 of 3 — Distribution List..."
    Dim report3Path As String
    report3Path = ProcessReport3_DistributionList()

    If report3Path = "" Then
        MsgBox "Report 3 failed. Please check the error and try again.", vbExclamation
        GoTo CleanUp
    End If

    ' Email: Draft reminder emails in Outlook
    Application.StatusBar = "Creating Outlook email drafts..."
    CreateReminderEmails report1Path, report2Path

    ' Plain MsgBox (NO emoji — MsgBox shows them as "??" on Windows) lists paths.
    ' The celebration is handled by the HTML success popup inside CreateReminderEmails.
    MsgBox "Commission run completed!" & vbNewLine & vbNewLine & _
           "Report 1 saved to:" & vbNewLine & report1Path & vbNewLine & vbNewLine & _
           "Report 2 saved to:" & vbNewLine & report2Path & vbNewLine & vbNewLine & _
           "Report 3 saved to:" & vbNewLine & report3Path & vbNewLine & vbNewLine & _
           "Email drafts created in nabcommercialbroker@nab.com.au Drafts folder!", _
           vbInformation, "Success"

CleanUp:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

'=============================================================================
' COLLECT ALL FILE PATHS UPFRONT
' Asks user for every file needed before any processing begins
'=============================================================================
Private Function CollectAllFilePaths() As Boolean

    CollectAllFilePaths = False

    ' ---- 1. SharePoint CSV (always ask — could be anywhere) ----
    MsgBox "Step 1 of 3: Select the SharePoint CSV file" & vbNewLine & vbNewLine & _
           "This is the 'All Workflow Items' export you downloaded from SharePoint." & vbNewLine & _
           "(Usually in your Downloads folder)", _
           vbInformation, "Select Files - SharePoint CSV"

    gSharePointCSVPath = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        Title:="Step 1/3: Select the SharePoint CSV file")

    If gSharePointCSVPath = "False" Or gSharePointCSVPath = "" Then
        Exit Function
    End If

    ' ---- 2. Full Loan Summary (auto-find from shared drive, confirm or override) ----
    Dim flsFolder As String
    flsFolder = gSharedDriveLetter & ":\NAB COMMERCIAL BROKER\Commission Run\BI Reports\Full Loan Summary\"

    gFullLoanSummaryPath = FindMostRecentFile(flsFolder, "NCB_Full_Loan_Summary_*.xls*")

    If gFullLoanSummaryPath <> "" Then
        ' Found one — ask user to confirm
        Dim flsConfirm As VbMsgBoxResult
        flsConfirm = MsgBox( _
            "Step 2 of 3: Full Loan Summary" & vbNewLine & vbNewLine & _
            "I found the most recent file:" & vbNewLine & _
            gFullLoanSummaryPath & vbNewLine & vbNewLine & _
            "Is this the correct file?" & vbNewLine & vbNewLine & _
            "Click YES to use this file" & vbNewLine & _
            "Click NO to browse for a different one", _
            vbYesNoCancel + vbQuestion, "Confirm Full Loan Summary")

        If flsConfirm = vbCancel Then Exit Function

        If flsConfirm = vbNo Then
            gFullLoanSummaryPath = Application.GetOpenFilename( _
                FileFilter:="Excel Files (*.xls*),*.xls*,All Files (*.*),*.*", _
                Title:="Step 2/3: Select the Full Loan Summary file")

            If gFullLoanSummaryPath = "False" Or gFullLoanSummaryPath = "" Then
                Exit Function
            End If
        End If
    Else
        ' Couldn't find it — ask user to browse
        MsgBox "Step 2 of 3: Full Loan Summary" & vbNewLine & vbNewLine & _
               "Could not auto-find the file in:" & vbNewLine & _
               flsFolder & vbNewLine & vbNewLine & _
               "Please browse to select it manually.", _
               vbInformation, "Select Full Loan Summary"

        gFullLoanSummaryPath = Application.GetOpenFilename( _
            FileFilter:="Excel Files (*.xls*),*.xls*,All Files (*.*),*.*", _
            Title:="Step 2/3: Select the Full Loan Summary file")

        If gFullLoanSummaryPath = "False" Or gFullLoanSummaryPath = "" Then
            Exit Function
        End If
    End If

    ' ---- 3. Verification Report (auto-find from shared drive, confirm or override) ----
    Dim vrFolder As String
    vrFolder = gSharedDriveLetter & ":\NAB COMMISSIONS\Monthly Validations - Retail Commission Payments\BI Reports\Verification Report\"

    gVerificationReportPath = FindMostRecentFile(vrFolder, "Verification Report_*.xlsx")

    If gVerificationReportPath <> "" Then
        Dim vrConfirm As VbMsgBoxResult
        vrConfirm = MsgBox( _
            "Step 3 of 3: Verification Report" & vbNewLine & vbNewLine & _
            "I found the most recent file:" & vbNewLine & _
            gVerificationReportPath & vbNewLine & vbNewLine & _
            "Is this the correct file?" & vbNewLine & vbNewLine & _
            "Click YES to use this file" & vbNewLine & _
            "Click NO to browse for a different one", _
            vbYesNoCancel + vbQuestion, "Confirm Verification Report")

        If vrConfirm = vbCancel Then Exit Function

        If vrConfirm = vbNo Then
            gVerificationReportPath = Application.GetOpenFilename( _
                FileFilter:="Excel Files (*.xls*),*.xls*,All Files (*.*),*.*", _
                Title:="Step 3/3: Select the Verification Report file")

            If gVerificationReportPath = "False" Or gVerificationReportPath = "" Then
                Exit Function
            End If
        End If
    Else
        MsgBox "Step 3 of 3: Verification Report" & vbNewLine & vbNewLine & _
               "Could not auto-find the file in:" & vbNewLine & _
               vrFolder & vbNewLine & vbNewLine & _
               "Please browse to select it manually.", _
               vbInformation, "Select Verification Report"

        gVerificationReportPath = Application.GetOpenFilename( _
            FileFilter:="Excel Files (*.xls*),*.xls*,All Files (*.*),*.*", _
            Title:="Step 3/3: Select the Verification Report file")

        If gVerificationReportPath = "False" Or gVerificationReportPath = "" Then
            Exit Function
        End If
    End If

    ' ---- 4. Master BUID List (always ask — comes from Teams chat, could be anywhere) ----
    MsgBox "Step 4 of 4: Select the Master BUID List (MBL)" & vbNewLine & vbNewLine & _
           "This is the most recent BUID list from your Teams chat." & vbNewLine & _
           "It has 3 tabs: Commercial_Broker_BABs_BUID_MAS, BDSB Bankers, SB Health Bankers", _
           vbInformation, "Select Files - Master BUID List"

    gMBLPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xls*),*.xls*,All Files (*.*),*.*", _
        Title:="Step 4/4: Select the Master BUID List (MBL)")

    If gMBLPath = "False" Or gMBLPath = "" Then
        Exit Function
    End If

    CollectAllFilePaths = True

End Function

'=============================================================================
' SHOW FINAL CONFIRMATION - Summary of everything before processing
'=============================================================================
Private Function ShowFinalConfirmation() As Boolean

    ShowFinalConfirmation = False

    ' Extract just filenames for cleaner display
    Dim csvName As String: csvName = ExtractFileName(gSharePointCSVPath)
    Dim flsName As String: flsName = ExtractFileName(gFullLoanSummaryPath)
    Dim vrName As String: vrName = ExtractFileName(gVerificationReportPath)
    Dim mblName As String: mblName = ExtractFileName(gMBLPath)

    Dim msg As String
    msg = "FINAL CONFIRMATION - Please review everything:" & vbNewLine & vbNewLine & _
          "═══ Run Details ═══" & vbNewLine & _
          "  Run Month:    " & gRunMonthName & " " & gRunMonthYear & vbNewLine & _
          "  Fiscal Year:  " & gFiscalYear & vbNewLine & _
          "  Today's Date: " & gTodayStr & vbNewLine & _
          "  Save Folder:  " & gSaveFolderPath & vbNewLine & vbNewLine & _
          "═══ Files Selected ═══" & vbNewLine & _
          "  1. SharePoint CSV:       " & csvName & vbNewLine & _
          "  2. Full Loan Summary:    " & flsName & vbNewLine & _
          "  3. Verification Report:  " & vrName & vbNewLine & _
          "  4. Master BUID List:     " & mblName & vbNewLine & vbNewLine & _
          "Click YES to start processing all reports." & vbNewLine & _
          "Click NO to cancel."

    #If DEBUG_MODE Then
        WriteToLog "DEBUG_MODE: Auto-confirming final confirmation"
        ShowFinalConfirmation = True
    #Else
        If MsgBox(msg, vbYesNo + vbQuestion, "Confirm & Start") = vbYes Then
            ShowFinalConfirmation = True
        End If
    #End If

End Function

'=============================================================================
' FIND THE MOST RECENT FILE IN A FOLDER MATCHING A PATTERN
' Returns full path of newest file, or "" if none found
'=============================================================================
Private Function FindMostRecentFile(folderPath As String, filePattern As String) As String

    Dim fileName As String
    Dim newestFile As String
    Dim newestDate As Date
    Dim fileDate As Date

    FindMostRecentFile = ""

    ' Check if folder exists
    On Error Resume Next
    If Dir(folderPath, vbDirectory) = "" Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' Loop through matching files
    fileName = Dir(folderPath & filePattern)

    Do While fileName <> ""
        fileDate = FileDateTime(folderPath & fileName)
        If newestFile = "" Or fileDate > newestDate Then
            newestFile = fileName
            newestDate = fileDate
        End If
        fileName = Dir()
    Loop

    If newestFile <> "" Then
        FindMostRecentFile = folderPath & newestFile
    End If

End Function

'=============================================================================
' EXTRACT FILENAME FROM FULL PATH
'=============================================================================
Private Function ExtractFileName(fullPath As String) As String

    Dim pos As Long
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        ExtractFileName = Mid(fullPath, pos + 1)
    Else
        ExtractFileName = fullPath
    End If

End Function

'=============================================================================
' DETECT THE SHARED DRIVE LETTER
' Scans all mapped drives (D: to Z:) looking for "NAB COMMERCIAL BROKER" folder.
' Returns the drive letter (e.g. "Z") or "" if not found.
' If not found, prompts user to browse for the shared drive root.
'=============================================================================
Private Function DetectSharedDriveLetter() As String

    Dim driveLetter As String
    Dim testPath As String
    Dim i As Long

    DetectSharedDriveLetter = ""

    ' Scan from Z down to D — most people pick letters at the end of the alphabet
    For i = Asc("Z") To Asc("D") Step -1
        driveLetter = Chr(i)
        testPath = driveLetter & ":\NAB COMMERCIAL BROKER"

        On Error Resume Next
        If Dir(testPath, vbDirectory) <> "" Then
            On Error GoTo 0
            DetectSharedDriveLetter = driveLetter
            Exit Function
        End If
        On Error GoTo 0
    Next i

    ' If not found, ask the user to browse
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox( _
        "Could not auto-detect the shared network drive." & vbNewLine & vbNewLine & _
        "The folder 'NAB COMMERCIAL BROKER' was not found on any mapped drive (D: to Z:)." & vbNewLine & vbNewLine & _
        "Click YES to browse and select the drive/folder manually." & vbNewLine & _
        "Click NO to cancel.", _
        vbYesNo + vbExclamation, "Shared Drive Not Found")

    If userResponse = vbYes Then
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)

        With fd
            .Title = "Select the root of your shared network drive (e.g. Z:\)"
            If .Show = -1 Then
                Dim selectedPath As String
                selectedPath = .SelectedItems(1)

                ' Extract the drive letter from the selected path
                If Len(selectedPath) >= 2 And Mid(selectedPath, 2, 1) = ":" Then
                    DetectSharedDriveLetter = Left(selectedPath, 1)
                Else
                    ' UNC path or something unexpected — store the first letter anyway
                    ' but this case is unlikely for a mapped drive
                    MsgBox "The selected path doesn't appear to be a mapped drive letter." & vbNewLine & _
                           "Path: " & selectedPath & vbNewLine & vbNewLine & _
                           "The macro will try to use it, but auto-detection may not work correctly.", _
                           vbExclamation
                    DetectSharedDriveLetter = Left(selectedPath, 1)
                End If
            End If
        End With
    End If

End Function

'=============================================================================
' BROWSE FOR FOLDER (lets user pick a folder via dialog)
'=============================================================================
Private Function BrowseForFolder(promptText As String) As String

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = promptText
        .InitialFileName = gSaveFolderPath
        If .Show = -1 Then
            BrowseForFolder = .SelectedItems(1)
        Else
            BrowseForFolder = ""
        End If
    End With

End Function

'=============================================================================
' REPORT 1: SHAREPOINT CSV PROCESSING
' Now uses the pre-collected gSharePointCSVPath instead of asking mid-run
' Returns the save path on success, "" on failure
'=============================================================================
Private Function ProcessReport1_SharePoint() As String

    Dim wbCSV As Workbook
    Dim wbNew As Workbook
    Dim wsSource As Worksheet
    Dim wsTab1 As Worksheet
    Dim wsTab2 As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim savePath As String

    ProcessReport1_SharePoint = ""

    On Error GoTo Report1Error

    ' ---- Open the CSV ----
    Application.StatusBar = "Report 1: Opening CSV file..."
    Set wbCSV = Workbooks.Open(gSharePointCSVPath, Local:=True)
    Set wsSource = wbCSV.Sheets(1)

    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    ' ---- Create new workbook ----
    Set wbNew = Workbooks.Add
    Do While wbNew.Sheets.Count > 1
        wbNew.Sheets(wbNew.Sheets.Count).Delete
    Loop

    ' ---- Tab 1: "Payments yet to be reviewed" ----
    Application.StatusBar = "Report 1: Filtering for Tab 1 - Payments yet to be reviewed..."
    Set wsTab1 = wbNew.Sheets(1)
    wsTab1.Name = "Payments yet to be reviewed"

    CopyFilteredData wsSource, wsTab1, lastRow, lastCol, _
        FilterType:="NewAndAllocated", _
        RunMonthFilter:=True

    ' ---- Tab 2: "Payments returned" ----
    Application.StatusBar = "Report 1: Filtering for Tab 2 - Payments returned..."
    Set wsTab2 = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.Count))
    wsTab2.Name = "Payments returned"

    CopyFilteredData wsSource, wsTab2, lastRow, lastCol, _
        FilterType:="ReturnedError", _
        RunMonthFilter:=False

    ' ---- Format both tabs ----
    Application.StatusBar = "Report 1: Formatting Tab 1..."
    FormatSharePointTab wsTab1

    Application.StatusBar = "Report 1: Formatting Tab 2..."
    FormatSharePointTab wsTab2

    ' ---- Save ----
    savePath = gSaveFolderPath & "NCB Payment Sharepoint " & gTodayStr & ".xlsx"

    Application.StatusBar = "Report 1: Saving..."
    wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook

    ' Close the CSV
    wbCSV.Close SaveChanges:=False

    ProcessReport1_SharePoint = savePath
    Exit Function

Report1Error:
    MsgBox "Error in Report 1: " & Err.Description, vbExclamation, "Report 1 Error"
    ProcessReport1_SharePoint = ""

End Function

'=============================================================================
' REPORT 2: BI ASSIST (Full Loan Summary + Verification Report)
' Creates one workbook with 2 tabs:
'   Tab 1: "Sharepoint Payments" from Full Loan Summary (filtered by Month Paid)
'   Tab 2: "NDF & TF Payments" from Verification Report (filtered by Month Paid + Product = NDF or TF)
' Returns the save path on success, "" on failure
'=============================================================================
Private Function ProcessReport2_BIAssist() As String

    Dim wbFLS As Workbook
    Dim wbVR As Workbook
    Dim wbNew As Workbook
    Dim wsTab1 As Worksheet
    Dim wsTab2 As Worksheet
    Dim flsHeaderRow As Long
    Dim vrHeaderRow As Long
    Dim savePath As String

    ProcessReport2_BIAssist = ""

    On Error GoTo Report2Error

    ' ---- PART A: Full Loan Summary ----
    Application.StatusBar = "Report 2: Opening Full Loan Summary..."
    Set wbFLS = Workbooks.Open(gFullLoanSummaryPath, Local:=True)
    Dim wsFLS As Worksheet
    Set wsFLS = wbFLS.Sheets(1)

    ' Find header row (smart detection — look for "BANKER NAME" or similar)
    flsHeaderRow = FindHeaderRow(wsFLS, "BANKER")
    If flsHeaderRow = 0 Then
        MsgBox "Could not find header row in Full Loan Summary. File format may have changed.", vbExclamation
        wbFLS.Close SaveChanges:=False
        Exit Function
    End If

    ' Delete rows before the header
    If flsHeaderRow > 1 Then
        wsFLS.Rows("1:" & flsHeaderRow - 1).Delete
        flsHeaderRow = 1
    End If

    ' Now flsHeaderRow is row 1
    Dim flsLastRow As Long, flsLastCol As Long
    flsLastRow = wsFLS.Cells(wsFLS.Rows.Count, "A").End(xlUp).Row
    flsLastCol = wsFLS.Cells(1, wsFLS.Columns.Count).End(xlToLeft).Column

    ' NOTE: Text to Columns will be applied ONLY on specific columns in the output tabs
    ' (not on source data to avoid corrupting it)

    ' Create new workbook for both tabs
    Set wbNew = Workbooks.Add
    Do While wbNew.Sheets.Count > 1
        wbNew.Sheets(wbNew.Sheets.Count).Delete
    Loop

    ' Filter and copy Full Loan Summary → Tab 1
    Application.StatusBar = "Report 2: Filtering Full Loan Summary..."
    Set wsTab1 = wbNew.Sheets(1)
    wsTab1.Name = "Sharepoint Payments"

    CopyFilteredDataReport2 wsFLS, wsTab1, flsLastRow, flsLastCol, "FullLoanSummary", True

    ' ---- PART B: Verification Report ----
    Application.StatusBar = "Report 2: Opening Verification Report..."
    Set wbVR = Workbooks.Open(gVerificationReportPath, Local:=True)
    Dim wsVR As Worksheet
    Set wsVR = wbVR.Sheets(1)

    ' Find header row (smart detection — look for "BANKER NAME" or similar)
    vrHeaderRow = FindHeaderRow(wsVR, "BANKER")
    If vrHeaderRow = 0 Then
        MsgBox "Could not find header row in Verification Report. File format may have changed.", vbExclamation
        wbVR.Close SaveChanges:=False
        wbFLS.Close SaveChanges:=False
        Exit Function
    End If

    ' Delete rows before the header
    If vrHeaderRow > 1 Then
        wsVR.Rows("1:" & vrHeaderRow - 1).Delete
        vrHeaderRow = 1
    End If

    ' Now vrHeaderRow is row 1
    Dim vrLastRow As Long, vrLastCol As Long
    vrLastRow = wsVR.Cells(wsVR.Rows.Count, "A").End(xlUp).Row
    vrLastCol = wsVR.Cells(1, wsVR.Columns.Count).End(xlToLeft).Column

    ' NOTE: Text to Columns will be applied ONLY on specific columns in the output tabs
    ' (not on source data to avoid corrupting it)

    ' Filter and copy Verification Report → Tab 2
    Application.StatusBar = "Report 2: Filtering Verification Report..."
    Set wsTab2 = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.Count))
    wsTab2.Name = "NDF & TF Payments"

    CopyFilteredDataReport2 wsVR, wsTab2, vrLastRow, vrLastCol, "VerificationReport", True

    ' ---- Format both tabs ----
    Application.StatusBar = "Report 2: Formatting Tab 1 - Full Loan Summary..."
    FormatReport2Tab1 wsTab1

    Application.StatusBar = "Report 2: Formatting Tab 2 - Verification Report..."
    FormatReport2Tab2 wsTab2

    ' ---- Save ----
    savePath = gSaveFolderPath & "NCB - " & gRunMonthName & " " & gRunMonthYear & " Payments " & gTodayStr & ".xlsx"

    Application.StatusBar = "Report 2: Saving..."
    wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook

    ' Close source files
    wbFLS.Close SaveChanges:=False
    wbVR.Close SaveChanges:=False

    ProcessReport2_BIAssist = savePath
    Exit Function

Report2Error:
    MsgBox "Error in Report 2: " & Err.Description, vbExclamation, "Report 2 Error"
    ProcessReport2_BIAssist = ""

End Function

'=============================================================================
' FIND HEADER ROW BY SEARCHING FOR A KEY WORD
' Smart detection — looks for a specific header name (e.g. "BANKER")
'=============================================================================
Private Function FindHeaderRow(ws As Worksheet, keywordHeader As String) As Long

    Dim i As Long
    Dim lastRow As Long
    Dim searchLimit As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    searchLimit = Application.Min(lastRow, 100)  ' Only search first 100 rows to be safe

    For i = 1 To searchLimit
        ' Check if any cell in this row contains the keyword
        Dim j As Long
        For j = 1 To ws.Cells(i, ws.Columns.Count).End(xlToLeft).Column
            If InStr(1, CStr(ws.Cells(i, j).Value), keywordHeader, vbTextCompare) > 0 Then
                FindHeaderRow = i
                Exit Function
            End If
        Next j
    Next i

    FindHeaderRow = 0

End Function

'=============================================================================
' APPLY TEXT TO COLUMNS ON ALL DATA (converts text-stored numbers to real numbers)
'=============================================================================
Private Sub ApplyTextToColumns(ws As Worksheet, lastRow As Long, lastCol As Long)

    ' This forces Excel to re-evaluate each cell as a real value
    ' by copying the value and pasting it back with Paste Special
    Dim rng As Range
    Dim i As Long
    Dim j As Long

    For i = 2 To lastRow
        For j = 1 To lastCol
            On Error Resume Next
            Dim cellVal As Variant
            cellVal = ws.Cells(i, j).Value

            ' Skip empty cells
            If Not IsEmpty(cellVal) Then
                ' Try to convert text to number
                If Not IsNumeric(cellVal) Then
                    ' If it looks like a date, convert it
                    If IsDate(cellVal) Then
                        ws.Cells(i, j).Value = CDate(cellVal)
                    Else
                        ' Try converting to number
                        ws.Cells(i, j).Value = CDbl(cellVal)
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo 0
        Next j
    Next i

End Sub

'=============================================================================
' APPLY TEXT TO COLUMNS ON SPECIFIC COLUMNS ONLY
' Use this for Report 2 tabs where only certain columns need Text to Columns
' Columns are 1-based (A=1, B=2, etc.)
'=============================================================================
Private Sub ApplyTextToColumnsSpecific(ws As Worksheet, lastRow As Long, colArray As Variant)

    Dim c As Variant
    Dim colNum As Long
    Dim rng As Range

    For Each c In colArray
        colNum = CLng(c)

        ' Skip if column is beyond the used range
        If colNum <= ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1 Then
            ' Only process if there's data beyond the header
            If lastRow >= 2 Then
                Set rng = ws.Range(ws.Cells(2, colNum), ws.Cells(lastRow, colNum))

                ' Skip if range is empty — TextToColumns throws "No data was selected to parse"
                If Application.WorksheetFunction.CountA(rng) = 0 Then GoTo NextColumn

                ' Apply Text to Columns — same as manual: Data > Text to Columns
                ' Step 1: Delimited (pre-selected)
                ' Step 2: Tab ✓, Text qualifier: " (double quote)
                ' Step 3: General format
                On Error Resume Next
                rng.TextToColumns _
                    Destination:=ws.Cells(2, colNum), _
                    DataType:=xlDelimited, _
                    Tab:=True, _
                    Semicolon:=False, _
                    Comma:=False, _
                    Space:=False, _
                    Other:=False, _
                    TextQualifier:=xlDoubleQuote, _
                    ConsecutiveDelimiter:=False
                On Error GoTo 0

NextColumn:
                On Error Resume Next
                Set rng = Nothing
                On Error GoTo 0
            End If
        End If
    Next c

End Sub

'=============================================================================
' COPY FILTERED DATA FOR REPORT 2
' Filters and copies rows from source to destination
'=============================================================================
Private Sub CopyFilteredDataReport2(wsSource As Worksheet, wsDest As Worksheet, _
                                     lastRow As Long, lastCol As Long, _
                                     FilterType As String, RunMonthFilter As Boolean)

    ' OPTIMISED: Uses helper-column + AutoFilter approach instead of row-by-row loop.
    ' For BI reports with 1000s of rows this is significantly faster.

    Dim monthPaidCol As Long
    Dim productCol As Long
    Dim helperCol As Long
    Dim i As Long

    ' Copy header row first
    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastCol)).Copy _
        wsDest.Range("A1")

    ' Find Month Paid column
    monthPaidCol = FindColumnByHeader(wsSource, "Month Paid")
    If monthPaidCol = 0 Then
        monthPaidCol = FindColumnByHeader(wsSource, "MonthPaid")
    End If
    If monthPaidCol = 0 Then
        monthPaidCol = FindColumnByHeader(wsSource, "Month_Paid")
    End If

    ' Add a helper column to flag which rows to include
    helperCol = lastCol + 1
    wsSource.Cells(1, helperCol).Value = "_IncludeRow"

    Application.StatusBar = FilterType & ": Evaluating filter criteria..."

    For i = 2 To lastRow
        Dim includeRow As Boolean
        includeRow = False

        Select Case FilterType
            Case "FullLoanSummary"
                ' Full Loan Summary: include rows with Month Paid matching run month (or blank)
                If monthPaidCol > 0 Then
                    Dim monthVal As String
                    monthVal = Trim(CStr(wsSource.Cells(i, monthPaidCol).Value))
                    If Len(monthVal) = 0 Then
                        includeRow = True
                    ElseIf MatchesRunMonth(monthVal) Then
                        includeRow = True
                    End If
                Else
                    includeRow = True
                End If

            Case "VerificationReport"
                ' Verification Report: filter by Month Paid AND Product = NDF or TF
                If monthPaidCol > 0 Then
                    Dim monthValVR As String
                    monthValVR = Trim(CStr(wsSource.Cells(i, monthPaidCol).Value))
                    If Len(monthValVR) > 0 And Not MatchesRunMonth(monthValVR) Then
                        GoTo MarkRow
                    End If
                End If

                ' Check Product column
                productCol = FindColumnByHeader(wsSource, "PRODUCT")
                If productCol > 0 Then
                    Dim productVal As String
                    productVal = Trim(CStr(wsSource.Cells(i, productCol).Value))
                    If InStr(1, productVal, "NDF", vbTextCompare) > 0 Or _
                       InStr(1, productVal, "TF", vbTextCompare) > 0 Then
                        includeRow = True
                    End If
                End If
        End Select

MarkRow:
        If includeRow Then
            wsSource.Cells(i, helperCol).Value = "YES"
        Else
            wsSource.Cells(i, helperCol).Value = "NO"
        End If
    Next i

    ' Apply AutoFilter on helper column to show only "YES" rows
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False

    Dim dataRange As Range
    Set dataRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, helperCol))

    dataRange.AutoFilter Field:=helperCol, Criteria1:="YES"

    ' Copy visible data rows (skip header since we already copied it)
    On Error Resume Next
    Dim visibleRows As Range
    ' Get data rows only (row 2 onwards)
    Set visibleRows = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, lastCol)).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleRows Is Nothing Then
        visibleRows.Copy wsDest.Range("A2")
    End If

    ' Count result rows
    Dim destRows As Long
    destRows = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row - 1
    Application.StatusBar = FilterType & ": Found " & destRows & " matching rows"

    ' Clean up: remove AutoFilter and helper column
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
    wsSource.Columns(helperCol).Delete

End Sub

'=============================================================================
' FORMAT REPORT 2 TAB 1 (Full Loan Summary)
' Currency: Q, Date: S, V
' Delete column U (QA Check Completed)
' Text to Columns: F, J, M, P, Q, S, U (columns 6, 10, 13, 16, 17, 19, 21)
' AutoFit, Freeze, Table
'=============================================================================
Private Sub FormatReport2Tab1(ws As Worksheet)

    Dim lastRow As Long, lastCol As Long
    Dim tbl As ListObject
    Dim tblName As String

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' *** TEXT TO COLUMNS ON SPECIFIC COLUMNS ONLY ***
    ' Tab 1: F, J, M, P, Q, S, U (columns 6, 10, 13, 16, 17, 19, 21)
    Application.StatusBar = "Report 2 Tab 1: Applying Text to Columns on specific columns..."
    Dim tab1Cols As Variant
    tab1Cols = Array(6, 10, 13, 16, 17, 19, 21)
    ApplyTextToColumnsSpecific ws, lastRow, tab1Cols

    ' Format Currency column Q (17)
    If 17 <= lastCol Then
        FormatColumnAsCurrency ws, 17, lastRow
    End If

    ' Format Date columns S (19), V (22)
    If 19 <= lastCol Then
        FormatColumnAsDate ws, 19, lastRow
    End If
    If 22 <= lastCol Then
        FormatColumnAsDate ws, 22, lastRow
    End If

    ' Delete column U (QA Check Completed) — before AutoFit
    Dim qaCol As Long
    qaCol = FindColumnByHeader(ws, "QA")
    If qaCol = 0 Then qaCol = FindColumnByHeader(ws, "QA Check")
    If qaCol > 0 Then
        ws.Columns(qaCol).Delete
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    End If

    ' AutoFit all columns
    ws.Cells.EntireColumn.AutoFit

    ' Set all row heights to 15
    ws.Cells.EntireRow.RowHeight = 15

    ' Freeze top row
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' Convert to Table
    tblName = CleanTableName("Tbl_Sharepoint_Payments")
    On Error Resume Next
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), _
        XlListObjectHasHeaders:=xlYes)

    If Not tbl Is Nothing Then
        tbl.Name = tblName
        tbl.TableStyle = "TableStyleMedium2"
    End If
    On Error GoTo 0

End Sub

'=============================================================================
' FORMAT REPORT 2 TAB 2 (Verification Report)
' Currency: I, N, O, P, Q, R
' Date: K, L, M
' Text to Columns: C, E, F, G, I, K, L, M, N, O, P, Q, R, V, W (columns 3, 5, 6, 7, 9, 11, 12, 13, 14, 15, 16, 17, 18, 22, 23)
' Delete column Y (Verified Flag)
' AutoFit, Freeze, Table
'=============================================================================
Private Sub FormatReport2Tab2(ws As Worksheet)

    Dim lastRow As Long, lastCol As Long
    Dim tbl As ListObject
    Dim tblName As String
    Dim c As Variant

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' *** TEXT TO COLUMNS ON SPECIFIC COLUMNS ONLY ***
    ' Tab 2: C, E, F, G, I, K, L, M, N, O, P, Q, R, V, W (columns 3, 5, 6, 7, 9, 11, 12, 13, 14, 15, 16, 17, 18, 22, 23)
    Application.StatusBar = "Report 2 Tab 2: Applying Text to Columns on specific columns..."
    Dim tab2Cols As Variant
    tab2Cols = Array(3, 5, 6, 7, 9, 11, 12, 13, 14, 15, 16, 17, 18, 22, 23)
    ApplyTextToColumnsSpecific ws, lastRow, tab2Cols

    ' Format Currency columns: I (9), N (14), O (15), P (16), Q (17), R (18)
    Dim currCols As Variant
    currCols = Array(9, 14, 15, 16, 17, 18)

    For Each c In currCols
        If CLng(c) <= lastCol Then
            FormatColumnAsCurrency ws, CLng(c), lastRow
        End If
    Next c

    ' Format Date columns: K (11), L (12), M (13)
    Dim dateCols As Variant
    dateCols = Array(11, 12, 13)

    For Each c In dateCols
        If CLng(c) <= lastCol Then
            FormatColumnAsDate ws, CLng(c), lastRow
        End If
    Next c

    ' Delete column Y (Verified Flag)
    Dim verCol As Long
    verCol = FindColumnByHeader(ws, "Verified")
    If verCol > 0 Then
        ws.Columns(verCol).Delete
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    End If

    ' AutoFit all columns
    ws.Cells.EntireColumn.AutoFit

    ' Set all row heights to 15
    ws.Cells.EntireRow.RowHeight = 15

    ' Freeze top row
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' Convert to Table
    tblName = CleanTableName("Tbl_NDF_TF_Payments")
    On Error Resume Next
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), _
        XlListObjectHasHeaders:=xlYes)

    If Not tbl Is Nothing Then
        tbl.Name = tblName
        tbl.TableStyle = "TableStyleMedium2"
    End If
    On Error GoTo 0

End Sub

'=============================================================================
' REPORT 3: DISTRIBUTION LIST
' Creates one workbook with 2 tabs:
'   Tab 1: "Banker Emails" — all banker emails from 3 MBL tabs (dedupe, no blanks)
'   Tab 2: "BDM Emails" — all BDM emails from 4th MBL tab + Nick Hill + Patrick Harris
' Returns the save path on success, "" on failure
'=============================================================================
Private Function ProcessReport3_DistributionList() As String

    Dim wbMBL As Workbook
    Dim wbNew As Workbook
    Dim wsTab1 As Worksheet
    Dim wsTab2 As Worksheet
    Dim savePath As String
    Dim bankerEmails As Object    ' Scripting.Dictionary for email dedup
    Dim bdmEmails As Object       ' Scripting.Dictionary for email dedup

    ProcessReport3_DistributionList = ""

    On Error GoTo Report3Error

    ' Open MBL workbook
    Application.StatusBar = "Report 3: Opening Master BUID List..."
    Set wbMBL = Workbooks.Open(gMBLPath, Local:=True)

    ' List all sheet names for debugging
    Dim sheetNames As String
    sheetNames = "Sheets in MBL file:" & vbNewLine
    Dim s As Variant
    For Each s In wbMBL.Sheets
        sheetNames = sheetNames & "- " & s.Name & vbNewLine
    Next s
    Debug.Print sheetNames

    ' ---- PART A: Collect Banker Emails from 3 tabs ----
    Application.StatusBar = "Report 3: Collecting banker emails..."

    ' Create dictionaries for email dedup (case-insensitive compare via keys)
    On Error Resume Next
    Set bankerEmails = CreateObject("Scripting.Dictionary")
    Set bdmEmails = CreateObject("Scripting.Dictionary")
    bankerEmails.CompareMode = vbTextCompare
    bdmEmails.CompareMode = vbTextCompare
    On Error GoTo 0

    Dim emailsBeforeTab1 As Long, emailsAfterTab1 As Long
    Dim emailsBeforeTab2 As Long, emailsAfterTab2 As Long
    Dim emailsBeforeTab3 As Long, emailsAfterTab3 As Long

    ' Try to find and process each banker tab - use partial matching
    Dim foundTabs As String
    foundTabs = ""

    ' Tab 1: Try multiple possible names for BDSB
    emailsBeforeTab1 = bankerEmails.Count
    If FindAndCollectEmails(wbMBL, "BDSB", bankerEmails) Then
        emailsAfterTab1 = bankerEmails.Count
        foundTabs = foundTabs & "BDSB: " & (emailsAfterTab1 - emailsBeforeTab1) & " emails" & vbNewLine
    End If

    ' Tab 2: Try multiple possible names for BAB
    emailsBeforeTab2 = bankerEmails.Count
    If FindAndCollectEmails(wbMBL, "BAB", bankerEmails) Then
        emailsAfterTab2 = bankerEmails.Count
        foundTabs = foundTabs & "BAB: " & (emailsAfterTab2 - emailsBeforeTab2) & " emails" & vbNewLine
    End If

    ' Tab 3: Try multiple possible names for SB Health
    emailsBeforeTab3 = bankerEmails.Count
    If FindAndCollectEmails(wbMBL, "SB Health", bankerEmails) Then
        emailsAfterTab3 = bankerEmails.Count
        foundTabs = foundTabs & "SB Health: " & (emailsAfterTab3 - emailsBeforeTab3) & " emails" & vbNewLine
    End If

    ' ---- PART B: Collect BDM Emails from 4th tab ----
    Application.StatusBar = "Report 3: Collecting BDM emails..."

    Dim bdmEmailsBefore As Long, bdmEmailsAfter As Long
    bdmEmailsBefore = bdmEmails.Count

    ' Try to find BDM sheet with partial matching
    If FindAndCollectBDMEmails(wbMBL, "BDM", bdmEmails) Then
        bdmEmailsAfter = bdmEmails.Count
        WriteToLog "BDM Sheet: Found " & (bdmEmailsAfter - bdmEmailsBefore) & " emails"
    End If

    ' Add Nick Hill and Patrick Harris if not already there
    AddEmailIfMissing bdmEmails, "nick.hill@nab.com.au"
    AddEmailIfMissing bdmEmails, "Patrick.T.Harris@nab.com.au"

    WriteToLog "BDM Emails (after adding Nick & Patrick): " & bdmEmails.Count

    ' ---- Create new workbook for both tabs ----
    Set wbNew = Workbooks.Add
    Do While wbNew.Sheets.Count > 1
        wbNew.Sheets(wbNew.Sheets.Count).Delete
    Loop

    ' ---- Tab 1: Banker Emails ----
    Application.StatusBar = "Report 3: Writing Banker Emails tab..."
    Set wsTab1 = wbNew.Sheets(1)
    wsTab1.Name = "Banker Emails"

    WriteEmailsToSheet wsTab1, bankerEmails

    ' ---- Tab 2: BDM Emails ----
    Application.StatusBar = "Report 3: Writing BDM Emails tab..."
    Set wsTab2 = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.Count))
    wsTab2.Name = "BDM Emails"

    WriteEmailsToSheet wsTab2, bdmEmails

    ' ---- Format both tabs ----
    Application.StatusBar = "Report 3: Formatting tabs..."
    FormatDistributionListTab wsTab1
    FormatDistributionListTab wsTab2

    ' ---- Save ----
    savePath = gSaveFolderPath & "Distribution List_Monthly Reminder Email_" & gTodayStr & ".xlsx"

    Application.StatusBar = "Report 3: Saving..."
    wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook

    ' Close MBL
    wbMBL.Close SaveChanges:=False

    ' Log collection results (only shown when DEBUG_MODE = True, to Immediate window)
    WriteToLog "Report 3 Collection Results: Tab1(BDSB)=" & (emailsAfterTab1 - emailsBeforeTab1) & _
               ", Tab2(BAB)=" & (emailsAfterTab2 - emailsBeforeTab2) & _
               ", Tab3(SB Health)=" & (emailsAfterTab3 - emailsBeforeTab3) & _
               ", Total Banker=" & bankerEmails.Count & ", Total BDM=" & bdmEmails.Count

    ProcessReport3_DistributionList = savePath
    Exit Function

Report3Error:
    MsgBox "Error in Report 3: " & Err.Description, vbExclamation, "Report 3 Error"
    ProcessReport3_DistributionList = ""

End Function

'=============================================================================
' COLLECT EMAILS FROM A SHEET (NEW VERSION WITH PARTIAL MATCHING)
' Finds sheet by partial name match, then finds EMAIL column and collects all emails
'=============================================================================
Private Function FindAndCollectEmails(wb As Workbook, partialSheetName As String, emailCollection As Object) As Boolean

    FindAndCollectEmails = False

    Dim ws As Worksheet
    Dim foundWs As Worksheet
    Dim emailCol As Long
    Dim headerRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim emailVal As String
    Dim emailCountBefore As Long
    Dim emailCountAfter As Long

    emailCountBefore = emailCollection.Count

    ' Find the sheet by partial name match
    Set foundWs = Nothing
    For Each ws In wb.Sheets
        If InStr(1, ws.Name, partialSheetName, vbTextCompare) > 0 Then
            Set foundWs = ws
            Exit For
        End If
    Next ws

    If foundWs Is Nothing Then
        Debug.Print "Sheet not found (partial match): " & partialSheetName
        Exit Function
    End If

    Debug.Print "Found sheet: " & foundWs.Name

    ' Find header row by looking for "EMAIL" keyword (like Report 2 does)
    headerRow = FindHeaderRow(foundWs, "EMAIL")
    If headerRow = 0 Then
        ' Fallback: try other common header keywords
        headerRow = FindHeaderRow(foundWs, "BANKER")
        If headerRow = 0 Then
            headerRow = FindHeaderRow(foundWs, "MI")  ' Mortgage Introducer
        End If
    End If
    If headerRow = 0 Then
        headerRow = 1  ' Default to row 1
    End If

    ' Delete rows before the header to make it row 1
    If headerRow > 1 Then
        foundWs.Rows("1:" & headerRow - 1).Delete
        headerRow = 1
    End If

    ' Find EMAIL column (smart detection: EMAIL, EMP EMAIL, EMPLOYEE EMAIL, etc.)
    emailCol = FindColumnByHeader(foundWs, "EMAIL")

    If emailCol = 0 Then
        Debug.Print "No EMAIL column in sheet: " & foundWs.Name
        Exit Function
    End If

    ' Find the actual last row with data
    Dim usedRange As Range
    On Error Resume Next
    Set usedRange = foundWs.UsedRange
    On Error GoTo 0

    If usedRange Is Nothing Then
        lastRow = 1
    Else
        lastRow = usedRange.Rows.Count + usedRange.Row - 1
    End If

    ' Collect all emails from row 2 onwards
    For i = 2 To lastRow
        emailVal = Trim(CStr(foundWs.Cells(i, emailCol).Value))
        If Len(emailVal) > 0 And InStr(emailVal, "@") > 0 Then
            ' Dictionary add — key is the email itself (case-insensitive compare)
            If Not emailCollection.Exists(emailVal) Then
                emailCollection(emailVal) = 1
            End If
        End If
    Next i

    emailCountAfter = emailCollection.Count
    Debug.Print "Collected " & (emailCountAfter - emailCountBefore) & " emails from: " & foundWs.Name

    FindAndCollectEmails = True

End Function

'=============================================================================
' COLLECT BDM EMAILS (WITH PARTIAL MATCHING)
'=============================================================================
Private Function FindAndCollectBDMEmails(wb As Workbook, partialSheetName As String, emailCollection As Object) As Boolean

    FindAndCollectBDMEmails = False

    Dim ws As Worksheet
    Dim foundWs As Worksheet
    Dim emailCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim emailVal As String
    Dim emailCountBefore As Long
    Dim emailCountAfter As Long

    emailCountBefore = emailCollection.Count

    ' Find the sheet by partial name match
    Set foundWs = Nothing
    For Each ws In wb.Sheets
        If InStr(1, ws.Name, partialSheetName, vbTextCompare) > 0 Then
            Set foundWs = ws
            Exit For
        End If
    Next ws

    If foundWs Is Nothing Then
        Debug.Print "BDM sheet not found (partial match): " & partialSheetName
        Exit Function
    End If

    Debug.Print "Found BDM sheet: " & foundWs.Name

    ' Find EMAIL column
    emailCol = FindColumnByHeader(foundWs, "EMAIL")

    If emailCol = 0 Then
        Debug.Print "No EMAIL column in BDM sheet: " & foundWs.Name
        Exit Function
    End If

    ' Find the actual last row with data
    Dim usedRange As Range
    On Error Resume Next
    Set usedRange = foundWs.UsedRange
    On Error GoTo 0

    If usedRange Is Nothing Then
        lastRow = 1
    Else
        lastRow = usedRange.Rows.Count + usedRange.Row - 1
    End If

    ' Collect all emails
    For i = 2 To lastRow
        emailVal = Trim(CStr(foundWs.Cells(i, emailCol).Value))
        If Len(emailVal) > 0 And InStr(emailVal, "@") > 0 Then
            ' Dictionary add — key is the email itself (case-insensitive compare)
            If Not emailCollection.Exists(emailVal) Then
                emailCollection(emailVal) = 1
            End If
        End If
    Next i

    emailCountAfter = emailCollection.Count
    Debug.Print "Collected " & (emailCountAfter - emailCountBefore) & " BDM emails from: " & foundWs.Name

    FindAndCollectBDMEmails = True

End Function

'=============================================================================
' ADD EMAIL IF NOT ALREADY IN DICTIONARY
'=============================================================================
Private Sub AddEmailIfMissing(emailCollection As Object, emailAddress As String)

    ' Dictionary: key = email (lowercase compare via CompareMode), value = 1
    If Not emailCollection.Exists(emailAddress) Then
        emailCollection(emailAddress) = 1
    End If

End Sub

'=============================================================================
' WRITE EMAILS TO SHEET (with deduplication) - NO TABLE
' Uses Scripting.Dictionary (case-insensitive keys) — single pass, no inner loop
'=============================================================================
Private Sub WriteEmailsToSheet(ws As Worksheet, emailCollection As Object)

    Dim i As Long
    Dim row As Long
    Dim emailVal As String
    Dim keys As Variant

    ' Write header
    ws.Cells(1, 1).Value = "EMAIL"
    ws.Range("A1").Font.Bold = True

    row = 2

    ' Dictionary.Keys() returns a 0-based array — single loop, no inner loop
    If emailCollection.Count = 0 Then
        Exit Sub
    End If

    keys = emailCollection.Keys
    For i = LBound(keys) To UBound(keys)
        emailVal = Trim(CStr(keys(i)))

        ' Skip blanks and invalid emails
        If Len(emailVal) = 0 Or InStr(emailVal, "@") = 0 Then
            GoTo NextEmail
        End If

        ws.Cells(row, 1).Value = emailVal
        row = row + 1

NextEmail:
    Next i

End Sub

'=============================================================================
' FORMAT DISTRIBUTION LIST TAB - NO TABLE, JUST PLAIN FORMATTING
'=============================================================================
Private Sub FormatDistributionListTab(ws As Worksheet)

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' AutoFit column width
    ws.Columns("A").AutoFit

    ' Set row height to 15
    ws.Cells.EntireRow.RowHeight = 15

    ' NO table conversion - just plain text with header

End Sub

'=============================================================================
' CREATE REMINDER EMAILS IN OUTLOOK
' Creates 3 draft emails:
'   Email 1: BDMs + HOs
'   Email 2+: BABs in batches of 450
'=============================================================================
Private Sub CreateReminderEmails(report1Path As String, report2Path As String)

    Dim outlookApp As Object
    Dim sharedMailboxRoot As Object
    Dim draftFolder As Object
    Dim fifthBizDay As Date
    Dim fifthBizDayStr As String
    Dim emailSubject As String
    Dim emailBody As String

    On Error GoTo EmailError

    ' ---- Step 1: Calculate 5th business day of next month ----
    Application.StatusBar = "Email: Calculating 5th business day..."
    fifthBizDay = GetFifthBusinessDay(gRunMonthDate)
    fifthBizDayStr = Format(fifthBizDay, "dddd d mmmm yyyy")  ' e.g. "Monday 7 September 2026"

    ' ---- Step 2: Find the nabcommercialbroker shared mailbox ----
    Application.StatusBar = "Email: Connecting to shared mailbox..."

    On Error Resume Next
    Set outlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If outlookApp Is Nothing Then
        MsgBox "Could not connect to Outlook. Please make sure Outlook is open.", vbExclamation
        Exit Sub
    End If

    Set sharedMailboxRoot = FindSharedMailboxRoot(outlookApp)

    If sharedMailboxRoot Is Nothing Then
        MsgBox "Could not find 'nabcommercialbroker' shared mailbox in Outlook." & vbNewLine & _
               "Please ensure the shared mailbox is accessible in your Outlook.", vbExclamation
        Exit Sub
    End If

    ' ---- Step 3: Confirm 5th business day with user ----
    #If DEBUG_MODE Then
        WriteToLog "DEBUG_MODE: Skipping date confirmation, using " & fifthBizDayStr
    #Else
        Dim confirmDate As VbMsgBoxResult
        confirmDate = MsgBox( _
            "5th Business Day of " & gRunMonthName & " (next month):" & vbNewLine & vbNewLine & _
            fifthBizDayStr & vbNewLine & vbNewLine & _
            "Is this correct?" & vbNewLine & vbNewLine & _
            "Click YES to proceed" & vbNewLine & _
            "Click NO to enter a different date", _
            vbYesNo + vbQuestion, "Confirm 5th Business Day")

        If confirmDate = vbNo Then
            Dim customDate As String
            customDate = InputBox("Enter the 5th business day (format: dd/mm/yyyy):", "Enter Custom Date", Format(fifthBizDay, "dd/mm/yyyy"))
            If customDate <> "" Then
                On Error Resume Next
                fifthBizDay = CDate(customDate)
                fifthBizDayStr = Format(fifthBizDay, "dddd d mmmm yyyy")
                On Error GoTo 0
            End If
        End If
    #End If

    ' ---- Step 4: Build email content ----
    Application.StatusBar = "Email: Building email content..."
    emailSubject = "***Reminder*** NAB Commercial Broker Commissions Payments - Action Required by " & fifthBizDayStr

    ' Hardcoded email body template — replaces the need to find a template email
    emailBody = GetHardcodedEmailTemplate()

    ' Then update it with current run details (dates, months, etc.)
    ' Pass the raw fifthBizDay Date variable to guarantee it overrides the template
    emailBody = UpdateEmailTemplate(emailBody, fifthBizDay)

    ' ---- Step 6: Create email drafts ----
    Application.StatusBar = "Email: Creating draft emails..."
    Set draftFolder = GetDraftFolder(sharedMailboxRoot)

    If draftFolder Is Nothing Then
        MsgBox "Could not access Drafts folder in nabcommercialbroker shared mailbox.", vbExclamation
        Exit Sub
    End If

    ' Email 1: BDMs + HOs
    CreateBDMEmailDraft outlookApp, draftFolder, emailSubject, emailBody, report1Path, report2Path

    ' Email 2+: BABs in batches of 450
    CreateBankerEmailDrafts outlookApp, draftFolder, emailSubject, emailBody, report1Path, report2Path

    ' Show success popup using HTML
    ShowSuccessPopup

    Exit Sub

EmailError:
    MsgBox "Error creating emails: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Line: " & Err.Source, vbExclamation, "Email Error"

End Sub

'=============================================================================
' SHOW SUCCESS POPUP (HTML-based, with emoji + formatted text)
' Creates a temp HTML file and opens it in the default browser
'=============================================================================
Private Sub ShowSuccessPopup()

    Dim html As String
    Dim htmlHead As String
    Dim htmlBody As String
    Dim fso As Object
    Dim ts As Object
    Dim tempFile As String

    ' Build HTML in parts to avoid VBA's 25-line-continuation limit

    ' Part 1: Head + CSS
    htmlHead = "<!DOCTYPE html><html><head><meta charset='UTF-8'>"
    htmlHead = htmlHead & "<title>Commission Run Success</title><style>"
    htmlHead = htmlHead & "body{font-family:'Segoe UI',Tahoma,sans-serif;background:linear-gradient(135deg,#667eea,#764ba2);"
    htmlHead = htmlHead & "min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;margin:0;}"
    htmlHead = htmlHead & ".card{background:#fff;border-radius:20px;padding:45px 55px;box-shadow:0 20px 60px rgba(0,0,0,.3);"
    htmlHead = htmlHead & "text-align:center;max-width:520px;animation:slideUp .5s ease-out;}"
    htmlHead = htmlHead & "@keyframes slideUp{from{opacity:0;transform:translateY(30px)}to{opacity:1;transform:translateY(0)}}"
    htmlHead = htmlHead & ".emoji{font-size:80px;display:block;margin-bottom:15px;}"
    htmlHead = htmlHead & "h1{font-size:30px;font-weight:800;color:#2d3748;margin:0 0 15px;font-style:italic;}"
    htmlHead = htmlHead & ".highlight{font-size:26px;font-weight:800;font-style:italic;color:#667eea;margin:18px 0;}"
    htmlHead = htmlHead & ".msg{font-size:18px;color:#4a5568;line-height:1.6;margin-bottom:22px;}"
    htmlHead = htmlHead & ".info{background:#f7fafc;border-left:4px solid #667eea;padding:15px 20px;margin:20px 0;text-align:left;border-radius:5px;font-size:15px;color:#4a5568;}"
    htmlHead = htmlHead & ".info p{margin:4px 0;}"
    htmlHead = htmlHead & ".btn{background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;border:none;padding:12px 40px;"
    htmlHead = htmlHead & "font-size:16px;font-weight:600;border-radius:50px;cursor:pointer;margin-top:20px;}"
    htmlHead = htmlHead & ".btn:hover{transform:translateY(-2px);box-shadow:0 5px 20px rgba(102,126,234,.4);}"
    htmlHead = htmlHead & "</style></head><body>"

    ' Part 2: Card content
    htmlBody = "<div class='card'>"
    htmlBody = htmlBody & "<span class='emoji'>&#127881;</span>"
    htmlBody = htmlBody & "<h1>Commission Run Complete!</h1>"
    htmlBody = htmlBody & "<p class='highlight'>&#9989; Yeahhh! The reminder mail is done!</p>"
    htmlBody = htmlBody & "<p class='msg'>Email drafts have been created successfully in the "
    htmlBody = htmlBody & "<strong>nabcommercialbroker@nab.com.au</strong> Drafts folder!</p>"
    htmlBody = htmlBody & "<div class='info'><p><strong>Next Steps:</strong></p>"
    htmlBody = htmlBody & "<p>&bull; Open Outlook</p>"
    htmlBody = htmlBody & "<p>&bull; Go to the nabcommercialbroker Drafts folder</p>"
    htmlBody = htmlBody & "<p>&bull; Review and send the reminder emails</p></div>"
    htmlBody = htmlBody & "<button class='btn' onclick='window.close()'>Close</button>"
    htmlBody = htmlBody & "</div></body></html>"

    ' Combine parts
    html = htmlHead & htmlBody

    ' Write to temp file and open in browser
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFile = Environ("TEMP") & "\commission_success.html"
    Set ts = fso.CreateTextFile(tempFile, True)
    ts.Write html
    ts.Close
    On Error GoTo 0

    Shell "cmd /c start """" """ & tempFile & """", vbNormalFocus

    Set fso = Nothing
    Set ts = Nothing

End Sub

'=============================================================================
' CALCULATE 5TH BUSINESS DAY OF NEXT MONTH
'=============================================================================
Private Function GetFifthBusinessDay(runMonthDate As Date) As Date

    Dim nextMonth As Date
    Dim businessDayCount As Long
    Dim currentDay As Date
    Dim dayOfWeek As Long

    ' Get first day of next month
    nextMonth = DateSerial(Year(runMonthDate), Month(runMonthDate) + 1, 1)

    ' Handle year rollover
    If Month(runMonthDate) = 12 Then
        nextMonth = DateSerial(Year(runMonthDate) + 1, 1, 1)
    End If

    currentDay = nextMonth
    businessDayCount = 0

    ' Count weekdays (vbSunday=1, vbMonday=2, ..., vbSaturday=7)
    ' Business days are Mon-Fri (vbMonday to vbFriday)
    Do While businessDayCount < 5
        dayOfWeek = VBA.Weekday(currentDay, vbSunday)
        If dayOfWeek >= vbMonday And dayOfWeek <= vbFriday Then
            businessDayCount = businessDayCount + 1
            If businessDayCount = 5 Then
                GetFifthBusinessDay = currentDay
                Exit Function
            End If
        End If
        currentDay = currentDay + 1
    Loop

End Function

'=============================================================================
' FIND NABCOMMERCIALBROKER SHARED MAILBOX ROOT FOLDER
' Returns the root folder of the shared mailbox (which has Inbox, Sent Items, Drafts, etc.)
'=============================================================================
Private Function FindSharedMailboxRoot(outlookApp As Object) As Object

    Dim namespace As Object
    Dim store As Object
    Dim folder As Object
    Dim i As Long

    On Error GoTo FindError

    Set namespace = outlookApp.GetNamespace("MAPI")

    ' Method 1: Search namespace.Folders (top-level folder entries in Outlook sidebar)
    ' Shared mailboxes added via "Open these additional mailboxes" appear here
    For Each folder In namespace.Folders
        If InStr(1, folder.Name, "nabcommercialbroker", vbTextCompare) > 0 Then
            Set FindSharedMailboxRoot = folder
            Set gSharedMailboxRoot = folder  ' Cache it globally
            Exit Function
        End If
    Next folder

    ' Method 2: Search namespace.Stores (if Folders didn't work)
    For Each store In namespace.Stores
        If InStr(1, store.DisplayName, "nabcommercialbroker", vbTextCompare) > 0 Then
            Set FindSharedMailboxRoot = store.GetRootFolder
            Set gSharedMailboxRoot = store.GetRootFolder  ' Cache it globally
            Exit Function
        End If
    Next store

    Set FindSharedMailboxRoot = Nothing
    Exit Function

FindError:
    Debug.Print "FindSharedMailboxRoot error: " & Err.Description
    Set FindSharedMailboxRoot = Nothing

End Function

'=============================================================================
' FIND TEMPLATE EMAIL IN SHARED MAILBOX SENT ITEMS
'=============================================================================
Private Function FindTemplateEmail(sharedMailboxRoot As Object) As Object

    Dim sentFolder As Object
    Dim items As Object
    Dim item As Object
    Dim i As Long

    On Error GoTo FindError

    If sharedMailboxRoot Is Nothing Then
        Set FindTemplateEmail = Nothing
        Exit Function
    End If

    ' Get Sent Items folder directly from shared mailbox root
    On Error Resume Next
    Set sentFolder = sharedMailboxRoot.Folders("Sent Items")
    On Error GoTo 0

    If sentFolder Is Nothing Then
        Set FindTemplateEmail = Nothing
        Exit Function
    End If

    Set items = sentFolder.Items
    On Error Resume Next
    items.Sort "[ReceivedTime]", False  ' Sort by date, newest first
    On Error GoTo 0

    ' Search for email with "Reminder" and "NAB Commercial Broker" and "Commissions"
    For i = 1 To items.Count
        Set item = items(i)
        If InStr(1, item.Subject, "***Reminder***", vbTextCompare) > 0 And _
           InStr(1, item.Subject, "NAB Commercial Broker", vbTextCompare) > 0 And _
           InStr(1, item.Subject, "Commissions", vbTextCompare) > 0 Then
            Set FindTemplateEmail = item
            Exit Function
        End If
    Next i

    Set FindTemplateEmail = Nothing
    Exit Function

FindError:
    Debug.Print "FindTemplateEmail error: " & Err.Description
    Set FindTemplateEmail = Nothing

End Function

'=============================================================================
' SELECT EMAIL MANUALLY VIA OUTLOOK FOLDER BROWSER
'=============================================================================
Private Function SelectEmailManually(outlookApp As Object) As Object

    MsgBox "Please manually select the template email from your Outlook sent items." & vbNewLine & _
           "Open Outlook and select an email, then click OK here.", vbInformation

    ' This is a placeholder — in real usage, user would select from Outlook manually
    ' For now, we'll try to find it again or return Nothing
    Set SelectEmailManually = Nothing

End Function

'=============================================================================
' HARDCODED EMAIL BODY TEMPLATE
' Returns the email body HTML with placeholder dates that UpdateEmailTemplate
' will replace with the correct dates for the current run month.
'=============================================================================
Private Function GetHardcodedEmailTemplate() As String

    ' Hardcoded fallback template (all 3 verified hyperlinks included).
    GetHardcodedEmailTemplate = GetFallbackTemplateBody()

End Function


' ---- Email template body: exact reconstruction from all 5 body photos ----
' Built verbatim from the WhatsApp body screenshots (03-Sep-2026).
' Placeholders that UpdateEmailTemplate replaces each run:
'   "August 2026"  -> next run month   |  "01.09.2026" -> gTodayStr
'   "7th of September 2026" / "7th September 2026" (with or without "of"/year)
'                    -> 5th business day ordinal in next month
'   "Friday"       -> actual weekday of that 5th business day
' Built with `body = body & "..."` (NO " _" continuations) so the 25-limit never triggers.
Private Function GetFallbackTemplateBody() As String

    Dim body As String

    body = "<font face='Calibri' size='3' color='#000000'>"
    body = body & "<p>Hi All,</p>"
    body = body & "<p>Thank you for your assistance in helping us to ensure your broker is paid for the great deals you are writing.</p>"
    body = body & "<p><b><u>NCB- August 2026 Payments 01.09.2026</u></b></p>"
    body = body & "<p>This spreadsheet shows the commercial loans that we have processed for August 2026 commission run however is pending for 4 Eye Check. If we come across any errors, we will contact you via email.</p>"
    body = body & "<p><b><u>NCB Payment SharePoint 01.09.2026</u></b></p>"
    body = body & "<p>This spreadsheet shows the commercial loans that have been submitted to our team for processing.</p>"
    body = body & "<p><b>Tab 1 &ndash; Payments yet to be reviewed.</b><br>"
    body = body & "&#10004; If you see the Payment Status of <b>&quot;New&quot;</b> or <b>&quot;Commission Allocated&quot;</b> recorded against your submission &ndash; this is yet to be reviewed. The team will review your submission and contact you via email if there are any errors which need to be rectified</p>"
    body = body & "<p><b>Tab 2 &ndash; Payments returned</b><br>"
    body = body & "&#10004; If you see the Payment Status of <b>&quot;Returned &ndash; Error&quot;</b> against your submission &ndash; this submission has been reviewed by the team and returned to you. You need to follow the instructions within the email received and rectify the errors identified</p>"
    body = body & "<p>If there are other commercial loans which were drawn in <b>August 2026</b> that <b><u>do not appear on these reports</u></b>, then please complete the "
    body = body & "<a href='https://nabcts.sharepoint.com/sites/s07_site2121/Lists/Commercial%20Broker%20Commission%20Payment%20Form/NewForm.aspx?Source=https%3a%2f%2fnabcts.sharepoint.com%2fsites%2fs07_site2121%2fLists%2fCommercial%2520Broker%2520Commission%2520Payment%2520Form%2fAllItems.aspx' style='color:#0563C1;text-decoration:underline;'><u>New item | Commercial Broker Commission Payment Form (sharepoint.com)</u></a>"
    body = body & " to enable your broker to receive commission.</p>"
    body = body & "<p>Here is a &#128203; <a href='https://nabcts.sharepoint.com/sites/intraGS115/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FintraGS115%2FShared%20Documents%2FNAB%20Commercial%20Broker%20Intranet%20site%2FHow%20To%20Guide%20-%20Commission%20Payment%20Process.pdf&parent=%2Fsites%2FintraGS115%2FShared%20Documents%2FNAB%20Commercial%20Broker%20Intranet%20site' style='color:#0563C1;text-decoration:underline;'><u>How To Guide - Commission Payment Process.pdf</u></a> if you need any assistance with completing the process</p>"
    body = body & "<ul>"
    body = body & "<li><b><u>Reminder</u></b> the cut-off date for submitting your Commission Payment Form via SharePoint for payments in the <b>August 2026</b> commission run is <b><u>COB Friday, 7th of September 2026</u></b>.</li>"
    body = body & "<li>Any Commission Payment Forms submitted after this date will be processed in the following commission run.</li>"
    body = body & "<li>If there are any issues with your Commission Payment Form, we will contact you via email advising what errors need to be rectified. <b><u>Please ensure you are checking your emails before COB Friday, 7th of September 2026</u></b> as errors must be resolved before the cut-off date, unless otherwise advised by a commission team member.</li>"
    body = body & "<li>Our team will review all completed requests received in the queue by <b><u>COB Friday, 7th of September 2026</u></b>.</li>"
    body = body & "</ul>"
    body = body & "<p>Any questions please reach out to the team inbox <a href='mailto:nabcommercialbroker@nab.com.au' style='color:#0563C1;text-decoration:underline;'><u>nabcommercialbroker@nab.com.au</u></a></p>"
    body = body & "<p>Thanks &amp; Regards,</p>"
    body = body & "</font>"

    GetFallbackTemplateBody = body

End Function


'=============================================================================
' UPDATE EMAIL TEMPLATE WITH CURRENT RUN DETAILS
'=============================================================================
Private Function UpdateEmailTemplate(htmlBody As String, fifthBizDay As Date) As String

    Dim updatedBody As String
    Dim nextMonthName As String
    Dim nextMonthDate As Date
    Dim fifthBizDayOrdinal As String
    Dim nextMonthNameOnly As String

    updatedBody = htmlBody

    ' Calculate next month for replacement
    nextMonthDate = DateSerial(Year(gRunMonthDate), Month(gRunMonthDate) + 1, 1)
    If Month(gRunMonthDate) = 12 Then
        nextMonthDate = DateSerial(Year(gRunMonthDate) + 1, 1, 1)
    End If

    nextMonthName = Format(nextMonthDate, "MMMM yyyy")
    nextMonthNameOnly = Format(nextMonthDate, "MMMM")

    ' Get the day name (Monday, Tuesday, etc.) for the 5th business day
    Dim fifthBizDayName As String
    fifthBizDayName = Format(fifthBizDay, "dddd")

    ' Build the full date string with ordinal suffix
    fifthBizDayOrdinal = Format(Day(fifthBizDay), "0") & GetOrdinalSuffix(Day(fifthBizDay)) & " of " & Format(fifthBizDay, "MMMM yyyy")

    ' --- REPLACEMENT 1: Replace filename dates (template date → today) ---
    updatedBody = Replace(updatedBody, "01.09.2026", gTodayStr, , -1)

    ' --- REPLACEMENT 2: Replace deadline date phrases with computed next month + 5th business day ---
    ' IMPORTANT: Deadline replacements MUST happen BEFORE body month replacement.
    ' Template hardcodes "September 2026" as deadline month (next month after August run).
    ' These replacements use the COMPUTED next month from the actual run month.
    updatedBody = Replace(updatedBody, "COB Friday, 7th of September 2026", "COB " & fifthBizDayName & ", " & fifthBizDayOrdinal, , -1)
    updatedBody = Replace(updatedBody, "COB Friday, 7th September 2026", "COB " & fifthBizDayName & ", " & Format(Day(fifthBizDay), "0") & GetOrdinalSuffix(Day(fifthBizDay)) & " " & nextMonthNameOnly & " " & CStr(Year(fifthBizDay)), , -1)
    updatedBody = Replace(updatedBody, "Friday, 7th of September 2026", fifthBizDayName & ", " & fifthBizDayOrdinal, , -1)
    updatedBody = Replace(updatedBody, "Friday, 7th September 2026", fifthBizDayName & ", " & Format(Day(fifthBizDay), "0") & GetOrdinalSuffix(Day(fifthBizDay)) & " " & nextMonthNameOnly & " " & CStr(Year(fifthBizDay)), , -1)
    updatedBody = Replace(updatedBody, "7th of September 2026", Format(Day(fifthBizDay), "0") & GetOrdinalSuffix(Day(fifthBizDay)) & " of " & Format(fifthBizDay, "MMMM yyyy"), , -1)
    updatedBody = Replace(updatedBody, "7th September 2026", Format(Day(fifthBizDay), "0") & GetOrdinalSuffix(Day(fifthBizDay)) & " " & nextMonthNameOnly & " " & CStr(Year(fifthBizDay)), , -1)

    ' --- REPLACEMENT 3: Replace template month with run month in body text ---
    ' Template hardcodes "August 2026" — replace with the user's actual run month.
    ' This does NOT catch deadline dates because they contain "September 2026"
    ' (already replaced above to computed next month).
    updatedBody = Replace(updatedBody, "for August 2026 commission run", "for " & gRunMonthName & " " & CStr(gRunMonthYear) & " commission run", , -1)
    updatedBody = Replace(updatedBody, "in August 2026 that do", "in " & gRunMonthName & " " & CStr(gRunMonthYear) & " that do", , -1)
    updatedBody = Replace(updatedBody, "in the August 2026 commission run is", "in the " & gRunMonthName & " " & CStr(gRunMonthYear) & " commission run is", , -1)
    updatedBody = Replace(updatedBody, "August 2026 Payments", gRunMonthName & " " & CStr(gRunMonthYear) & " Payments", , -1)

    UpdateEmailTemplate = updatedBody

End Function

' Helper function to get ordinal suffix (st, nd, rd, th)
Private Function GetOrdinalSuffix(dayNum As Long) As String
    Select Case dayNum Mod 10
        Case 1
            If dayNum Mod 100 <> 11 Then
                GetOrdinalSuffix = "st"
            Else
                GetOrdinalSuffix = "th"
            End If
        Case 2
            If dayNum Mod 100 <> 12 Then
                GetOrdinalSuffix = "nd"
            Else
                GetOrdinalSuffix = "th"
            End If
        Case 3
            If dayNum Mod 100 <> 13 Then
                GetOrdinalSuffix = "rd"
            Else
                GetOrdinalSuffix = "th"
            End If
        Case Else
            GetOrdinalSuffix = "th"
    End Select
End Function

'=============================================================================
' GET DRAFT FOLDER FROM SHARED MAILBOX
'=============================================================================
Private Function GetDraftFolder(sharedMailboxRoot As Object) As Object

    Dim draftFolder As Object

    On Error Resume Next
    If Not sharedMailboxRoot Is Nothing Then
        Set draftFolder = sharedMailboxRoot.Folders("Drafts")
    End If
    On Error GoTo 0

    Set GetDraftFolder = draftFolder

End Function

'=============================================================================
' CREATE BDM EMAIL DRAFT
' Reads BDM emails from the Distribution List file's "BDM Emails" tab
'=============================================================================
Private Sub CreateBDMEmailDraft(outlookApp As Object, draftFolder As Object, _
                                subject As String, body As String, _
                                report1Path As String, report2Path As String)

    Dim mailItem As Object
    Dim hoEmails() As String
    Dim i As Long
    Dim recip As Object
    Dim wbDist As Workbook
    Dim wsBDM As Worksheet
    Dim lastRow As Long
    Dim bdmEmailList As String
    Dim emailVal As String

    On Error GoTo EmailError

    ' HOs list
    hoEmails = Split("Anita.Lindsay@nab.com.au;John.A.Shillington@nab.com.au;Peter.A.Bugler@nab.com.au;Sam.Turri@nab.com.au;nabcommercialbroker@nab.com.au", ";")

    ' ---- Read BDM emails from Distribution List file ----
    Dim distListPath As String
    distListPath = gSaveFolderPath & "Distribution List_Monthly Reminder Email_" & gTodayStr & ".xlsx"

    On Error Resume Next
    Set wbDist = Workbooks.Open(distListPath, Local:=True)
    On Error GoTo 0

    If wbDist Is Nothing Then
        MsgBox "Could not open Distribution List file for BDM emails.", vbExclamation
        Exit Sub
    End If

    Set wsBDM = wbDist.Sheets("BDM Emails")
    lastRow = wsBDM.Cells(wsBDM.Rows.Count, "A").End(xlUp).Row

    bdmEmailList = ""
    For i = 2 To lastRow
        emailVal = Trim(CStr(wsBDM.Cells(i, 1).Value))
        If Len(emailVal) > 0 And InStr(emailVal, "@") > 0 Then
            If Len(bdmEmailList) > 0 Then
                bdmEmailList = bdmEmailList & ";"
            End If
            bdmEmailList = bdmEmailList & emailVal
        End If
    Next i

    wbDist.Close SaveChanges:=False

    If Len(bdmEmailList) = 0 Then
        MsgBox "No BDM emails found in Distribution List. BDM email not created.", vbExclamation
        Exit Sub
    End If

    ' ---- Create the BDM email draft ----
    Set mailItem = outlookApp.CreateItem(0)  ' 0 = olMailItem

    With mailItem
        .Subject = subject
        .BodyFormat = 3  ' olFormatHTML

        ' Add the reminder body, then append the user's Outlook signature
        .HTMLBody = body & GetOutlookSignature()

        ' Add BDMs to TO field (semicolons work as separators in .To)
        .To = bdmEmailList

        ' Add HOs to CC field
        For i = LBound(hoEmails) To UBound(hoEmails)
            Set recip = .Recipients.Add(Trim(hoEmails(i)))
            recip.Type = 2  ' olCC
        Next i

        .Attachments.Add report1Path
        .Attachments.Add report2Path
        .Move draftFolder
    End With

    Set mailItem = Nothing
    Exit Sub

EmailError:
    MsgBox "Error creating BDM email: " & Err.Description & " (Line: " & Err.Number & ")", vbExclamation
    Set mailItem = Nothing

End Sub

'=============================================================================
' CREATE BANKER EMAIL DRAFTS (BATCHED AT 450)
'=============================================================================
Private Sub CreateBankerEmailDrafts(outlookApp As Object, draftFolder As Object, _
                                     subject As String, body As String, _
                                     report1Path As String, report2Path As String)

    Dim wbDist As Workbook
    Dim wsTab1 As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim emailList As String
    Dim batchCount As Long
    Dim batchNum As Long
    Dim mailItem As Object
    Dim distListPath As String

    ' Build the Distribution List path
    distListPath = gSaveFolderPath & "Distribution List_Monthly Reminder Email_" & gTodayStr & ".xlsx"

    ' Open the Distribution List workbook
    On Error Resume Next
    Set wbDist = Workbooks.Open(distListPath, Local:=True)
    On Error GoTo 0

    If wbDist Is Nothing Then
        MsgBox "Could not open Distribution List file. Banker emails not created.", vbExclamation
        Exit Sub
    End If

    Set wsTab1 = wbDist.Sheets("Banker Emails")

    lastRow = wsTab1.Cells(wsTab1.Rows.Count, "A").End(xlUp).Row

    ' Create email drafts in batches of 450
    batchNum = 1
    batchCount = 0
    emailList = ""

    For i = 2 To lastRow
        Dim emailVal As String
        emailVal = Trim(CStr(wsTab1.Cells(i, 1).Value))

        If Len(emailVal) > 0 And InStr(emailVal, "@") > 0 Then
            If Len(emailList) > 0 Then
                emailList = emailList & ";"
            End If
            emailList = emailList & emailVal
            batchCount = batchCount + 1

            ' Create draft when batch reaches 450 or at end of list
            If batchCount >= 450 Or i = lastRow Then
                Set mailItem = outlookApp.CreateItem(0)  ' 0 = olMailItem

                With mailItem
                    .Subject = subject
                    .BodyFormat = 3  ' olFormatHTML

                    ' Add the reminder body, then append the user's Outlook signature
                    .HTMLBody = body & GetOutlookSignature()

                    .To = "nabcommercialbroker@nab.com.au"
                    .BCC = emailList
                    .Attachments.Add report1Path
                    .Attachments.Add report2Path
                    .Move draftFolder
                End With

                Set mailItem = Nothing

                ' Reset for next batch
                emailList = ""
                batchCount = 0
                batchNum = batchNum + 1
            End If
        End If
    Next i

    wbDist.Close SaveChanges:=False

End Sub

'=============================================================================
' GET OUTLOOK SIGNATURE
' Reads the default Outlook signature HTML from the user's Signatures folder
'=============================================================================
Private Function GetOutlookSignature() As String

    ' Reads the user's DEFAULT "new mail" Outlook signature from the filesystem.
    ' Strategy (safe — no .Display, no Word-generated CSS, no "invalid style" errors):
    '   1) Return cached value instantly if already fetched in this run.
    '   2) Read the default signature NAME from the registry (Office 16.0 / 15.0).
    '   3) Open ONLY the matching .htm file in %APPDATA%\Microsoft\Signatures\.
    '      No fallback scanning — avoids picking up template/other signatures.
    '   4) Extract the inner <body> content (strips <html>/<head>/<style> wrappers).
    '   5) Convert relative image paths (e.g. "SignatureName_files/image001.png")
    '      to absolute file:/// URIs so Outlook renders logos, photos, GIFs etc.
    '   6) Return clean HTML ready to concatenate after the email template body.
    '   7) If nothing found at all, return "" — email still works, just no signature.

    Static cachedSignature As String
    Static hasChecked As Boolean

    If hasChecked Then
        GetOutlookSignature = cachedSignature
        Exit Function
    End If

    hasChecked = True
    cachedSignature = ""

    On Error Resume Next

    Dim sigFolder As String
    Dim sigFile As String
    Dim sigName As String
    Dim fso As Object
    Dim ts As Object
    Dim wshShell As Object
    Dim rawHtml As String

    sigFolder = Environ("APPDATA") & "\Microsoft\Signatures\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sigFolder) Then GoTo SigDone

    ' --- Step 1: Read the default "new message" signature name from registry ---
    '     This is the ONLY signature we want — prevents pulling templates/extras.
    Set wshShell = CreateObject("WScript.Shell")

    ' Office 365 / 2019 / 2016 stores it under 16.0
    sigName = ""
    sigName = wshShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\MailSettings\NewSignature")

    ' If 16.0 didn't work, try 15.0 (Office 2013)
    If Len(sigName) = 0 Then
        sigName = wshShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Common\MailSettings\NewSignature")
    End If

    Err.Clear

    ' --- Step 2: Find the EXACT .htm file for that signature ---
    '     NO fallback to index.htm or first-found .htm — that's what caused
    '     multiple signatures / wrong signature being pulled.
    If Len(sigName) = 0 Then GoTo SigDone
    If Not fso.FileExists(sigFolder & sigName & ".htm") Then GoTo SigDone
    sigFile = sigFolder & sigName & ".htm"

    ' --- Step 3: Read the file and extract just the <body> content ---
    Set ts = fso.OpenTextFile(sigFile, 1, False, 0)   ' 0 = TristateFalse (ANSI / system default)
    rawHtml = ts.ReadAll
    ts.Close
    Set ts = Nothing

    Dim bodyStart As Long
    Dim bodyEnd As Long

    bodyStart = InStr(1, rawHtml, "<body", vbTextCompare)
    If bodyStart > 0 Then
        bodyStart = InStr(bodyStart, rawHtml, ">", vbBinaryCompare)
        If bodyStart > 0 Then bodyStart = bodyStart + 1
    End If

    bodyEnd = InStrRev(rawHtml, "</body>", -1, vbTextCompare)

    If bodyStart > 0 And bodyEnd > bodyStart Then
        cachedSignature = Mid$(rawHtml, bodyStart, bodyEnd - bodyStart)
    ElseIf bodyStart > 0 Then
        cachedSignature = Mid$(rawHtml, bodyStart)
    Else
        cachedSignature = rawHtml
        Dim headEnd As Long
        headEnd = InStr(1, cachedSignature, "</head>", vbTextCompare)
        If headEnd > 0 Then
            cachedSignature = Mid$(cachedSignature, headEnd + 7)
        End If
    End If

    cachedSignature = Trim$(cachedSignature)

    ' Clean up any stray closing tags
    If LCase$(Right$(cachedSignature, 7)) = "</html>" Then
        cachedSignature = Left$(cachedSignature, Len(cachedSignature) - 7)
    End If

    ' --- Step 4: Fix image paths so logos/photos/GIFs render in Outlook ---
    '     Outlook signature .htm files reference images with relative paths like:
    '       src="SignatureName_files/image001.png"
    '     These break when pasted into an email's HTMLBody. Convert them to
    '     absolute file:/// URIs so Outlook can find and embed them.
    Dim sigFilesFolder As String
    sigFilesFolder = sigName & "_files/"

    If InStr(1, cachedSignature, sigFilesFolder, vbTextCompare) > 0 Then
        ' Convert relative paths to absolute file:/// URIs
        Dim absPath As String
        absPath = "file:///" & Replace(sigFolder, "\", "/") & sigName & "_files/"
        cachedSignature = Replace(cachedSignature, """" & sigFilesFolder, """" & absPath, , -1, vbTextCompare)
        cachedSignature = Replace(cachedSignature, "'" & sigFilesFolder, "'" & absPath, , -1, vbTextCompare)
    End If

SigDone:
    On Error GoTo 0
    GetOutlookSignature = cachedSignature

    Set ts = Nothing
    Set fso = Nothing
    Set wshShell = Nothing

End Function

'=============================================================================
' LEGACY ENTRY POINT - Report 1 only (standalone)
' Kept for backward compatibility — runs Report 1 by itself
'=============================================================================
Public Sub RunReport1_SharePoint()

    Dim savePath As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Get run month
    If Not GetRunMonthInput() Then
        MsgBox "Operation cancelled.", vbInformation
        GoTo CleanUp
    End If

    ' Ask for SharePoint CSV
    gSharePointCSVPath = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        Title:="Select the downloaded SharePoint CSV file")

    If gSharePointCSVPath = "False" Or gSharePointCSVPath = "" Then
        MsgBox "No file selected. Operation cancelled.", vbInformation
        GoTo CleanUp
    End If

    ' Process
    savePath = ProcessReport1_SharePoint()

    If savePath <> "" Then
        MsgBox "Report 1 completed successfully!" & vbNewLine & vbNewLine & _
               "Saved to:" & vbNewLine & savePath, vbInformation, "Success"
    End If

CleanUp:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

'=============================================================================
' GET RUN MONTH INPUT FROM USER
'=============================================================================
Private Function GetRunMonthInput() As Boolean

    Dim userInput As String
    Dim testDate As Date

    GetRunMonthInput = False

    ' Ask for run month
    userInput = InputBox( _
        "Enter the Commission Run Month and Year:" & vbNewLine & vbNewLine & _
        "Examples: Aug 2026, August 2026, 08/2026" & vbNewLine & vbNewLine & _
        "This will be used to filter Month Paid and determine the save folder.", _
        "Commission Run Month")

    If userInput = "" Then Exit Function

    ' Try to parse the input into a date
    On Error Resume Next

    ' Try direct parse first (e.g., "August 2026" or "Aug 2026")
    testDate = CDate("1 " & userInput)

    If Err.Number <> 0 Then
        Err.Clear
        ' Try "MM/YYYY" format
        If InStr(userInput, "/") > 0 Then
            Dim parts() As String
            parts = Split(userInput, "/")
            If UBound(parts) = 1 Then
                testDate = DateSerial(CLng(parts(1)), CLng(parts(0)), 1)
            End If
        End If
    End If

    If Err.Number <> 0 Then
        MsgBox "Could not understand '" & userInput & "'." & vbNewLine & _
               "Please use format like: August 2026, Aug 2026, or 08/2026", _
               vbExclamation
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    ' Set global variables
    gRunMonthDate = DateSerial(Year(testDate), Month(testDate), 1)
    gRunMonthName = Format(gRunMonthDate, "MMMM")
    gRunMonthYear = Year(gRunMonthDate)
    gRunMonthMM = Format(gRunMonthDate, "MM")
    gTodayStr = Format(Date, "DD.MM.YYYY")

    ' Calculate Fiscal Year (Oct-Sep: Oct onwards = next year's FY)
    If Month(gRunMonthDate) >= 10 Then
        gFiscalYear = "FY" & Right(CStr(gRunMonthYear + 1), 2)
    Else
        gFiscalYear = "FY" & Right(CStr(gRunMonthYear), 2)
    End If

    ' Build save folder path
    gSaveFolderPath = gSharedDriveLetter & ":\NAB COMMERCIAL BROKER\Commission System\" & _
                      gFiscalYear & "\" & _
                      CStr(gRunMonthYear) & " - " & gRunMonthMM & " NCB Monthly Commission\"

    ' Verify folder exists and confirm with user
    If Dir(gSaveFolderPath, vbDirectory) = "" Then
        ' Folder not found — let user browse
        MsgBox "Default save folder not found:" & vbNewLine & gSaveFolderPath & vbNewLine & vbNewLine & _
               "You'll need to select a folder manually.", vbInformation, "Save Location"

        Dim fallbackFolder As String
        fallbackFolder = BrowseForFolder("Select the folder to save reports to:")

        If fallbackFolder = "" Then Exit Function
        If Right(fallbackFolder, 1) <> "\" Then fallbackFolder = fallbackFolder & "\"
        gSaveFolderPath = fallbackFolder
    Else
        ' Folder found — confirm or let user change it
        Dim saveConfirm As VbMsgBoxResult
        saveConfirm = MsgBox( _
            "Reports will be saved to:" & vbNewLine & vbNewLine & _
            gSaveFolderPath & vbNewLine & vbNewLine & _
            "Click YES to use this location" & vbNewLine & _
            "Click NO to choose a different folder", _
            vbYesNoCancel + vbQuestion, "Confirm Save Location")

        If saveConfirm = vbCancel Then Exit Function

        If saveConfirm = vbNo Then
            Dim customFolder As String
            customFolder = BrowseForFolder("Select the folder to save reports to:")

            If customFolder = "" Then Exit Function
            If Right(customFolder, 1) <> "\" Then customFolder = customFolder & "\"
            gSaveFolderPath = customFolder
        End If
    End If

    GetRunMonthInput = True

End Function

'=============================================================================
' COPY FILTERED DATA FROM SOURCE TO DESTINATION SHEET
'=============================================================================
Private Sub CopyFilteredData(wsSource As Worksheet, wsDest As Worksheet, _
                              lastRow As Long, lastCol As Long, _
                              FilterType As String, RunMonthFilter As Boolean)

    ' OPTIMISED: Evaluates filter logic per row into a helper column, then
    ' uses AutoFilter + SpecialCells bulk copy instead of row-by-row .Copy.
    ' The evaluation loop is fast (just string reads + writes). The big speed
    ' gain comes from replacing 30k individual .Copy calls with ONE bulk copy.

    Dim paymentStatusCol As Long
    Dim monthPaidCol As Long
    Dim helperCol As Long
    Dim i As Long

    ' Find Payment Status column
    paymentStatusCol = FindColumnByHeader(wsSource, "Payment Status")
    If paymentStatusCol = 0 Then
        paymentStatusCol = FindColumnByHeader(wsSource, "PaymentStatus")
    End If

    If paymentStatusCol = 0 Then
        MsgBox "Could not find 'Payment Status' column in the CSV. Please check the file.", vbExclamation
        Exit Sub
    End If

    ' Find Month Paid column (needed for run-month filter)
    monthPaidCol = FindColumnByHeader(wsSource, "Month Paid")
    If monthPaidCol = 0 Then
        monthPaidCol = FindColumnByHeader(wsSource, "MonthPaid")
    End If
    If monthPaidCol = 0 Then
        monthPaidCol = FindColumnByHeader(wsSource, "Month_Paid")
    End If

    ' Clear any existing AutoFilter on source
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False

    ' Add a helper column to flag which rows to include
    helperCol = lastCol + 1
    wsSource.Cells(1, helperCol).Value = "_IncludeRow"

    Application.StatusBar = FilterType & ": Evaluating filter criteria..."

    For i = 2 To lastRow
        Dim matchesFilter As Boolean
        Dim cellVal As String
        matchesFilter = False

        cellVal = Trim(CStr(wsSource.Cells(i, paymentStatusCol).Value))

        Select Case FilterType
            Case "NewAndAllocated"
                If InStr(1, cellVal, "1. New", vbTextCompare) > 0 Or _
                   InStr(1, cellVal, "New", vbTextCompare) > 0 Or _
                   InStr(1, cellVal, "2. Commission Allocated", vbTextCompare) > 0 Or _
                   InStr(1, cellVal, "Commission Allocated", vbTextCompare) > 0 Then
                    matchesFilter = True
                End If

            Case "ReturnedError"
                If InStr(1, cellVal, "3. Returned", vbTextCompare) > 0 Or _
                   InStr(1, cellVal, "Returned - Error", vbTextCompare) > 0 Or _
                   InStr(1, cellVal, "Returned", vbTextCompare) > 0 Then
                    matchesFilter = True
                End If
        End Select

        ' Apply Month Paid filter if required
        ' Blanks are INCLUDED — only exclude rows with a different month
        If matchesFilter And RunMonthFilter And monthPaidCol > 0 Then
            Dim monthVal As String
            monthVal = Trim(CStr(wsSource.Cells(i, monthPaidCol).Value))
            If Len(monthVal) > 0 Then
                If Not MatchesRunMonth(monthVal) Then
                    matchesFilter = False
                End If
            End If
        End If

        If matchesFilter Then
            wsSource.Cells(i, helperCol).Value = "YES"
        Else
            wsSource.Cells(i, helperCol).Value = "NO"
        End If
    Next i

    ' AutoFilter on helper column to show only "YES" rows + header
    Dim fullRange As Range
    Set fullRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, helperCol))
    fullRange.AutoFilter Field:=helperCol, Criteria1:="YES"

    ' Copy header row separately (always visible)
    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastCol)).Copy _
        wsDest.Range("A1")

    ' Bulk-copy visible data rows (row 2 onwards, exclude helper column)
    On Error Resume Next
    Dim visibleRows As Range
    Set visibleRows = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, lastCol)).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleRows Is Nothing Then
        visibleRows.Copy wsDest.Range("A2")
    End If

    ' Count result rows
    Dim destRows As Long
    destRows = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row - 1
    Application.StatusBar = FilterType & ": Found " & destRows & " matching rows"

    ' Clean up: remove AutoFilter and delete helper column
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
    wsSource.Columns(helperCol).Delete

End Sub

'=============================================================================
' CHECK IF A MONTH VALUE MATCHES THE RUN MONTH
'=============================================================================
Private Function MatchesRunMonth(monthVal As String) As Boolean

    Dim testDate As Date
    MatchesRunMonth = False

    If Len(Trim(monthVal)) = 0 Then Exit Function

    On Error Resume Next

    ' Try parsing as a date directly
    testDate = CDate(monthVal)

    If Err.Number = 0 Then
        If Month(testDate) = Month(gRunMonthDate) And Year(testDate) = Year(gRunMonthDate) Then
            MatchesRunMonth = True
        End If
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear

    ' Try matching text like "August 2026" or "Aug 2026"
    If InStr(1, monthVal, gRunMonthName, vbTextCompare) > 0 And _
       InStr(1, monthVal, CStr(gRunMonthYear), vbTextCompare) > 0 Then
        MatchesRunMonth = True
        On Error GoTo 0
        Exit Function
    End If

    ' Try matching short month name
    If InStr(1, monthVal, Left(gRunMonthName, 3), vbTextCompare) > 0 And _
       InStr(1, monthVal, CStr(gRunMonthYear), vbTextCompare) > 0 Then
        MatchesRunMonth = True
        On Error GoTo 0
        Exit Function
    End If

    ' Try matching MM/YYYY pattern
    If InStr(1, monthVal, gRunMonthMM & "/" & CStr(gRunMonthYear), vbTextCompare) > 0 Then
        MatchesRunMonth = True
    End If

    ' Try matching YYYY-MM pattern
    If InStr(1, monthVal, CStr(gRunMonthYear) & "-" & gRunMonthMM, vbTextCompare) > 0 Then
        MatchesRunMonth = True
    End If

    On Error GoTo 0

End Function

'=============================================================================
' FIND A COLUMN NUMBER BY HEADER TEXT (case-insensitive)
'=============================================================================
Private Function FindColumnByHeader(ws As Worksheet, headerText As String) As Long

    Dim lastCol As Long
    Dim i As Long
    Dim cellVal As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For i = 1 To lastCol
        cellVal = Trim(CStr(ws.Cells(1, i).Value))
        If StrComp(cellVal, headerText, vbTextCompare) = 0 Then
            FindColumnByHeader = i
            Exit Function
        End If
    Next i

    ' Try partial match
    For i = 1 To lastCol
        cellVal = Trim(CStr(ws.Cells(1, i).Value))
        If InStr(1, cellVal, headerText, vbTextCompare) > 0 Then
            FindColumnByHeader = i
            Exit Function
        End If
    Next i

    FindColumnByHeader = 0

End Function

'=============================================================================
' FORMAT A SHAREPOINT TAB
' Order: Format dates > Format currency > Expand columns > Delete H & J >
'        Freeze top row > Convert to Table
'=============================================================================
Private Sub FormatSharePointTab(ws As Worksheet)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim tbl As ListObject
    Dim tblName As String

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' ---- 1. Format Date columns: D,G,I,L,AC,AT,BL,CD,CQ ----
    Dim dateCols As Variant
    dateCols = Array(4, 7, 9, 12, 29, 46, 64, 82, 95)

    Dim c As Variant
    For Each c In dateCols
        If CLng(c) <= lastCol Then
            FormatColumnAsDate ws, CLng(c), lastRow
        End If
    Next c

    ' ---- 2. Format Currency columns: W,AE,AS,AW,BK,BO,CC,CG ----
    Dim currCols As Variant
    currCols = Array(23, 31, 45, 49, 63, 67, 81, 85)

    For Each c In currCols
        If CLng(c) <= lastCol Then
            FormatColumnAsCurrency ws, CLng(c), lastRow
        End If
    Next c

    ' ---- 3. Auto-fit (expand) all columns ----
    ws.Cells.EntireColumn.AutoFit

    ' ---- 3b. Set all row heights to default (15) ----
    ws.Cells.EntireRow.RowHeight = 15

    ' ---- 4. Delete columns H (8) and J (10) ----
    ' Delete J first (col 10) then H (col 8) so indices don't shift
    Dim colH_Header As String
    Dim colJ_Header As String
    colH_Header = Trim(CStr(ws.Cells(1, 8).Value))
    colJ_Header = Trim(CStr(ws.Cells(1, 10).Value))

    ' Check col J = "QA Allocation"
    If InStr(1, colJ_Header, "QA", vbTextCompare) > 0 Or _
       InStr(1, colJ_Header, "Allocation", vbTextCompare) > 0 Then
        ws.Columns(10).Delete
    Else
        Dim qaCol As Long
        qaCol = FindColumnByHeader(ws, "QA Allocation")
        If qaCol = 0 Then qaCol = FindColumnByHeader(ws, "QA_Allocation")
        If qaCol > 0 Then ws.Columns(qaCol).Delete
    End If

    ' Check col H = "Payment Team Allocation"
    colH_Header = Trim(CStr(ws.Cells(1, 8).Value))
    If InStr(1, colH_Header, "Payment Team", vbTextCompare) > 0 Or _
       InStr(1, colH_Header, "Allocation", vbTextCompare) > 0 Then
        ws.Columns(8).Delete
    Else
        Dim ptCol As Long
        ptCol = FindColumnByHeader(ws, "Payment Team Allocation")
        If ptCol = 0 Then ptCol = FindColumnByHeader(ws, "Payment_Team_Allocation")
        If ptCol > 0 Then ws.Columns(ptCol).Delete
    End If

    ' Recalculate dimensions after deletion
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' ---- 5. Freeze top row ----
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' ---- 6. Convert to Table ----
    tblName = Replace(ws.Name, " ", "_")
    tblName = CleanTableName(tblName)

    On Error Resume Next
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), _
        XlListObjectHasHeaders:=xlYes)

    If Not tbl Is Nothing Then
        tbl.Name = tblName
        tbl.TableStyle = "TableStyleMedium2"
    End If
    On Error GoTo 0

End Sub

'=============================================================================
' FORMAT A COLUMN AS SHORT DATE
'=============================================================================
Private Sub FormatColumnAsDate(ws As Worksheet, colNum As Long, lastRow As Long)

    Dim rng As Range
    Dim cell As Range
    Dim cellVal As Variant

    Set rng = ws.Range(ws.Cells(2, colNum), ws.Cells(lastRow, colNum))

    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            On Error Resume Next
            If Not IsDate(cell.Value) Then
                cellVal = CDate(cell.Value)
                If Err.Number = 0 Then
                    cell.Value = cellVal
                End If
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next cell

    rng.NumberFormat = "dd/mm/yyyy"

End Sub

'=============================================================================
' FORMAT A COLUMN AS CURRENCY
'=============================================================================
Private Sub FormatColumnAsCurrency(ws As Worksheet, colNum As Long, lastRow As Long)

    Dim rng As Range
    Dim cell As Range

    Set rng = ws.Range(ws.Cells(2, colNum), ws.Cells(lastRow, colNum))

    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            On Error Resume Next
            If Not IsNumeric(cell.Value) Then
                cell.Value = CDbl(cell.Value)
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next cell

    rng.NumberFormat = "$#,##0.00"

End Sub

'=============================================================================
' CLEAN TABLE NAME (remove invalid characters)
'=============================================================================
Private Function CleanTableName(name As String) As String

    Dim result As String
    Dim i As Long
    Dim ch As String

    result = ""
    For i = 1 To Len(name)
        ch = Mid(name, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            result = result & ch
        End If
    Next i

    If Len(result) > 0 And IsNumeric(Left(result, 1)) Then
        result = "Tbl_" & result
    End If

    If Len(result) = 0 Then result = "Table1"

    CleanTableName = result

End Function
