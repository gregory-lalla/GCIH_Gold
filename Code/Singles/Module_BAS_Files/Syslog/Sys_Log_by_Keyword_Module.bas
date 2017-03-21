Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text

Sub Syslog_Log_by_Keyword()

'------------------------------------------------------------------------------
'Make VBA Code Run Faster
'------------------------------------------------------------------------------
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.EnableEvents = False

'------------------------------------------------------------------------------
'Prompt User to Open keywords.txt file
'powerspreadsheets.com/vba-open-workbook
'------------------------------------------------------------------------------
Dim keywordFilePath As Variant
Dim textFile As String

textFile = FreeFile

MsgBox "Select the keyword text file to open"

keywordFilePath = Application.GetOpenFilename(Title:="Open Keyword File")
If keywordFilePath <> False Then
    Open keywordFilePath For Input As textFile
Else
    MsgBox "You need to first select the Keyword File"
    ActiveWorkbook.Close savechanges:=False
    Exit Sub
End If

'------------------------------------------------------------------------------
'Create an array with contents of keywords.txt file
'**some of this code from thespreadsheetguru.com/blog/vba-guide-text-files
'------------------------------------------------------------------------------
Dim keywordList As String
Dim keywordArray() As String

keywordList = Input(LOF(textFile), textFile)
keywordArray = Split(keywordList, vbCrLf)

Close textFile

'--------------------------------------------------------------------------
'This Code retrieves computer name from cell B1
'--------------------------------------------------------------------------
Dim hostName As String

hostName = Cells(1, 2).Value

'--------------------------------------------------------------------------
'Insert Header Row
'--------------------------------------------------------------------------
Sheets(1).Rows(1).Insert Shift:=xlToRight

'------------------------------------------------------------------------------
'Find last row
'www.rondebruin.nl/win/s9/win005.htm
'------------------------------------------------------------------------------
Dim lastRow As Long

With ActiveSheet.UsedRange
    lastRow = .Rows(.Rows.Count).Row
End With

'------------------------------------------------------------------------------
'This code section finds rows that have keywords in them
'------------------------------------------------------------------------------
Dim q As Long
Dim boolTest As Boolean
Dim keepRow() As Variant
Dim reDimNumber As Long
Dim sizeKeepRow As Long

boolTest = False
ActiveWorkbook.Sheets(1).Range("A1").Select

'------------------------------------------------------------------------------
'Finds rows to keep and adds to an array to track
'blogs.office.com/2008/10/03/what-is-the-fastest-way-to-scan-a-large-range-in-excel
'------------------------------------------------------------------------------
Dim rowNumb As Long
Dim maxRows As Long
Dim colNumb As Long
Dim maxCols As Long

maxRows = Range("A1").CurrentRegion.Rows.Count
maxCols = Range("A1").CurrentRegion.Columns.Count

sizeKeepRow = 1

For rowNumb = 2 To maxRows

    For colNumb = 1 To maxCols

        For q = 0 To UBound(keywordArray)
        
            If Not InStr(1, Cells(rowNumb, colNumb), keywordArray(q), vbTextCompare) > 0 Then
                boolTest = True
            Else
                boolTest = False
                Exit For
            End If
        Next

        If boolTest = False Then
            reDimNumber = sizeKeepRow
            ReDim Preserve keepRow(reDimNumber)
            keepRow(sizeKeepRow - 1) = rowNumb
            sizeKeepRow = sizeKeepRow + 1
            Exit For
        End If
        
    Next
    
Next

'------------------------------------------------------------------------------
'This code section tests to see if array is empty
'stackoverflow.com/questions/9874086/vba-dont-go-into-loop-when-array-is-empty
'------------------------------------------------------------------------------
If Len(Join(keepRow, "")) = 0 Then

    MsgBox "No Keyword Hits Detected"

Else

    '------------------------------------------------------------------------------
    'This code section removes duplicates from the array
    'blogs.technet.microsoft.com/heyscriptingguy/2006/10/27/how-can-i-delete-duplicate-items-from-an-array
    '------------------------------------------------------------------------------
    Dim objDictionary As Object
    Dim strItem As Variant
    Dim strKey As Variant
    Dim intItems As Long
    Dim k As Long

    Set objDictionary = CreateObject("Scripting.Dictionary")

    For Each strItem In keepRow
        If Not objDictionary.Exists(strItem) Then
            objDictionary.Add strItem, strItem
        End If
    Next

    intItems = objDictionary.Count - 1

    ReDim keepRow(intItems)

    k = 0

    For Each strKey In objDictionary.Keys
        keepRow(k) = strKey
        k = k + 1
    Next

    '------------------------------------------------------------------------------
    'This code section removes empty elements in the arrays
    '------------------------------------------------------------------------------
    Dim rowsToKeep() As Variant
    Dim a As Long
    Dim j As Long

    j = 0

    ReDim rowsToKeep(LBound(keepRow) To UBound(keepRow))

    For a = LBound(keepRow) To UBound(keepRow)
        If Not IsEmpty(keepRow(a)) Then
            rowsToKeep(j) = keepRow(a)
            j = j + 1
        End If
    Next

    ReDim Preserve rowsToKeep(LBound(keepRow) To (j - 1))

    '------------------------------------------------------------------------------
    'Sort array descending
    'mrexcel.com/forum/excel-questions/690718-visual-basic-applications-sort-array-numbers.htm
    '------------------------------------------------------------------------------
    Dim srtTemp As Variant
    Dim g As Long
    Dim h As Long

    For g = LBound(rowsToKeep) To UBound(rowsToKeep)
        For h = g + 1 To UBound(rowsToKeep)
            If rowsToKeep(g) < rowsToKeep(h) Then
                srtTemp = rowsToKeep(h)
                rowsToKeep(h) = rowsToKeep(g)
                rowsToKeep(g) = srtTemp
            End If
        Next h
    Next g

    '------------------------------------------------------------------------------
    'Delete Rows with No Keywords
    '------------------------------------------------------------------------------
    Dim x As Variant
    Dim b As Variant

    b = lastRow + 1

    For Each x In rowsToKeep
        ActiveWorkbook.Sheets(1).Rows(x).Copy
        ActiveWorkbook.Sheets(1).Rows(b).Insert Shift:=xlToRight
        b = b + 1
    Next

    For b = lastRow To 2 Step -1
        ActiveSheet.Rows(b).EntireRow.Delete
    Next
    
End If

'------------------------------------------------------------------------------
'Date/Time Column is formatted
'------------------------------------------------------------------------------
Sheets(1).Columns("A").NumberFormat = "mm/dd/yyyy hh:mm:ss"

'--------------------------------------------------------------------------
'Insert Columns
'--------------------------------------------------------------------------
Sheets(1).Columns("B").Insert Shift:=xlToRight

'------------------------------------------------------------------------------
'Change Row Headers
'------------------------------------------------------------------------------
Sheets(1).Cells(1, "A") = "Date/Time"
Sheets(1).Cells(1, "B") = "Account"
Sheets(1).Cells(1, "C") = "Computer"
Sheets(1).Cells(1, "D") = "Description"
Sheets(1).Cells(1, "E") = "Details"
Sheets(1).Cells(1, "F") = "Properties"
Sheets(1).Cells(1, "G") = "Miscellaneous"
Sheets(1).Cells(1, "H") = "Artifacts"

'------------------------------------------------------------------------------
'Find last row
'www.rondebruin.nl/win/s9/win005.htm
'------------------------------------------------------------------------------
Dim lastRow3 As Long

With ActiveSheet.UsedRange
    lastRow3 = .Rows(.Rows.Count).Row
End With

'------------------------------------------------------------------------------
'Inserts Account Name, Computer Name and Artifact Name in each cell of each column
'------------------------------------------------------------------------------
Dim n As Long

For n = 2 To lastRow3
    Cells(n, "B").Value = "N/A"
    Cells(n, "C").Value = hostName
    Cells(n, "H").Value = "Syslog Log"
Next

'------------------------------------------------------------------------------
'Freeze the first row.
'------------------------------------------------------------------------------
Sheets(1).Rows("2:2").Select
ActiveWindow.FreezePanes = True

'------------------------------------------------------------------------------
'Bold the first row.
'------------------------------------------------------------------------------
ActiveWorkbook.Sheets(1).Rows(1).Font.Bold = True

'------------------------------------------------------------------------------
'Filtering is enabled on columns.
'------------------------------------------------------------------------------
ActiveSheet.Range("A1").AutoFilter

'------------------------------------------------------------------------------
'Turn Off Wrap Text, Autofit Column Widths and Align All Cells to the Left
'------------------------------------------------------------------------------
With ActiveWorkbook.Sheets(1)
    .Columns.WrapText = False
    .Columns.HorizontalAlignment = xlLeft
    .Columns.AutoFit
End With

'--------------------------------------------------------------------------
'This code section prompts the user to filter events by date. In most cases,
'Event Viewer logs should be filtered by date on export. Uncomment this
'section if this feature is desired.
'--------------------------------------------------------------------------
'Dim startDate As Date
'Dim endDate As Date
'Dim yesNo As Variant
'Dim u As Long
'
'yesNo = MsgBox("Do you want to filter by a time frame? Yes or No", vbYesNo)
'
'If yesNo = vbYes Then
'    Selection.NumberFormat = "mm/dd/yyyy"
'    startDate = Application.InputBox("Enter the Start Date in format mm/dd/yyyy")
'    endDate = Application.InputBox("Enter the End Date in format mm/dd/yyyy")
'
'    u = 2
'    Range("A2").Select
'
'    Do Until IsEmpty(ActiveCell)
'
'        If Cells(u, "A") < startDate Or Cells(u, "A") > endDate + 1 Then
'            ActiveSheet.Rows(u).EntireRow.Delete
'        Else
'            u = u + 1
'            ActiveCell.Offset(1, 0).Select
'        End If
'
'    Loop
'
'End If

'------------------------------------------------------------------------------
'Clear Clipboard of data
'------------------------------------------------------------------------------
Application.CutCopyMode = False
    
'------------------------------------------------------------------------------
'Turn Off VBA Code Optimization
'------------------------------------------------------------------------------
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.EnableEvents = False

End Sub
