Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text

Sub Security_Event_Log_XML_Output_To_IR_Format_By_Keyword()

'------------------------------------------------------------------------------
'Make VBA Code Run Faster
'------------------------------------------------------------------------------
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.EnableEvents = False

'------------------------------------------------------------------------------
'Convert Table To a Range
'------------------------------------------------------------------------------
Dim rList As Range

With Worksheets("Sheet1").ListObjects("Table1")
    Set rList = .Range
    .Unlist                     'convert the table back to a range
End With

With rList
    .Interior.ColorIndex = xlColorIndexNone
    .Font.ColorIndex = xlColorIndexAutomatic
    .Borders.LineStyle = xlLineStyleNone
End With

'------------------------------------------------------------------------------
'Prompt User to Open keywords.txt file
'powerspreadsheets.com/vba-open-workbook
'------------------------------------------------------------------------------
Dim keywordFilePath As Variant
Dim textFile As String

textFile = FreeFile

MsgBox "Browse to the keyword text file to open"

keywordFilePath = Application.GetOpenFilename(Title:="Open Keyword File")
If keywordFilePath <> False Then
    Open keywordFilePath For Input As textFile
Else
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
'This Code section prompts the analyst to enter the computer name
'--------------------------------------------------------------------------
Dim hostName As String

hostName = Application.InputBox("Enter the Computer Name associated with this file")

'------------------------------------------------------------------------------
'Define Columns to Keep
'------------------------------------------------------------------------------
Dim eventName As Variant
Dim eventID As Variant
Dim eventRecordID As Variant
Dim eventUserID As Variant
Dim eventSubjectUserName As Variant
Dim eventMessage As Variant
Dim eventData As Variant
Dim eventName2 As Variant
Dim systemTime As Variant
Dim hexData As Variant

systemTime = Application.Match("SystemTime", Sheets(1).Rows(1), 0)
eventName = Application.Match("Name", Sheets(1).Rows(1), 0)
eventID = Application.Match("ns?:EventID", Sheets(1).Rows(1), 0)
eventRecordID = Application.Match("ns?:EventRecordID", Sheets(1).Rows(1), 0)

If Not IsError(Application.Match("UserID", Sheets(1).Rows(1), 0)) Then
    eventUserID = Application.Match("UserID", Sheets(1).Rows(1), 0)
Else
    eventUserID = 1000
End If

If Not IsError(Application.Match("ns?:SubjectUserName", Sheets(1).Rows(1), 0)) Then
    eventSubjectUserName = Application.Match("ns?:SubjectUserName", Sheets(1).Rows(1), 0)
Else
    eventSubjectUserName = 1001
End If

If Not IsError(Application.Match("ns?:Message", Sheets(1).Rows(1), 0)) Then
    eventMessage = Application.Match("ns?:Message", Sheets(1).Rows(1), 0)
Else
    eventMessage = 1002
End If

If Not IsError(Application.Match("ns?:Data", Sheets(1).Rows(1), 0)) Then
    eventData = Application.Match("ns?:Data", Sheets(1).Rows(1), 0)
Else
    eventData = 1003
End If

If Not IsError(Application.Match("Name2", Sheets(1).Rows(1), 0)) Then
    eventName2 = Application.Match("Name2", Sheets(1).Rows(1), 0)
Else
    eventName2 = 1004
End If

If Not IsError(Application.Match("ns?:Binary", Sheets(1).Rows(1), 0)) Then
    hexData = Application.Match("ns?:Binary", Sheets(1).Rows(1), 0)
Else
    hexData = 1005
End If

'------------------------------------------------------------------------------
'This code section deletes specific columns
'stackoverflow.com/questions/16597841/deleting-all-columns-with-certain-headings
'------------------------------------------------------------------------------
Dim deleteColumn As Variant
Dim headerName As String

For deleteColumn = ActiveSheet.UsedRange.Columns.Count To 1 Step -1

    headerName = deleteColumn

    Select Case headerName

        Case eventName
            ActiveSheet.Columns(eventName).Delete
        Case eventID
            'Do Nothing
        Case eventRecordID
            'Do Nothing
        Case eventUserID
            'Do Nothing
        Case eventSubjectUserName
            'Do Nothing
        Case eventMessage
            'Do Nothing
        Case systemTime
            'Do Nothing
        Case hexData
            'Do Nothing
        Case eventData
            'Do Nothing
        Case eventName2
            'Do Nothing
        Case Else
            ActiveSheet.Columns(deleteColumn).Delete
    End Select

Next

'------------------------------------------------------------------------------
'In this code section, moving columns
'------------------------------------------------------------------------------
Sheets(1).Columns("A").Cut
Sheets(1).Columns("F").Insert Shift:=xlToRight

Sheets(1).Columns("B").Cut
Sheets(1).Columns("F").Insert Shift:=xlToRight

If Not eventUserID = 1000 Then
    Sheets(1).Columns("E").Cut
    Sheets(1).Columns("C").Insert Shift:=xlToRight
Else
    Sheets(1).Columns("B").Insert Shift:=xlToRight
End If

Sheets(1).Columns("D").Cut
Sheets(1).Columns("C").Insert Shift:=xlToRight

If Not eventSubjectUserName = 1001 Then
    Sheets(1).Columns("G").Insert Shift:=xlToRight
End If

Sheets(1).Columns("C").Insert Shift:=xlToRight

'------------------------------------------------------------------------------
'Find last row and last column
'www.rondebruin.nl/win/s9/win005.htm
'------------------------------------------------------------------------------
Dim lastRow As Long
Dim lastColumn As Long

With ActiveSheet.UsedRange
    lastRow = .Rows(.Rows.Count).Row
    lastColumn = .Columns(.Columns.Count).Column
End With

'------------------------------------------------------------------------------
'Fill blank cells with a hyphon. Part of code from
'www.rondebruin.nl/win/s9/win005.htm
'------------------------------------------------------------------------------
Dim w As Long
Dim d As Long

For w = 1 To lastColumn
    For d = 1 To lastRow
        If IsEmpty(ActiveSheet.Cells(d, w).Value) Then
            ActiveSheet.Cells(d, w).Value = "-"
        End If
    Next
Next

'------------------------------------------------------------------------------
'Find new binary data column
'------------------------------------------------------------------------------
Dim eventBin As Variant

eventBin = Application.Match("ns?:Binary", Sheets(1).Rows(1), 0)

'------------------------------------------------------------------------------
'Convert Hex to ASCII
'------------------------------------------------------------------------------
Dim fnd As Variant
Dim rplc As Variant
Dim binData() As String
Dim convertData As String
Dim l As Long
Dim z As Long
Dim y As Long
Dim c As Long

l = 2
fnd = "00"
rplc = ""

If Not hexData = 1005 Then

    '------------------------------------------------------------------------------
    'Replace Unicode 2 bytes with ASCII 1 byte
    '------------------------------------------------------------------------------
    Sheets(1).Cells(l, eventBin).Select

    Do Until IsEmpty(ActiveCell)

        Sheets(1).Cells(l, eventBin).Replace what:=fnd, Replacement:=rplc, _
        LookAt:=xlPart, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False

        l = l + 1
        ActiveCell.Offset(1, 0).Select

    Loop

    '------------------------------------------------------------------------------
    'Fill blank cells with a hyphon so Hex to Ascii does not stop on Binary Cells
    'filled with all zeros
    '------------------------------------------------------------------------------
    For d = 1 To lastRow
        If IsEmpty(ActiveSheet.Cells(d, eventBin).Value) Then
            ActiveSheet.Cells(d, eventBin).Value = "-"
        End If
    Next

    '------------------------------------------------------------------------------
    'Starts Conversion of Hex to ASCII
    '------------------------------------------------------------------------------
    z = 2
    Sheets(1).Cells(1, eventBin).Select

    Do Until IsEmpty(ActiveCell)

    convertData = ""
    y = 0

    ReDim binData(Len(Cells(z, eventBin).Value))

    For c = 0 To (Len(Cells(z, eventBin).Value) / 2)

        binData(c) = Chr(Val("&H" & Mid(Cells(z, eventBin).Value, y + 1, 2)))
        convertData = convertData & binData(c)
        y = y + 2

    Next

    Cells(z, eventBin) = convertData

    ActiveCell.Offset(1, 0).Select

    z = z + 1

    Loop

End If

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
'Change Formatting of date/time in SystemTime Column
'------------------------------------------------------------------------------
Dim findChar As Variant
Dim replaceChar As Variant
Dim s As Long

findChar = "T"
replaceChar = " "

s = 2
Sheets(1).Cells(s, "A").Select

'------------------------------------------------------------------------------
'Remove "T" in XML Time Format and replace with " "
'------------------------------------------------------------------------------
Do Until IsEmpty(ActiveCell)

    Sheets(1).Cells(s, "A").Replace what:=findChar, Replacement:=replaceChar, _
    LookAt:=xlPart, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False

    s = s + 1
    ActiveCell.Offset(1, 0).Select

Loop

'------------------------------------------------------------------------------
'Remove Milliseconds
'------------------------------------------------------------------------------
Dim fullCell As String
Dim splitMilli() As String
Dim correctTime As String
Dim m As Long

m = 2
Sheets(1).Cells(m, "A").Select

Do Until IsEmpty(ActiveCell)

    fullCell = ActiveCell.Value
    splitMilli = Split(fullCell, ".")
    correctTime = splitMilli(0)
    Sheets(1).Cells(m, "A").Value = correctTime

    m = m + 1
    ActiveCell.Offset(1, 0).Select

Loop

'------------------------------------------------------------------------------
'Make Excel recognize cell as a Date/Time
'stackoverflow.com/questions/20375233/excel/convert-text-to-date
'------------------------------------------------------------------------------
With ActiveSheet.UsedRange.Columns("A").Cells
    .TextToColumns Destination:=.Cells(1), DataType:=xlFixedWidth, FieldInfo:=Array(0, xlMDYFormat)
    .NumberFormat = "mm/dd/yyyy hh:mm:ss"
End With

'------------------------------------------------------------------------------
'Find last row
'www.rondebruin.nl/win/s9/win005.htm
'------------------------------------------------------------------------------
Dim lastRow3 As Long

With ActiveSheet.UsedRange
    lastRow3 = .Rows(.Rows.Count).Row
End With

'------------------------------------------------------------------------------
'Combine Column, Inserts Account Name, Computer Name and Artifact Name in each
'cell of each column
'------------------------------------------------------------------------------
Dim f As Long

For f = 2 To lastRow3
    Cells(f, "B").Value = "N/A"
    Cells(f, "C").Value = hostName
    Cells(f, "F") = "Evt ID: " & Cells(f, "F")
    Cells(f, "G") = "Evt Record #: " & Cells(f, "G")
    Cells(f, "H").Value = "Security Event Log"
Next

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
Sheets(1).Cells(1, "H") = "Artifact"

'------------------------------------------------------------------------------
'Worksheet is sorted by column A, oldest to newest
'------------------------------------------------------------------------------
Sheets(1).Columns.Sort Key1:=Sheets(1).Range("A1"), Header:=xlYes

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

''--------------------------------------------------------------------------
''This code section picks dates to filter on
''--------------------------------------------------------------------------
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

