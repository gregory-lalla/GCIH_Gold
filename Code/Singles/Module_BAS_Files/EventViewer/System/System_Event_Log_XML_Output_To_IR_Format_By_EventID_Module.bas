Attribute VB_Name = "Module1"
Option Explicit

Sub System_Event_Log_XML_Output_To_IR_Format_By_EventID()

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
'Prompt User to Open Event ID File
'powerspreadsheets.com/vba-open-workbook
'------------------------------------------------------------------------------
Dim evtIdFilePath As Variant
Dim textFile As String

textFile = FreeFile

MsgBox "Select the file that contains the System EventIDs to Keep"

evtIdFilePath = Application.GetOpenFilename(Title:="Open System EventIDs File")

If evtIdFilePath <> False Then
    Open evtIdFilePath For Input As textFile
Else
    ActiveWorkbook.Close savechanges:=False
    Exit Sub
End If

'------------------------------------------------------------------------------
'Create an array with contents of System EventIDs file
'**some of this code from thespreadsheetguru.com/blog/vba-guide-text-files
'------------------------------------------------------------------------------
Dim evtIdList As String
Dim evtIdArray() As String

evtIdList = Input(LOF(textFile), textFile)
evtIdArray = Split(evtIdList, vbCrLf)

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
            'Do Nothing
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
            ActiveSheet.Columns(eventName2).Delete
            'Do Nothing
        Case Else
            ActiveSheet.Columns(deleteColumn).Delete
    End Select

Next

'------------------------------------------------------------------------------
'In this code section, moving columns
'------------------------------------------------------------------------------
Sheets(1).Columns("A").Cut
Sheets(1).Columns("D").Insert Shift:=xlToRight

Sheets(1).Columns("A").Cut
Sheets(1).Columns("D").Insert Shift:=xlToRight

If Not eventUserID = 1000 Then
    Sheets(1).Columns("E").Cut
    Sheets(1).Columns("B").Insert Shift:=xlToRight
Else
    Sheets(1).Columns("B").Insert Shift:=xlToRight
End If

If Not eventSubjectUserName = 1001 Then
    Sheets(1).Columns("G").Insert Shift:=xlToRight
End If

Sheets(1).Columns("C").Insert Shift:=xlToRight

'------------------------------------------------------------------------------
'Find Event ID Column
'------------------------------------------------------------------------------
Dim evtID As Variant

evtID = Application.Match("ns?:EventID", Sheets(1).Rows(1), 0)

'------------------------------------------------------------------------------
'This code section keeps the rows that have event ids listed in the text file
'------------------------------------------------------------------------------
Dim p As Long
Dim q As Long
Dim boolTest As Boolean

p = 2
Sheets(1).Cells(p, evtID).Select
boolTest = False

Do Until IsEmpty(ActiveCell)
    For q = 0 To UBound(evtIdArray)
        If Not ActiveCell.Value = evtIdArray(q) Then
            boolTest = True
        Else
            boolTest = False
            Exit For
        End If
    Next

    If boolTest = True Then
        ActiveCell.EntireRow.Delete
    Else
        ActiveCell.Offset(1, 0).Select
    End If
Loop

'------------------------------------------------------------------------------
'Check for no hits
'------------------------------------------------------------------------------
If IsEmpty(Cells(2, 1)) Then
    MsgBox "No Event IDs Found"
    ActiveWorkbook.Close savechanges:=False
    Exit Sub
End If

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
    If eventUserID = 1000 Then
        Cells(f, "B").Value = "N/A"
    End If
    Cells(f, "C").Value = hostName
    Cells(f, "E") = "Evt ID: " & Cells(f, "E") & " | Evt Record #: " & Cells(f, "F")
    Cells(f, "I").Value = "System Event Log"
Next

'------------------------------------------------------------------------------
'Delete Column
'------------------------------------------------------------------------------
Sheets(1).Columns("F").EntireColumn.Delete

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

