Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text

Sub User_Shellbag_by_Keyword()

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
'This Code section prompts the analyst to enter the user profile name and
'computer name
'--------------------------------------------------------------------------
Dim hostName As String
Dim profileName As String

profileName = Application.InputBox("Enter the User Name associated with this file")
hostName = Application.InputBox("Enter the Computer Name associated with this file")

'--------------------------------------------------------------------------
'This code section deletes specific columns
'--------------------------------------------------------------------------
Sheets(1).Range("A:D,F:H,M:O,R:R").EntireColumn.Delete

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

'--------------------------------------------------------------------------
'This code section removes elements from a string in a cell in the new
'first row
'--------------------------------------------------------------------------
Dim sht As Worksheet
Dim fnd1 As Variant
Dim fnd2 As Variant
Dim rplc As Variant

fnd1 = " +00:00"
fnd2 = " +00:00 "
rplc = ""

For Each sht In ActiveWorkbook.Worksheets

sht.Cells.Replace what:=fnd1, Replacement:=rplc, _
LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
SearchFormat:=False, ReplaceFormat:=False

sht.Cells.Replace what:=fnd2, Replacement:=rplc, _
LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
SearchFormat:=False, ReplaceFormat:=False

Next sht

'--------------------------------------------------------------------------
'This code section removes the first unknown element in a string in a cell
'in the new first row
'--------------------------------------------------------------------------
Dim txt As String
Dim i As Long
Dim y As Long
Dim f As Long
Dim CellContent As Variant
Dim arrayFields As Long
Dim newOutput As String
Dim nextField As String

y = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    newOutput = ""
    f = 0
    txt = Cells(y, 1).Value
    CellContent = Split(txt, "\")
    arrayFields = UBound(CellContent)
    
    If arrayFields <> 0 Then
    
        For i = 0 To arrayFields
            f = f + 1
            If f > arrayFields Then
                f = f
            Else
                nextField = Trim(CellContent(f))
                newOutput = newOutput & nextField & "\"
            End If
        Next i
    
    Else
        newOutput = CellContent(0)
    End If
    
    Cells(y, 1).Value = newOutput
    
    y = y + 1
    ActiveCell.Offset(1, 0).Select

Loop

'--------------------------------------------------------------------------
'This code section makes blank "text" fields either Empty or Date for the
'AccessedOn column (D)
'--------------------------------------------------------------------------
Dim c As Long

c = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    If WorksheetFunction.IsText(Cells(c, "D")) Then
        Cells(c, "D") = Cells(c, "F")
    End If
    
    c = c + 1
    ActiveCell.Offset(1, 0).Select

Loop

'--------------------------------------------------------------------------
'This code section changes Empty cells from FirstExplored column (F)
'--------------------------------------------------------------------------
Dim l As Long

l = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    If Not IsDate(Cells(l, "F")) Then
        Cells(l, "F") = Cells(l, "D")
    End If
    
    l = l + 1
    ActiveCell.Offset(1, 0).Select

Loop

'--------------------------------------------------------------------------
'This code section calculates the earliest date of access by the user
'account
'--------------------------------------------------------------------------
Dim z As Long

z = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    If Cells(z, "F") < Cells(z, "D") Then
        Cells(z, "H") = Cells(z, "F")
    Else
        Cells(z, "H") = Cells(z, "D")
    End If
    
    z = z + 1
    ActiveCell.Offset(1, 0).Select

Loop

'--------------------------------------------------------------------------
'This code section deletes specific columns
'--------------------------------------------------------------------------
Sheets(1).Range("B:D,F:G").EntireColumn.Delete

'--------------------------------------------------------------------------
'In this code section, moving columns
'--------------------------------------------------------------------------
Columns("C").Cut
Columns("B").Insert Shift:=xlToRight

'--------------------------------------------------------------------------
'In this code section the First Access and LastWriteTime are compared. If
'the time difference is 3 seconds or less, it is considered the same
'connection and only counted once.
'--------------------------------------------------------------------------
Dim t As Long
Dim dtDiff As Double
Dim lastWrite As Date
Dim firstAccess As Date

t = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

If WorksheetFunction.IsText(Cells(t, "B")) = True Then
    dtDiff = 0.0003
Else
    lastWrite = Cells(t, "C")
    firstAccess = Cells(t, "B")
    dtDiff = lastWrite - firstAccess
End If

If dtDiff > 0.0003 Then
    ActiveCell.EntireRow.Select
    Selection.Copy
    Selection.Insert Shift:=xlDown

'--------------------------------------------------------------------------
'The line below will clear the clipboard of all the copied data
'--------------------------------------------------------------------------
    Application.CutCopyMode = False

    ActiveCell.Offset(0, 1).Value = ""
    ActiveCell.Offset(1, 2).Value = ""

    If Cells(t, "B") = 0 Then
        Cells(t, "D") = Cells(t, "C")
        Cells(t, "E") = "Last Accessed"
        Cells(t + 1, "D") = Cells(t + 1, "B")
        Cells(t + 1, "E") = "First Accessed"
    Else
        Cells(t, "D") = Cells(t, "B")
        Cells(t, "E") = "Last Accessed"
        Cells(t + 1, "D") = Cells(t + 1, "C")
        Cells(t + 1, "E") = "First Accessed"
    End If

    t = t + 2
    ActiveCell.Offset(2, 0).Select
Else
    ActiveCell.Offset(0, 1).Value = ""

    Cells(t, "D") = Cells(t, "C")
    Cells(t, "E") = "Last Accessed"

    t = t + 1
    ActiveCell.Offset(1, 0).Select
End If

Loop

'--------------------------------------------------------------------------
'This code section deletes rows with no date/time value
'--------------------------------------------------------------------------
Dim e As Long

e = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    If IsEmpty(Cells(e, "D")) Then
        ActiveSheet.Rows(e).EntireRow.Delete
    Else
        e = e + 1
        ActiveCell.Offset(1, 0).Select
    End If

Loop

'--------------------------------------------------------------------------
'Delete specific columns
'--------------------------------------------------------------------------
Sheets(1).Range("B:C").EntireColumn.Delete

'--------------------------------------------------------------------------
'In this code section, moving more columns
'--------------------------------------------------------------------------
Columns("C").Cut
Columns("E").Insert Shift:=xlToRight

Columns("A").Cut
Columns("E").Insert Shift:=xlToRight

Columns("C").Insert Shift:=xlToRight

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
Dim lastRow2 As Long

With ActiveSheet.UsedRange
    lastRow2 = .Rows(.Rows.Count).Row
End With

'------------------------------------------------------------------------------
'Inserts Account Name, Computer Name and Artifact Name in each cell of each column
'------------------------------------------------------------------------------
Dim n As Long

For n = 2 To lastRow2
    Cells(n, "B").Value = profileName
    Cells(n, "C").Value = hostName
    Cells(n, "H").Value = "Shellbags"
Next

'--------------------------------------------------------------------------
'In this code section, all date/times are formatted
'--------------------------------------------------------------------------
Sheets(1).Columns("A").NumberFormat = "mm/dd/yyyy hh:mm:ss"

'--------------------------------------------------------------------------
'Workseet is sorted by column A, oldest to newest.
'--------------------------------------------------------------------------
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
''This code section prompts the user to filter events by date. In most cases,
''Event Viewer logs should be filtered by date on export. Uncomment this
''section if this feature is desired.
''--------------------------------------------------------------------------
'Dim yesNo As Variant
'Dim startDate As Date
'Dim endDate As Date
'Dim u As Long
'
'yesNo = MsgBox("Do you want to filter by a time frame?", vbYesNo)
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
