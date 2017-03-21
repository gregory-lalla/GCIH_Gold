Attribute VB_Name = "Module1"
Option Explicit

Sub EventLogs_System_CSV_Unique_IDs_per_Sheet()

'------------------------------------------------------------------------------
'Make VBA Code Run Faster
'------------------------------------------------------------------------------
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.EnableEvents = False


'--------------------------------------------------------------------------
'This code section prompts the user to filter events by date. In most cases,
'Event Viewer logs should be filtered by date on export. Uncomment this
'section if this feature is desired.
'--------------------------------------------------------------------------
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
'    Range("B2").Select
'
'    Do Until IsEmpty(ActiveCell)
'
'        If Cells(u, "B") < startDate Or Cells(u, "B") > endDate + 1 Then
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
'This code section removes the newline character and carriage return character
'in a string in a cell
'------------------------------------------------------------------------------
Dim txt As String
Dim y As Integer

y = 2

Sheets(1).Range("F2").Select
Do Until IsEmpty(ActiveCell)

    txt = ""
    txt = Sheets(1).Cells(y, 6).Value
    txt = Replace(txt, Chr(13), "#")
    txt = Replace(txt, Chr(10), "")
    
    Sheets(1).Cells(y, 6).Value = WorksheetFunction.Trim(txt)
    Sheets(1).Cells(y, 6).WrapText = False
    
    y = y + 1
    ActiveCell.Offset(1, 0).Select

Loop

'------------------------------------------------------------------------------
'Converts text in cells to columns delimited with the # character
'------------------------------------------------------------------------------
Sheets(1).Range("F:F").TextToColumns DataType:=xlDelimited, Other:=True, OtherChar:="#"

'------------------------------------------------------------------------------
'This code section deletes columns
'------------------------------------------------------------------------------
Sheets(1).Range("A:A,C:C").EntireColumn.Delete

'------------------------------------------------------------------------------
'Create an Array of Event IDs from Event ID Column
'blogs.office.com/2008/10/03/what-is-the-fastest-way-to-scan-a-large-range-in-excel
'------------------------------------------------------------------------------
Dim evtIDArray() As Variant
Dim rowNumb As Long
Dim maxRows As Long
Dim reDimNumber As Long
Dim sizeKeepRow As Long

maxRows = Sheets(1).Range("B1").CurrentRegion.Rows.Count
sizeKeepRow = 1

For rowNumb = 2 To maxRows
    reDimNumber = sizeKeepRow
    ReDim Preserve evtIDArray(reDimNumber)
    evtIDArray(sizeKeepRow - 1) = Sheets(1).Cells(rowNumb, "B").Value
    sizeKeepRow = sizeKeepRow + 1
Next

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

For Each strItem In evtIDArray
    If Not objDictionary.Exists(strItem) Then
        objDictionary.Add strItem, strItem
    End If
Next

intItems = objDictionary.Count - 1

ReDim evtIDArray(intItems)

k = 0

For Each strKey In objDictionary.Keys
    evtIDArray(k) = strKey
    k = k + 1
Next

'------------------------------------------------------------------------------
'This code section removes empty elements in the arrays
'------------------------------------------------------------------------------
Dim idsToKeep() As Variant
Dim a As Long
Dim j As Long

j = 0

ReDim idsToKeep(LBound(evtIDArray) To UBound(evtIDArray))

For a = LBound(evtIDArray) To UBound(evtIDArray)
    If Not IsEmpty(evtIDArray(a)) Then
        idsToKeep(j) = evtIDArray(a)
        j = j + 1
    End If
Next

ReDim Preserve idsToKeep(LBound(evtIDArray) To (j - 1))

'------------------------------------------------------------------------------
'Sort array descending
'mrexcel.com/forum/excel-questions/690718-visual-basic-applications-sort-array-numbers.htm
'------------------------------------------------------------------------------
Dim srtTemp As Variant
Dim g As Long
Dim h As Long

For g = LBound(idsToKeep) To UBound(idsToKeep)
    For h = g + 1 To UBound(idsToKeep)
        If idsToKeep(g) > idsToKeep(h) Then
            srtTemp = idsToKeep(h)
            idsToKeep(h) = idsToKeep(g)
            idsToKeep(g) = srtTemp
        End If
    Next h
Next g

'------------------------------------------------------------------------------
'This code section formats date/time correctly
'------------------------------------------------------------------------------
Sheets(1).Columns("A").NumberFormat = "mm/dd/yyyy hh:mm:ss"

'------------------------------------------------------------------------------
'Change Column Header Names
'------------------------------------------------------------------------------
Sheets(1).Cells(1, "A") = "Date/Time"
Sheets(1).Cells(1, "B") = "Evt ID"
Sheets(1).Cells(1, "C") = "Description"

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
Sheets(1).Rows(1).Font.Bold = True

'------------------------------------------------------------------------------
'Filtering is enabled on columns.
'------------------------------------------------------------------------------
Sheets(1).Range("A1").AutoFilter

'------------------------------------------------------------------------------
'Turn Off Wrap Text, Autofit Column Widths and Align All Cells to the Left
'------------------------------------------------------------------------------
With ActiveWorkbook.Sheets(1)
    .Columns.WrapText = False
    .Columns.HorizontalAlignment = xlLeft
    .Columns.AutoFit
End With

'--------------------------------------------------------------------------
'Create Sheets for each Unique Event ID in Array
'stackoverflow.com/questions/20697706/how-to-add-a-named-sheet-at-the-end-of-all-excel-sheets
'--------------------------------------------------------------------------
Dim eventID As Variant
Dim ws As Worksheet

For Each eventID In idsToKeep

    With ActiveWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.Name = eventID
    End With
    
    '------------------------------------------------------------------------------
    'Change Column Header Names
    '------------------------------------------------------------------------------
    Dim convNum As String
    
    convNum = eventID
    
    Sheets(convNum).Cells(1, "A") = "Date/Time"
    Sheets(convNum).Cells(1, "B") = "Evt ID"
    Sheets(convNum).Cells(1, "C") = "Description"

    '------------------------------------------------------------------------------
    'Copy Rows with same EventID to Tab for that EventID
    '------------------------------------------------------------------------------
    Dim newRow As Long
    
    newRow = 2
    
    For rowNumb = 2 To maxRows
        If Sheets(1).Cells(rowNumb, "B").Value = eventID Then
            Sheets(1).Rows(rowNumb).Copy
            Sheets(convNum).Rows(newRow).Insert Shift:=xlToRight
        End If
    Next

    '------------------------------------------------------------------------------
    'Find last row
    'www.rondebruin.nl/win/s9/win005.htm
    '------------------------------------------------------------------------------
    Dim lastRow As Long
    
    With ActiveSheet.UsedRange
        lastRow = .Rows(.Rows.Count).Row
    End With
    
    '------------------------------------------------------------------------------
    'Find last column
    'www.rondebruin.nl/win/s9/win005.htm
    '------------------------------------------------------------------------------
    Dim lastColumn As Long
    
    With ActiveSheet.UsedRange
        lastColumn = .Columns(.Columns.Count).Column
    End With

    '------------------------------------------------------------------------------
    'Delete Empty Columns
    '------------------------------------------------------------------------------
    Dim w As Long
    Dim d As Long
    Dim boolTest As Boolean
    
    boolTest = False
    
    For w = 1 To lastColumn
    
        For d = 1 To lastRow
        
            If Not IsEmpty(Sheets(convNum).Cells(d, w)) Then
                boolTest = False
                Exit For
            Else
                boolTest = True
            End If
            
        Next
        
        If boolTest = True Then
            Sheets(convNum).Columns(w).EntireColumn.Delete
        End If
    
    Next

    '------------------------------------------------------------------------------
    'This code section formats date/time correctly
    '------------------------------------------------------------------------------
    Sheets(convNum).Range("A1").NumberFormat = "mm/dd/yyyy hh:mm:ss"

    '------------------------------------------------------------------------------
    'Worksheet is sorted by column A, oldest to newest
    '------------------------------------------------------------------------------
    Sheets(convNum).Columns.Sort Key1:=Sheets(convNum).Range("A1"), Header:=xlYes

    '------------------------------------------------------------------------------
    'Freeze the first row.
    '------------------------------------------------------------------------------
    Sheets(convNum).Rows("2:2").Select
    ActiveWindow.FreezePanes = True

    '------------------------------------------------------------------------------
    'Bold the first row.
    '------------------------------------------------------------------------------
    Sheets(convNum).Rows(1).Font.Bold = True

    '------------------------------------------------------------------------------
    'Filtering is enabled on columns.
    '------------------------------------------------------------------------------
    Sheets(convNum).Range("A1").AutoFilter

    '------------------------------------------------------------------------------
    'Turn Off Wrap Text, Autofit Column Widths and Align All Cells to the Left
    '------------------------------------------------------------------------------
    With ActiveWorkbook.Sheets(convNum)
        .Columns.WrapText = False
        .Columns.HorizontalAlignment = xlLeft
        .Columns.AutoFit
    End With
    
Next

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
