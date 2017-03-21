Attribute VB_Name = "Module1"
Option Explicit

Sub Bro_Conn_Log_Standard_Format()

'------------------------------------------------------------------------------
'Make VBA Code Run Faster
'------------------------------------------------------------------------------
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.EnableEvents = False

'--------------------------------------------------------------------------
'This Code section prompts the analyst to enter the computer name
'--------------------------------------------------------------------------
Dim hostName As String

hostName = Application.InputBox("Enter the Computer Name associated with this file")

'--------------------------------------------------------------------------
'This code section deletes the first three rows
'--------------------------------------------------------------------------
Sheets(1).Rows(1).EntireRow.Delete
Sheets(1).Rows(1).EntireRow.Delete
Sheets(1).Rows(1).EntireRow.Delete
Sheets(1).Rows(1).EntireRow.Delete
Sheets(1).Rows(1).EntireRow.Delete
Sheets(1).Rows(1).EntireRow.Delete
Sheets(1).Rows(2).EntireRow.Delete

'--------------------------------------------------------------------------
'This code section deletes the first cell and shifts contents of first row left
'--------------------------------------------------------------------------
Sheets(1).Cells(1, 1).Delete Shift:=xlToLeft

'--------------------------------------------------------------------------
'This code section convert epoch time to human readable time
'Calculation by Oorang
'http://stackoverflow.com/questions/2259324/how-to-get-seconds-since-epoch-1-1-1970-in-vba
'--------------------------------------------------------------------------
Dim e As Long
Dim cellValue As Variant
Dim humanTime As Date

e = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    If IsNumeric(Cells(e, "A")) Then
        cellValue = Cells(e, "A").Value
        humanTime = cellValue / 86400# + #1/1/1970#
        Cells(e, "A") = humanTime
    End If
    
    e = e + 1
    ActiveCell.Offset(1, 0).Select
    
Loop

'--------------------------------------------------------------------------
'This code section to delete extra headers
'--------------------------------------------------------------------------
Dim c As Long

c = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    If Not IsDate(ActiveCell) Then
        Sheets(1).Rows(c).EntireRow.Delete
    Else
        c = c + 1
        ActiveCell.Offset(1, 0).Select
    End If

Loop

'--------------------------------------------------------------------------
'This code section deletes a column
'--------------------------------------------------------------------------
Sheets(1).Range("G:G,I:U").EntireColumn.Delete

'--------------------------------------------------------------------------
'This code section deletes empty rows
'--------------------------------------------------------------------------
Dim p As Long

p = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    If IsEmpty(Cells(p, "A")) Then
        ActiveSheet.Rows(p).EntireRow.Delete
    Else
        p = p + 1
        ActiveCell.Offset(1, 0).Select
    End If

Loop

'--------------------------------------------------------------------------
'This code section combines columns
'--------------------------------------------------------------------------
Dim f As Long

f = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    Cells(f, "C") = "Orig IP: " & Cells(f, "C") & " | Orig Prt: " & Cells(f, "D")
    Cells(f, "E") = "Resp IP: " & Cells(f, "E") & " | Resp Prt: " & Cells(f, "F")
    
    f = f + 1
    ActiveCell.Offset(1, 0).Select

Loop

'--------------------------------------------------------------------------
'This code section deletes a column
'--------------------------------------------------------------------------
Sheets(1).Range("D:D,F:F").EntireColumn.Delete

'------------------------------------------------------------------------------
'Date/Time Column is formatted
'------------------------------------------------------------------------------
Sheets(1).Columns("A").NumberFormat = "mm/dd/yyyy hh:mm:ss"

'--------------------------------------------------------------------------
'Insert Columns
'--------------------------------------------------------------------------
Sheets(1).Columns("D").Cut
Sheets(1).Columns("B").Insert Shift:=xlToRight

Sheets(1).Columns("E").Cut
Sheets(1).Columns("C").Insert Shift:=xlToRight

Sheets(1).Columns("E").Cut
Sheets(1).Columns("D").Insert Shift:=xlToRight

Sheets(1).Columns("B").Insert Shift:=xlToRight
Sheets(1).Columns("C").Insert Shift:=xlToRight

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
    Cells(n, "G") = "UID: " & Cells(n, "G")
    Cells(n, "H").Value = "Bro Conn Log"
Next

'------------------------------------------------------------------------------
'Freeze the first row.
'Code by Dannnit
'http://stackoverflow.com/questions/3232920/how-can-i-programmatically-freeze-the-top-row-of-an-excel-worksheet-in-excel-200
'------------------------------------------------------------------------------
Dim r As Range
Set r = ActiveCell
Range("A2").Select
With ActiveWindow
    .FreezePanes = False
    .ScrollRow = 1
    .ScrollColumn = 1
    .FreezePanes = True
    .ScrollRow = r.Row
End With
r.Select

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
