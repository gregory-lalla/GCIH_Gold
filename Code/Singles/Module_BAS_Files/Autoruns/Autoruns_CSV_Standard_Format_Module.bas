Attribute VB_Name = "Module1"
Option Explicit

Sub Autoruns_CSV_Standard_Format()

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
'Format date/times in specific columns
'--------------------------------------------------------------------------
Sheets(1).Columns("A").NumberFormat = "mm/dd/yyyy hh:mm:ss"

'--------------------------------------------------------------------------
'This code section deletes specific columns
'--------------------------------------------------------------------------
Sheets(1).Range("D:D,E:E,H:I,K:K").EntireColumn.Delete

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

'--------------------------------------------------------------------------
'Combine Columns
'--------------------------------------------------------------------------
For d = 2 To lastRow
    Cells(d, "B").Value = Cells(d, "B") & " - " & Cells(d, "C")
    Cells(d, "G").Value = "Launch String: " & Cells(d, "G")
Next

'--------------------------------------------------------------------------
'This code section deletes specific columns
'--------------------------------------------------------------------------
Sheets(1).Columns("C").EntireColumn.Delete

'--------------------------------------------------------------------------
'Insert Columns
'--------------------------------------------------------------------------
Sheets(1).Columns("C").Cut
Sheets(1).Columns("B").Insert Shift:=xlToRight

Sheets(1).Columns("C").Insert Shift:=xlToRight

Sheets(1).Columns("D").Cut
Sheets(1).Columns("F").Insert Shift:=xlToRight

'------------------------------------------------------------------------------
'Inserts Account Name, Computer Name and Artifact Name in each cell of each column
'------------------------------------------------------------------------------
Dim n As Long

For n = 2 To lastRow
    Cells(n, "C").Value = hostName
    Cells(n, "H").Value = "Autoruns"
Next

'------------------------------------------------------------------------------
'Worksheet is sorted by column A, oldest to newest
'------------------------------------------------------------------------------
Sheets(1).Columns.Sort Key1:=Sheets(1).Range("A1"), Header:=xlYes

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
