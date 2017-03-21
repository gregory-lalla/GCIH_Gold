Attribute VB_Name = "Module1"
Option Explicit

Sub MFT_Entries_Standard_Format()

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
'This code section deletes a column
'--------------------------------------------------------------------------
Sheets(1).Range("A:G, K:K, O:O, Q:BZ").EntireColumn.Delete

'--------------------------------------------------------------------------
'Fix date/time values in specific columns
'--------------------------------------------------------------------------
Sheets(1).Columns("B:G").Value = Sheets(1).Columns("B:G").Value

'------------------------------------------------------------------------------
'Date/Time Column is formatted
'------------------------------------------------------------------------------
Sheets(1).Columns("B:G").NumberFormat = "mm/dd/yyyy hh:mm:ss"

'--------------------------------------------------------------------------
'This code section deletes rows with "NOFNRecord" or "Corrupt MFT Record"
'in Filename Column
'--------------------------------------------------------------------------
Dim p As Long

p = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    If (Cells(p, "A") = "NoFNRecord") Or (Cells(p, "A") = "Corrupt MFT Record") Then
        ActiveSheet.Rows(p).EntireRow.Delete
    Else
        p = p + 1
        ActiveCell.Offset(1, 0).Select
    End If

Loop

'------------------------------------------------------------------------------
'Find last row
'www.rondebruin.nl/win/s9/win005.htm
'------------------------------------------------------------------------------
Dim lastRow As Long

With ActiveSheet.UsedRange
    lastRow = .Rows(.Rows.Count).Row
End With

'------------------------------------------------------------------------------
'Change "/" in paths to "\"
'------------------------------------------------------------------------------
Dim findChar As Variant
Dim replaceChar As Variant
Dim s As Long

findChar = "/"
replaceChar = "\"

s = 2
Sheets(1).Cells(s, "A").Select

Do Until IsEmpty(ActiveCell)

    Sheets(1).Cells(s, "A").Replace what:=findChar, Replacement:=replaceChar, _
    LookAt:=xlPart, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False

    s = s + 1
    ActiveCell.Offset(1, 0).Select

Loop

'------------------------------------------------------------------------------
'In this code section, moving columns
'------------------------------------------------------------------------------
ActiveWorkbook.Sheets(1).Columns("E").Cut
ActiveWorkbook.Sheets(1).Columns("A").Insert Shift:=xlToRight

ActiveWorkbook.Sheets(1).Columns("B").Insert Shift:=xlToRight

'------------------------------------------------------------------------------
'Worksheet is sorted by column A, oldest to newest
'------------------------------------------------------------------------------
Sheets(1).Columns.Sort Key1:=ActiveWorkbook.Sheets(1).Range("A1"), Header:=xlYes

'------------------------------------------------------------------------------
'Fill Description column
'------------------------------------------------------------------------------
Dim z As Long

z = 2
Sheets(1).Cells(z, "A").Select

Do Until IsEmpty(ActiveCell)

    Cells(z, "B") = "Create Date (FN Info)"

    z = z + 1
    ActiveCell.Offset(1, 0).Select

Loop

'------------------------------------------------------------------------------
'Append Std Info Creation Date column and add other Std Info Timestamps
'------------------------------------------------------------------------------
Dim f As Long

f = 2
Sheets(1).Cells(f, "A").Select

Do Until IsEmpty(ActiveCell)

    Cells(f, "D") = "Std Info - Create: " & Cells(f, "D").Value & _
    ", Modify: " & Cells(f, "E").Value & ", Entry: " & Cells(f, "F").Value

    f = f + 1
    ActiveCell.Offset(1, 0).Select

Loop

'------------------------------------------------------------------------------
'Append FN Modifications Info Creation Date column and add FN Info Entry Timestamps
'------------------------------------------------------------------------------
Dim v As Long

v = 2
Sheets(1).Cells(v, "A").Select

Do Until IsEmpty(ActiveCell)

    Cells(v, "G") = "FN Info - Modify: " & Cells(v, "G").Value & ", Entry: " & Cells(v, "H").Value

    v = v + 1
    ActiveCell.Offset(1, 0).Select

Loop

'--------------------------------------------------------------------------
'This code section deletes a column
'--------------------------------------------------------------------------
Sheets(1).Range("E:F, H:H").EntireColumn.Delete

'--------------------------------------------------------------------------
'Insert columns
'--------------------------------------------------------------------------
ActiveWorkbook.Sheets(1).Columns("B").Insert Shift:=xlToRight
ActiveWorkbook.Sheets(1).Columns("C").Insert Shift:=xlToRight

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

For n = 2 To lastRow
    Cells(n, "B").Value = "N/A"
    Cells(n, "C").Value = hostName
    Cells(n, "H").Value = "MFT Entry"
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

