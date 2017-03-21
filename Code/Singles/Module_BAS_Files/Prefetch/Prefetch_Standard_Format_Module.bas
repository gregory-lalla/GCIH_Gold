Attribute VB_Name = "Module1"
Option Explicit

Sub Prefetch_File_Standard_Format()

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

'------------------------------------------------------------------------------
'Fix Date Column
'------------------------------------------------------------------------------
Dim findChar As Variant
Dim replaceChar As Variant
Dim fullDate As String
Dim splitDate() As String
Dim correctDate As Date
Dim m As Long

findChar = "  "
replaceChar = " "

m = 2
Sheets(1).Cells(m, "D").Select

Do Until IsEmpty(ActiveCell)

    Sheets(1).Cells(m, "D").Replace what:=findChar, Replacement:=replaceChar, _
    LookAt:=xlPart, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
    
    fullDate = ActiveCell.Value
    splitDate = Split(fullDate, " ")
    correctDate = splitDate(1) & " " & splitDate(2) & " " & splitDate(4) & " " & splitDate(3)
    Sheets(1).Cells(m, "D").Value = correctDate

    m = m + 1
    ActiveCell.Offset(1, 0).Select

Loop

'------------------------------------------------------------------------------
'Worksheet is sorted by column D, oldest to newest
'------------------------------------------------------------------------------
Sheets(1).Columns.Sort Key1:=Sheets(1).Range("D1"), Header:=xlYes

'------------------------------------------------------------------------------
'In this code section, moving columns
'------------------------------------------------------------------------------
Sheets(1).Columns("D").Cut
Sheets(1).Columns("A").Insert Shift:=xlToRight

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
'Date/Time Column is formatted
'------------------------------------------------------------------------------
Sheets(1).Columns("A").NumberFormat = "mm/dd/yyyy hh:mm:ss"

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
    Cells(n, "F").Value = "Number of Time Run: " & Cells(n, "F")
    Cells(n, "H").Value = "Prefetch Entry"
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
'This code section picks dates to filter on
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

