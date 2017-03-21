Attribute VB_Name = "Module1"
Option Explicit

Sub IE_History_Output_Standard_Format()

'------------------------------------------------------------------------------
'Make VBA Code Run Faster
'------------------------------------------------------------------------------
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.EnableEvents = False

'--------------------------------------------------------------------------
'This Code section prompts the analyst to enter the user profile name and
'computer name
'--------------------------------------------------------------------------
Dim hostName As String
Dim profileName As String

profileName = Application.InputBox("Enter the User Name associated with this file")
hostName = Application.InputBox("Enter the Computer Name associated with this file")

'--------------------------------------------------------------------------
'This code section deletes the first two rows
'--------------------------------------------------------------------------
Sheets(1).Rows(1).EntireRow.Delete
Sheets(1).Rows(1).EntireRow.Delete

'--------------------------------------------------------------------------
'This code section deletes specific columns
'--------------------------------------------------------------------------
Sheets(1).Range("C:C,F:F").EntireColumn.Delete

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

''--------------------------------------------------------------------------
''This code section prompts the user to filter events by date. In most cases,
''Event Viewer logs should be filtered by date on export. Uncomment this
''section if this feature is desired.
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
'        If Cells(u, "C") < startDate Or Cells(u, "C") > endDate + 1 Then
'            ActiveSheet.Rows(u).EntireRow.Delete
'        Else
'            u = u + 1
'            ActiveCell.Offset(1, 0).Select
'        End If
'
'    Loop
'
'End If

'--------------------------------------------------------------------------
'Insert Columns
'--------------------------------------------------------------------------
Sheets(1).Columns("C").Cut
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
'Inserts Account Name, Computer Name and Artifact Name in each cell of each column
'------------------------------------------------------------------------------
Dim n As Long

For n = 2 To lastRow
    Cells(n, "B").Value = profileName
    Cells(n, "C").Value = hostName
    Cells(n, "H").Value = "IE History"
Next

'------------------------------------------------------------------------------
'Date/Time Column is formatted
'------------------------------------------------------------------------------
Sheets(1).Columns("A").NumberFormat = "mm/dd/yyyy hh:mm:ss"

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
