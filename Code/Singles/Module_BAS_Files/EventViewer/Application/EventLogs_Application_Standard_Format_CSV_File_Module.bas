Attribute VB_Name = "Module1"
Option Explicit

Sub EventLogs_Application_Standard_Format()

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

'--------------------------------------------------------------------------
'This Code section prompts the analyst to enter the computer name
'--------------------------------------------------------------------------
Dim hostName As String

hostName = Application.InputBox("Enter the Computer Name associated with this file")

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
Sheets(1).Range("A:A,C:C,E:E").EntireColumn.Delete

'------------------------------------------------------------------------------
'In this code section, moving and adding columns
'------------------------------------------------------------------------------
Sheets(1).Columns("B").Insert Shift:=xlToRight
Sheets(1).Columns("C").Insert Shift:=xlToRight

Sheets(1).Columns("D").Cut
Sheets(1).Columns("F").Insert Shift:=xlToRight

'------------------------------------------------------------------------------
'This code section formats date/time correctly
'------------------------------------------------------------------------------
Sheets(1).Columns("A").NumberFormat = "mm/dd/yyyy hh:mm:ss"

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
'Combine Column, Inserts Account Name, Computer Name and Artifact Name in each
'cell of each column
'------------------------------------------------------------------------------
Dim f As Long

For f = 2 To lastRow
    Cells(f, "B").Value = "N/A"
    Cells(f, "C").Value = hostName
    Cells(f, "E") = "Evt ID: " & Cells(f, "E")
'    Cells(f, "H").Value = "Application Event Log"
Next

'------------------------------------------------------------------------------
'Change Column Header Names
'------------------------------------------------------------------------------
Sheets(1).Cells(1, "A") = "Date/Time"
Sheets(1).Cells(1, "B") = "Account"
Sheets(1).Cells(1, "C") = "Computer"
Sheets(1).Cells(1, "D") = "Description"
Sheets(1).Cells(1, "E") = "Details"
'Sheets(1).Cells(1, "F") = "Properties"
'Sheets(1).Cells(1, "G") = "Miscellaneous"
'Sheets(1).Cells(1, "H") = "Artifact"

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
