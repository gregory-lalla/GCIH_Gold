Attribute VB_Name = "Module1"
Option Explicit

Sub Apache_Access_Log_Standard_Format()

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
Sheets(1).Range("B:B,C:C,H:H,J:J").EntireColumn.Delete

'--------------------------------------------------------------------------
'Insert Header Row
'--------------------------------------------------------------------------
Sheets(1).Rows(1).Insert Shift:=xlToRight

'------------------------------------------------------------------------------
'Adjust Date/Time Column according to timezone value
'Split code from Alex K.
'http://stackoverflow.com/questions/13195583/split-string-into-array-of-characters
'------------------------------------------------------------------------------
Dim plusMinus As String
Dim timeZone As Long
Dim zoneHour As Long
Dim zoneMinute As Long
Dim zoneParts() As String
Dim e As Long

e = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    timeZone = Cells(e, "C")
    zoneParts = Split(StrConv(timeZone, vbUnicode), Chr$(0))
    ReDim Preserve zoneParts(UBound(zoneParts) - 1)
    zoneHour = zoneParts(1)
    zoneMinute = zoneParts(2) & zoneParts(3)
    plusMinus = zoneParts(0)
    
    If plusMinus = "-" Then
        Cells(e, "B") = Cells(e, "B") + (zoneHour) / 24
    Else
        Cells(e, "B") = Cells(e, "B") - (zoneHour) / 24
    End If
    
    e = e + 1
    ActiveCell.Offset(1, 0).Select

Loop

'--------------------------------------------------------------------------
'This code section combines columns
'--------------------------------------------------------------------------
Dim f As Long

f = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    Cells(f, "A") = "Client IP: " & Cells(f, "A")
    Cells(f, "F") = "Status Code: " & Cells(f, "F")
    
    f = f + 1
    ActiveCell.Offset(1, 0).Select

Loop

'--------------------------------------------------------------------------
'This code section deletes a column
'--------------------------------------------------------------------------
Sheets(1).Range("C:C").EntireColumn.Delete

'--------------------------------------------------------------------------
'Insert Columns
'--------------------------------------------------------------------------
Sheets(1).Columns("B").Insert Shift:=xlToRight
Sheets(1).Columns("C").Insert Shift:=xlToRight

Sheets(1).Columns("D").Cut
Sheets(1).Columns("A").Insert Shift:=xlToRight

Sheets(1).Columns("B").Cut
Sheets(1).Columns("E").Insert Shift:=xlToRight

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
    Cells(n, "H").Value = "Apache Access Entry"
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
