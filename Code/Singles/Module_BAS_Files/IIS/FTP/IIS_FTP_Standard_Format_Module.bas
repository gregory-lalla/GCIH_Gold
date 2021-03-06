Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text

Sub IIS_FTP_Log_Standard_Format()

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

'--------------------------------------------------------------------------
'This code section deletes the first cell and shifts contents of first row left
'--------------------------------------------------------------------------
Sheets(1).Cells(1, 1).Delete Shift:=xlToLeft

'--------------------------------------------------------------------------
'This code section to delete extra headers
'--------------------------------------------------------------------------
'Use a IsDate Loop to delete remaining headers in file
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
'This code section combines columns
'--------------------------------------------------------------------------
Dim f As Long

f = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    Cells(f, "A") = Cells(f, "A") + Cells(f, "B")
    Cells(f, "E") = "Server IP: " & Cells(f, "E") & " | Server Prt: " & Cells(f, "F")
    Cells(f, "G") = Cells(f, "G") & ": " & Cells(f, "H")
    
    f = f + 1
    ActiveCell.Offset(1, 0).Select

Loop

'------------------------------------------------------------------------------
'Date/Time Column is formatted
'------------------------------------------------------------------------------
Sheets(1).Columns("A").NumberFormat = "mm/dd/yyyy hh:mm:ss"

'--------------------------------------------------------------------------
'This code section deletes a column
'--------------------------------------------------------------------------
Sheets(1).Range("B:B,D:D,F:F,H:H,J:L").EntireColumn.Delete

'--------------------------------------------------------------------------
'Insert Columns
'--------------------------------------------------------------------------
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
    Cells(n, "D") = "Client IP: " & Cells(n, "D")
    Cells(n, "G") = "Status Code: " & Cells(n, "G")
    Cells(n, "H").Value = "IIS FTP Entry"
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
