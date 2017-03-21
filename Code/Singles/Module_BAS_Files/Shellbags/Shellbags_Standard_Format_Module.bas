Attribute VB_Name = "Module1"
Option Explicit

Sub Shellbag_File_Standard_Format()

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
'This code section deletes specific columns
'--------------------------------------------------------------------------
Sheets(1).Range("A:D,F:H,M:O,R:R").EntireColumn.Delete

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
Dim x As Long
Dim CellContent As Variant
Dim arrayFields As Long
Dim newOutput As String
Dim nextField As String

y = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    newOutput = ""
    x = 0
    txt = Cells(y, 1).Value
    CellContent = Split(txt, "\")
    arrayFields = UBound(CellContent)
    
    If arrayFields <> 0 Then
    
        For i = 0 To arrayFields
            x = x + 1
            If x > arrayFields Then
                x = x
            Else
                nextField = Trim(CellContent(x))
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
Dim b As Long

b = 2
Range("A2").Select

Do Until IsEmpty(ActiveCell)

    If IsEmpty(Cells(b, "D")) Then
        ActiveSheet.Rows(b).EntireRow.Delete
    Else
        b = b + 1
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
Dim lastRow As Long

With ActiveSheet.UsedRange
    lastRow = .Rows(.Rows.Count).Row
End With

'------------------------------------------------------------------------------
'Inserts Account Name, Computer Name and Artifact Name in each cell of each column
'------------------------------------------------------------------------------
Dim n As Long

For n = 2 To lastRow
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

