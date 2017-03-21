Attribute VB_Name = "Module1"
Option Explicit

Sub Binary_Hex2Ascii_Conversion()

'------------------------------------------------------------------------------
'Make VBA Code Run Faster
'------------------------------------------------------------------------------
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.EnableEvents = False

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

'------------------------------------------------------------------------------
'Find binary data column
'------------------------------------------------------------------------------
Dim eventBin As Variant

eventBin = Application.Match("ns?:Binary", Sheets(1).Rows(1), 0)

'------------------------------------------------------------------------------
'Convert Hex to ASCII
'------------------------------------------------------------------------------
Dim fnd As Variant
Dim rplc As Variant
Dim binData() As String
Dim convertData As String
Dim l As Long
Dim z As Long
Dim y As Long
Dim c As Long

l = 2
fnd = "00"
rplc = ""

'------------------------------------------------------------------------------
'Replace Unicode 2 bytes with ASCII 1 byte
'------------------------------------------------------------------------------
Sheets(1).Cells(l, eventBin).Select

Do Until IsEmpty(ActiveCell)

    Sheets(1).Cells(l, eventBin).Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False

    l = l + 1
    ActiveCell.Offset(1, 0).Select

Loop

'------------------------------------------------------------------------------
'Fill blank cells with a hyphon so Hex to Ascii does not stop on Binary Cells
'filled with all zeros
'------------------------------------------------------------------------------
For d = 1 To lastRow
    If IsEmpty(ActiveSheet.Cells(d, eventBin).Value) Then
        ActiveSheet.Cells(d, eventBin).Value = "-"
    End If
Next

'------------------------------------------------------------------------------
'Starts Conversion of Hex to ASCII
'------------------------------------------------------------------------------
z = 2
Sheets(1).Cells(1, eventBin).Select

Do Until IsEmpty(ActiveCell)

convertData = ""
y = 0

ReDim binData(Len(Cells(z, eventBin).Value))

For c = 0 To (Len(Cells(z, eventBin).Value) / 2)

    binData(c) = Chr(Val("&H" & Mid(Cells(z, eventBin).Value, y + 1, 2)))
    convertData = convertData & binData(c)
    y = y + 2

Next

Cells(z, eventBin) = convertData

ActiveCell.Offset(1, 0).Select

z = z + 1

Loop

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

