VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Summary()
For Each ws In Worksheets

'Table Labels
ws.Cells(1, 11).Value = "Ticker"
ws.Cells(1, 12).Value = "Yearly Change"
ws.Cells(1, 13).Value = "Percent Change"
ws.Cells(1, 14).Value = "Total Stock Volume"
ws.Cells(2, 17).Value = "Greatest % Increase"
ws.Cells(3, 17).Value = "Greatest % Decrease"
ws.Cells(4, 17).Value = "Greatest Total Volume"
ws.Cells(1, 18).Value = "Ticker"
ws.Cells(1, 19).Value = "Value"





'Calculating Annual Totals

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim counter1 As Long
Dim BeginningRow As Long
Dim i As Long
Dim q As Long
volume = 0
counter1 = 2
BeginningRow = 2

For i = 2 To LastRow
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
VolumeRange = Range("G" & BeginningRow & ":G" & i)
ws.Cells(counter1, 14) = WorksheetFunction.Sum(VolumeRange)
ws.Cells(counter1, 12) = ws.Cells(i, 6).Value - ws.Cells(BeginningRow, 3).Value
    If ws.Cells(counter1, 12) > 0 Then
    ws.Cells(counter1, 12).Interior.ColorIndex = 4
    Else: ws.Cells(counter1, 12).Interior.ColorIndex = 3
    End If
ws.Cells(counter1, 13) = ws.Cells(counter1, 12) / ws.Cells(BeginningRow, 3).Value
ws.Cells(counter1, 13).Style = "Percent"
ws.Cells(counter1, 11) = ws.Cells(i, 1).Value
counter1 = counter1 + 1
BeginningRow = i + 1
volume = 0
End If

Next i

'Finding the greatest in the totals list

Dim GreatestPI As Double
Dim GreatestPD As Double
Dim GreatestVolume As LongLong
Dim GPItick As String
Dim GPDtick As String
Dim GVtick As String

GreatestPI = ws.Cells(2, 13).Value
GPItick = ws.Cells(2, 11).Value
GreatestPD = ws.Cells(2, 13).Value
GPDtick = ws.Cells(2, 11).Value
GreatestVolume = ws.Cells(2, 14).Value
GVtick = ws.Cells(2, 11).Value

For q = 3 To (counter1 - 1)
If ws.Cells(q, 13).Value > GreatestPI Then
GreatestPI = ws.Cells(q, 13).Value
GPItick = ws.Cells(q, 11).Value
End If
If ws.Cells(q, 13).Value < GreatestPD Then
GreatestPD = ws.Cells(q, 13).Value
GPDtick = ws.Cells(q, 11).Value
End If
If ws.Cells(q, 14).Value > GreatestVolume Then
GreatestVolume = ws.Cells(q, 14).Value
GVtick = ws.Cells(q, 11).Value
End If

Next q

'Fill greatest in totals table
ws.Cells(2, 19).Value = GreatestPI
ws.Cells(2, 19).Style = "Percent"
ws.Cells(3, 19).Value = GreatestPD
ws.Cells(3, 19).Style = "Percent"
ws.Cells(4, 19).Value = GreatestVolume
ws.Cells(2, 18).Value = GPItick
ws.Cells(3, 18).Value = GPDtick
ws.Cells(4, 18).Value = GVtick
ws.Range("K1:S1").EntireColumn.AutoFit

Next ws
MsgBox ("Summaries Added to All Worksheets")

End Sub

