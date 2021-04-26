Attribute VB_Name = "Module11"
Sub loop_workbooks_for_loop()

Dim a As Integer
Dim ws_num As Integer

Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet
ws_num = ThisWorkbook.Worksheets.Count

For a = 1 To ws_num
ThisWorkbook.Worksheets(a).Activate
    
Dim ticker As String

Dim startPrice As Long
Dim endPrice As Long
Dim yearly_change As Long

Dim percent_change As Long

Dim stockVolume As LongLong

Dim summary_row As Integer
summary_row = 2

Dim sht As Worksheet
Dim LastRow As Long
Set sht = ActiveSheet
LastRow = sht.Range("A2").CurrentRegion.Rows.Count


startPrice = Cells(2, 3).Value

For i = 2 To LastRow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

ticker = Cells(i, 1).Value
Range("j" & summary_row).Value = ticker

endPrice = Cells(i, 6).Value
yearly_change = startPrice - endPrice
Range("k" & summary_row).Value = yearly_change


If startPrice = 0 Then
Range("l" & summary_row).Value = "Null"
Else
percent_change = (yearly_change / startPrice) * 100
Range("l" & summary_row).Value = percent_change
End If
 
 If Range("l" & summary_row).Value >= 0 Then
 Range("l" & summary_row).Interior.Color = vbGreen
 Else
 Range("l" & summary_row).Interior.Color = vbRed
 End If
 

stockVolume = stockVolume + Cells(i, 7).Value
Range("m" & summary_row).Value = stockVolume

summary_row = summary_row + 1
startPrice = Cells(i + 1, 3).Value
stockVolume = 0

Else

stockVolume = stockVolume + Cells(i, 7).Value

End If

Next i

Next

starting_ws.Activate

End Sub
