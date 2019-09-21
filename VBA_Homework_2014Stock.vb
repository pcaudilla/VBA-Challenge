Sub Multi_Stock_2014()

Dim Ticker As String

Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock As Double
Total_Stock = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim OpeningDate As Long
OpeningDate = "20140101"

Dim ClosingDate As Long
ClosingDate = "20141231"

Dim OpeningPrice As Double
Dim ClosingPrice As Double

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Range("I1:L1").WrapText = True


For i = 2 To lastrow

   If Cells(i, 2).Value = OpeningDate Then
   
     OpeningPrice = Cells(i, 3).Value
   
   End If
   
   
   If Cells(i, 2).Value = ClosingDate Then
     
     ClosingPrice = Cells(i, 6).Value

   End If
   
   
   
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
     Ticker = Cells(i, 1).Value
     Yearly_Change = ClosingPrice - OpeningPrice
     Percent_Change = Yearly_Change / OpeningPrice
     Total_Stock = Total_Stock + Cells(i, 7).Value
     
     Range("I" & Summary_Table_Row).Value = Ticker
     Range("J" & Summary_Table_Row).Value = Yearly_Change
     Range("K" & Summary_Table_Row).Value = FormatPercent(Percent_Change)
     Range("L" & Summary_Table_Row).Value = Total_Stock
     
     Summary_Table_Row = Summary_Table_Row + 1
     Total_Stock = 0
   
   Else
   
     Total_Stock = Total_Stock + Cells(i, 7).Value
     
   End If
   


   If Range("J" & Summary_Table_Row).Value > 0 Then
   
     Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
   
   ElseIf Range("J" & Summary_Table_Row).Value <= 0 Then
     
     Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
     

   End If
     
   
  Next i



End Sub
