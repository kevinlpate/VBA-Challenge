Attribute VB_Name = "RibbonX_Code"
Sub StockHomework()


Dim ticker As String
Dim volume As Variant
volume = 0
Dim Summary_Table_Row As Double
Summary_Table_Row = 2
Last_Row = 70926
Dim Yearly_Change As Double

Dim Start_Price As Double
Dim End_Price As Double

Range("I1").Value = "ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Volume"
Range("M1").Value = "Opening Price"
Range("N1").Value = "Closing Price"


For i = 2 To Last_Row

   If Cells(i, 2).Value = 20160101 Then
   Start_Price = Cells(i, 6).Value
   
   Range("M" & Summary_Table_Row).Value = Start_Price
   
   End If
   
   If Cells(i, 2).Value = 20161230 Then
   End_Price = Cells(i, 6).Value
   
   Range("N" & Summary_Table_Row).Value = End_Price

   End If
   
   Yearly_Change = End_Price - Start_Price
   
Range("J" & Summary_Table_Row).Value = Yearly_Change

If Range("J" & Summary_Table_Row).Value > 0 Then

    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

    Else: Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    
    End If
    
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
       ticker = Cells(i, 1).Value
       
       volume = volume + Cells(i, 7).Value
       
       Range("I" & Summary_Table_Row).Value = ticker
       
       Range("L" & Summary_Table_Row).Value = volume
       
       Summary_Table_Row = Summary_Table_Row + 1
       
       volume = 0
Else
   
   Vol = Vol + Cells(i, 7)
End If
Next i

If ActiveSheet.Index = Worksheets.Count Then

    Worksheets(1).Activate

    Else: ActiveSheet.Next.Activate
    
End If

End Sub

