Attribute VB_Name = "Module3"
Sub stock_data_2016()
        
Dim Ticker As String
Ticker = 0

Dim Open_Price As Double
Open_Price = 0

Dim Close_Price As Double
Close_Price = 0

Dim yearly_change As Double
yearly_change = 0

Dim percent_change As Double
percent_change = 0

Dim Summary_Table_Row As Long
Summary_Table_Row = 2
        
Open_Price = Cells(2, 3).Value
       
For i = 2 To 797711
            
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
             
Ticker = Cells(i, 1).Value
Close_Price = Cells(i, 6).Value
yearly_change = Close_Price - Open_Price
               
If Open_Price <> 0 Then
percent_change = (yearly_change / Open_Price)
End If
                
total_volume = total_volume + Cells(i, 7).Value
                          
Range("I" & Summary_Table_Row).Value = Ticker
Range("J" & Summary_Table_Row).Value = yearly_change
Range("K" & Summary_Table_Row).Value = percent_change
Range("L" & Summary_Table_Row).Value = total_volume
                            
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
                            
                           
Summary_Table_Row = Summary_Table_Row + 1

yearly_change = 0
Close_Price = 0
Open_Price = Cells(i + 1, 3).Value
percent_change = 0
total_volume = 0
                                                
Else
total_volume = total_volume + Cells(i, 7).Value
End If
Next i
End Sub



