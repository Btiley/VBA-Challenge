Attribute VB_Name = "Module1"
Sub stock_counter():

    Dim lastRow As Long
    Dim start As Long
    Dim Opening_Value As Double
    Dim Closing_Value As Double
    Dim ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim totalVolume As Double
    

Range("I1").value = "Ticker"
Range("J1").value = "Yearly Change"
Range("K1").value = "Percent_Change"
Range("L1").value = "Total_Stock"
  
  lastRow = Cells(Rows.Count, "A").End(xlUp).Row
 
    
  start = 2
  nr = 0
  totalVolume = 0
       
For Row = 2 To lastRow:
    
    totalVolume = totalVolume + Cells(Row, 7).value

If Cells(Row, 1).value <> Cells(Row + 1, 1).value Then
 
 'Ticker Value
        ticker = Cells(start, 1).value

        Range("I" & 2 + nr).value = ticker
    
'Open/Close Value
    Opening_Value = Cells(start, 3).value
 
    Closing_Value = Cells(Row, 6).value
 
'Yearly Change Calc
    Yearly_Change = Closing_Value - Opening_Value
    
    Range("J" & 2 + nr).value = Yearly_Change
'Total Stock Calc

'Overflow might be caused by dividing 0?
    Range("L" & 2 + nr).value = totalVolume
If Yearly_Change Or Opening_Value = 0 Then
    
        Percent_Change = 0
    
Else
        Percent_Change = Yearly_Change / Opening_Value
        
        Range("K" & 2 + nr).value = FormatPercent(Percent_Change, 2)
End If
    

If Range("K" & 2 + nr).value < 0 Then
        Cells(Row, 11).Interior.ColorIndex = 3
    
Else
        Cells(Row, 11).Interior.ColorIndex = 4
    
End If
    
    start = Row + 1
    nr = nr + 1
    totalVolume = 0
    
    End If
    
    
    Next Row
    
End Sub


