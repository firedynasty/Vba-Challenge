Sub Leonard()

' LOOP THROUGH ALL SHEETS

' Declare Current as a worksheet object variable.
Dim ws As Worksheet
Dim Ticker As String
Dim LastRow As Double
Dim Summary_Table_Row As Integer
Dim StockWorkSheet As String
Dim Total_Stock_Volume As Double

Dim Yearly_Change As Double
Yearly_Change = 0
Dim Percent_Change As Double
Percent_Change = 0
Dim DFinRow As Double
Dim PerCalc As Long
Dim YrOp As Double
Dim YrCl As Double
Dim i As Long

For Each ws In Worksheets
    'Lastrow of worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Total_Stock_Volume = 0
    
    'StockWorkSheet = ws.Name
    'MsgBox StockWorkSheet


    Summary_Table_Row = 2
    
   
   


    For i = 2 To LastRow

    ' If the next row does not match the stock name then you want to do something but before that you want
    ' to be able to continue adding

 Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
 
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      Ticker = ws.Cells(i, 1).Value
      ws.Range("J" & Summary_Table_Row).Value = Ticker
      ws.Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
      
      Summary_Table_Row = Summary_Table_Row + 1
      Total_Stock_Volume = 0
    End If
  Next i
Next ws
End Sub
