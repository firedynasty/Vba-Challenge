' I want to add the colors


Sub Homework_2_VBA()

Dim ws As Worksheet

Dim Ticker As String

Dim Summary_Row As Integer

Dim Total_Stock_Volume As Double

' Declare beginning stock value
 
Dim Beg_stock_value As Double
 
Dim End_stock_value As Double

' Variable for keeping track of row for Beg_stock_value

Dim Beg_var As Long

' Declare variable for percent change

Dim percentage_change As Double

Dim total As Double

Dim i As Long

For Each ws In Worksheets
    'Lastrow of worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' Declare the stock volume
    Total_Stock_Volume = 0
    
    'StockWorkSheet = ws.Name
    'MsgBox StockWorkSheet

    Summary_Row = 2

    Beg_var = 2

 


    'Declare beginning stock value

    For i = 2 To LastRow

    ' If the next row does not match the stock name then you want to do something but before that you want
    ' to be able to continue adding

      'The stock volume gets added in this loop
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
      
      End_stock_value = Cells(i, 6).Value

      'how would I get the beginning stock value?
      ' I would need to declare it from the very beginning'

       Beg_stock_value = Cells(Beg_var, 6).Value

      'When the stock ticker is different in value, then it breaks the addition

      ' How do you check the zero value?

      'add two columns first then mess around with it later


      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          Ticker = ws.Cells(i, 1).Value

          total = (End_stock_value - Beg_stock_value)
          ws.Range("J" & Summary_Row).Value = Ticker
          ws.Range("K" & Summary_Row).Value = total
          ws.Range("L" & Summary_Row).Value = Beg_stock_value
          ws.Range("M" & Summary_Row).Value = ""
          ws.Range("N" & Summary_Row).Value = Total_Stock_Volume

         'percentage_change = total / Beg_stock_value
        'percentage_change = ((End_stock_value - Beg_stock_value) / (Beg_stock_value))
          'ws.Range("L" & Summary_Row).Value = percentage_change
        Summary_Row = Summary_Row + 1
        Total_Stock_Volume = 0
        Beg_var = i + 1
    End If
Next i

last_row_summary = ws.Cells(Rows.Count, 10).End(xlUp).Row

Dim j As Double


For j = 2 To last_row_summary

If ws.Cells(j, 12) <> 0 Then
  ws.Range("M" & j).Value = ws.Range("K" & j) / ws.Range("L" & j)
' get the name of the stock too'
End If
Next

'get last row of stock_summary_table
'from the 2nd cell then go all the way to the bottom to format to percentage



last_row_summary = ws.Cells(Rows.Count, 13).End(xlUp).Row

For k = 2 To last_row_summary
''
ws.Cells(k, 13).NumberFormat = "0.00%"

Next k



ws.Range("J1:M" & last_row_summary).Columns.AutoFit


For l = 2 To last_row_summary

If ws.Cells(l, 11).Value > 0 Then
  ws.Cells(l, 11).Interior.ColorIndex = 4
Else
 ws.Cells(l, 11).Interior.ColorIndex = 3
End If

Next

    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Beg_value"
    ws.Range("M1").Value = "Percentage change"
    ws.Range("N1").Value = "Total Volume"



Next ws
End Sub






