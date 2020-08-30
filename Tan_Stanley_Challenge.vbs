Sub Tester()


Dim ws As Worksheet
For Each ws In Worksheets


' Loop through the cells and if the value is the highest at the end print the i'
last_row_percentage = ws.Cells(Rows.Count, 12).End(xlUp).Row

Dim m As Double

Dim n As Double
Dim p As String
'tester = Range("B1:B6")
' I don't need that range
m = 0
For n = 2 To last_row_percentage
If ws.Cells(n, 13).Value > m Then
m = ws.Cells(n, 13).Value
p = ws.Cells(n, 10).Value

' get the name of the stock too'
End If
Next

ws.Range("P2").Value = "Max percentage"
ws.Range("Q2").Value = p
ws.Range("R2").Value = m
ws.Range("R2").NumberFormat = "0.00%"

'make the cells in O to reveal the max '
' and write out the name for it to

' Loop through the cells and if the value is the highest at the end print the i'

last_row_percentage = ws.Cells(Rows.Count, 12).End(xlUp).Row

Dim q As Double

Dim r As Double
Dim s As String
'tester = Range("B1:B6")
' I don't need that range
q = 0
For r = 2 To last_row_percentage
If ws.Cells(r, 13).Value < q Then
q = ws.Cells(r, 13).Value
s = ws.Cells(r, 10).Value

' get the name of the stock too'
End If
Next

ws.Range("P3").Value = "Min percentage"
ws.Range("Q3").Value = s
ws.Range("R3").Value = q
ws.Range("R3").NumberFormat = "0.00%"

'make the cells in O to reveal the max '
' and write out the name for it to

' Loop through the cells and if the value is the highest at the end print the i'

last_row_volume = ws.Cells(Rows.Count, 10).End(xlUp).Row

Dim tt As Double

Dim u As Double
Dim vv As String
'tester = Range("B1:B6")
' I don't need that range
tt = 0

For u = 2 To last_row_volume

If ws.Cells(u, 14).Value > tt Then

tt = ws.Cells(u, 14).Value

vv = ws.Cells(u, 10).Value

' get the name of the stock too'
End If
Next

ws.Range("P4").Value = "Most total volume"
ws.Range("Q4").Value = vv
ws.Range("R4").Value = tt

'make the cells in O to reveal the max '
' and write out the name for it to

ws.Range("P2:R4").Columns.AutoFit
Next ws
    
End Sub

