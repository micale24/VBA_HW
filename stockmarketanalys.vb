Sub stock_market()

 
Dim ws As Worksheet
For Each ws In Worksheets

Dim ticker As String
Dim stock_vol, beg_stock_yr, end_stock_yr, stock_delta, percent_change As Double
Dim j, k, max, total_vol, min As Integer


Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Lastrow_percent = ws.Cells(Rows.Count, 11).End(xlUp).Row
j = 2
stock_vol = 0
percent_change = 0

'Summary Table Titles
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


'Loop to add up the stock volume and chart the ticker symbol
For i = 2 To Lastrow
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        stock_vol = stock_vol + ws.Cells(i, 7).Value

    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        stock_vol = stock_vol + ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1).Value
        ws.Cells(j, 9).Value = ticker
        ws.Cells(j, 12).Value = stock_vol
        j = j + 1
        stock_vol = 0
    End If
Next i

'Loop for the Percent change and Yearly change and color of interior cells
k = 2
j = 2
For i = 2 To Lastrow
    If Right(ws.Cells(i, 2).Value, 4) = "0101" Then
        beg_stock_yr = ws.Cells(i, 3).Value

    ElseIf Right(ws.Cells(i, 2).Value, 4) = "1230" Then
        end_stock_yr = ws.Cells(i, 6).Value
        'Calculation yearly change
         yearly_change = (end_stock_yr - beg_stock_yr)
         ws.Cells(j, 10).Value = yearly_change
         j = j + 1
         percent_change = (yearly_change / beg_stock_yr)
         ws.Range("K2:K" & Lastrow).EntireColumn.NumberFormat = "00.00%"
         ws.Cells(k, 11).Value = percent_change
          If ws.Cells(k, 10).Value > 0 Then
                ws.Cells(k, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(k, 10).Value < 0 Then
                ws.Cells(k, 10).Interior.ColorIndex = 3
            End If
         k = k + 1
         

    ElseIf Right(ws.Cells(i, 2).Value, 4) = "1231" Then
         end_stock_yr = ws.Cells(i, 6).Value
         yearly_change = (end_stock_yr - beg_stock_yr)
         ws.Cells(j, 10).Value = yearly_change
         j = j + 1
         percent_change = (yearly_change / beg_stock_yr)
         ws.Range("K2:K" & Lastrow).EntireColumn.NumberFormat = "00.00%"
         ws.Cells(k, 11).Value = percent_change
         If ws.Cells(k, 10).Value > 0 Then
                ws.Cells(k, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(k, 10).Value < 0 Then
                ws.Cells(k, 10).Interior.ColorIndex = 3
            End If
         k = k + 1
        
    

    End If

Next i

'Challenge Table
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

   

ws.Range("Q2").FormulaR1C1 = "=MAX(C[-6])"
ws.Range("Q3").FormulaR1C1 = "=MIN(C[-6])"
ws.Range("Q4").FormulaR1C1 = "=MAX(C[-5])"
ws.Range("P4").FormulaR1C1 = "=INDEX(C[-7],MATCH(RC[1],C[-4],0))"
ws.Range("P3").FormulaR1C1 = "=INDEX(C[-7],MATCH(RC[1],C[-5],0))"
ws.Range("P2").FormulaR1C1 = "=INDEX(C[-7],MATCH(RC[1],C[-5],0))"
ws.Range("Q2:Q3").NumberFormat = "00.00%"
ws.Columns("I:L").AutoFit
ws.Columns("O:Q").AutoFit

Next ws
End Sub







