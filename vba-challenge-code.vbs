Sub Stock_Market()


'Set initial variables
Dim ws As Worksheet
Dim opening_price As Double
Dim closing_price As Double
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock As Double
Dim Summary_Table_Row As Integer


'Loop through each worksheet
For Each ws In Worksheets

'Set total stock volume to 0
total_stock = 0

'Set summary table integer
Summary_Table_Row = 2

'Set last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Set the column headers for summary table
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


'Loop through all tickers
For i = 2 To lastrow
    
    'Check if current ticker is not the same to the previous ticker
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    opening_price = ws.Cells(i, 3)

    'Check if current ticker is not the same to the next ticker
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    closing_price = ws.Cells(i, 6).Value
    
        'Set ticker name
        ticker = ws.Cells(i, 1).Value
        
        'Calculate yearly change
        yearly_change = closing_price - opening_price
        
        'Calculate percentage yearly change
        percent_change = (yearly_change / opening_price)
        
        'Calculate total stock value
        total_stock = total_stock + ws.Cells(i, 7).Value
        
        'Print summary table variable
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("J" & Summary_Table_Row).Value = yearly_change
            ws.Columns("J:J").NumberFormat = "0.00"
        ws.Range("K" & Summary_Table_Row).Value = percent_change
            ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Range("L" & Summary_Table_Row).Value = total_stock
        
        'Add one to the summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset total stock volume
        total_stock = 0
        yearly_change = 0
        
    'If the same
    Else
        
        'Add to the total stock value
        total_stock = total_stock + ws.Cells(i, 7).Value
        
    End If
    
Next i

'---------------------------------------------

'Set initial variables
Dim greatest_increase As Double
Dim increase_ticker As String
Dim greatest_decrease As Double
Dim decrease_ticker As String
Dim greatest_total As Double
Dim total_ticker As String

'Set final last row
final_last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row


'Set column headers
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


'Loop through greatest values
For j = 2 To final_last_row

    'Check for greatest % increase maximum value and ticker
    If greatest_increase < ws.Cells(j, 11).Value Then
    
        greatest_increase = ws.Cells(j, 11).Value
        increase_ticker = ws.Cells(j, 9).Value
    
    End If
    
    'Check for greatest % decrease maximum value and ticker
    If greatest_decrease > ws.Cells(j, 11).Value Then
        
        greatest_decrease = ws.Cells(j, 11).Value
        decrease_ticker = ws.Cells(j, 9).Value
    
    End If
    
    'Check for greatest total volume value and ticker
    If greatest_total < ws.Cells(j, 12).Value Then
        
        greatest_total = ws.Cells(j, 12).Value
        total_ticker = ws.Cells(j, 9).Value
    
    End If

Next j

'---------------------------------------------

'Print values into the table
ws.Range("P2").Value = increase_ticker
ws.Range("Q2").Value = greatest_increase
    ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("P3").Value = decrease_ticker
ws.Range("Q3").Value = greatest_decrease
    ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("P4").Value = total_ticker
ws.Range("Q4").Value = greatest_total


'Column formatting
For k = 2 To final_last_row
    
    'Yearly Change conditional formatting
    If ws.Cells(k, 10) > 0 Then
        ws.Cells(k, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(k, 10).Interior.ColorIndex = 3
    End If
Next k


ws.Range("I1:Q1").Font.Bold = True
ws.Range("O2:O4").Font.Bold = True
ws.Columns("I:Q").AutoFit


Next ws
End Sub
