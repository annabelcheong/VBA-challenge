Sub stocks()

For Each ws In Worksheets

Dim row_counter As Integer
Dim vol As Long
Dim yearly_change As Double
Dim opening_price As Double
Dim closing_price As Double
Dim percent_change As Double
Dim a As Integer
'Challenge Exercise variables
Dim b As Integer
Dim g_percent_inc As Double
Dim g_percent_dec As Double
Dim ticker_inc As String
Dim ticker_dec As String
Dim ticker_vol_max As String
Dim vol_max As Double

row_counter = 1
a = 0


'Heading Columns of Solution
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"



LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow


    'if above row does not equal row
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
        'print ticker onto summary table
        ws.Cells(row_counter + 1, 9).Value = ws.Cells(i, 1).Value
        
        'open price is first row,3 and closing price is last row,6. a is the row range of company.
        closing_price = ws.Cells(i, 6).Value
        opening_price = ws.Cells(i - a, 3).Value
                
        'yearly range and print yearly change to summary table
        yearly_change = closing_price - opening_price
        ws.Cells(row_counter + 1, 10) = yearly_change
               
            '---------------------------
            'SHADE IN CELLS GREEN OR RED FOR + OR - YEARLY CHANGE
               
            'if yearly_change is more than 0, shade cell in green
            If ws.Cells(row_counter + 1, 10) > 0 Then
            ws.Cells(row_counter + 1, 10).Interior.ColorIndex = 4
                
            'if yearly_change is less than 0, shade cell in red
            ElseIf ws.Cells(row_counter + 1, 10) < 0 Then
            ws.Cells(row_counter + 1, 10).Interior.ColorIndex = 3
                
            End If
                
            '-------------------------
            'TROUBLESHOOTING FOR TICKER WITH OPENING PRICE AT 0.00
            If opening_price = 0 Then
            MsgBox ("The opening price is 0.00 for" + " " + ws.Cells(i, 1).Value)
            End If
                
                '--------------------------
                
            'Print percent change in column K/column 11
            If opening_price <> 0 Then
            ws.Cells(row_counter + 1, 11) = Format(yearly_change / opening_price, "0.00%")
            
            End If
                 
        'reset value a
        a = 0
            
        'print volume onto summary table
        ws.Cells(row_counter + 1, 12).Value = volume + ws.Cells(i, 7)
        
        'reset volume for next ticker type
        volume = 0
        
        'reset yearly_change for next ticker type
        yearly_change = 0
           
        'row counter for summary table to print on new row
        row_counter = row_counter + 1
        
    '  ---------------------
    'if above row equals row
    ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        'stored closing volume = Cells(i, 7)
        volume = volume + ws.Cells(i, 7).Value
        
        'increment the range for use in formula for when ticker does not match
        a = a + 1
        
    End If
    
Next i

'Challenge Exercise

b = 3
'initialise greatest/lowest percent to first row value of percentage change (col K), initialise volume
    g_percent_inc = ws.Cells(b - 1, 11).Value
    ws.Cells(2, 16).Value = Format(g_percent_inc, "0.00%")
    
    g_percent_dec = ws.Cells(b - 1, 11).Value
    ws.Cells(b, 16).Value = Format(g_percent_dec, "0.00%")
    
    vol_max = ws.Cells(b - 1, 12).Value

'initialise ticker name for greatest/lowest percent,vol max into challenge summary table
    ticker_inc = ws.Cells(b - 1, 9).Value
    ticker_dec = ws.Cells(b - 1, 9).Value
    ticker_vol_max = ws.Cells(b - 1, 9).Value

LastRow1 = Cells(Rows.Count, 9).End(xlUp).Row
For b = 3 To LastRow1


    
        'loop to find the greatest % value
        If ws.Cells(b, 11).Value > g_percent_inc Then
            g_percent_inc = ws.Cells(b, 11).Value
            ticker_inc = ws.Cells(b, 9)
            'print greatest percentage change into challenge summary table
            ws.Cells(2, 16).Value = Format(g_percent_inc, "0.00%")
            ws.Cells(2, 15).Value = ticker_inc
        End If
        
        'loop to find the lowest % value
        If ws.Cells(b, 11).Value < g_percent_dec Then
            g_percent_dec = ws.Cells(b, 11).Value
            ticker_dec = ws.Cells(b, 9).Value
            'print lowest percentage change into challenge summary table
            ws.Cells(3, 16).Value = Format(g_percent_dec, "0.00%")
            ws.Cells(3, 15).Value = ticker_dec
        End If
    
        'loop to find the greatest volume
        If ws.Cells(b, 12).Value > vol_max Then
            vol_max = ws.Cells(b, 12).Value
            ticker_vol_max = ws.Cells(b, 9).Value
            'Print greatest volume into challenge summary table
            ws.Cells(4, 16).Value = vol_max
            ws.Cells(4, 15).Value = ticker_vol_max
        End If
        
        
Next b

Next ws

End Sub
