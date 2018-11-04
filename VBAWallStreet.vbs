'Create a VBA script that will loop through all stocks and create a column for:
'1)    Total Volume of each Stock’s Sales throughout the year,
'2)    Stock’ Ticker Symbol,
'3)    Yearly Price Change (Opening price to Closing Price)
'4)    Yearly Percent Change (Opening price to Closing Price)
    'a.    Conditional Format – Positive (Green) & Negative (Red)

    '1)    Create Table (4x3 as below) W/ Ticker Symbol & Value of:
        'a.    ‘Greatest % Increase’,
        'b.    ‘Greatest % Decrease’,
        'c.    ‘Greatest Total Volume’.
'-------------------------------------------

Sub VBAWallStreet():

Dim Ticker As String
Dim Total_Volume_of_Stock, Yearly_Change, Percent_Change, Opening_Price, Closing_Price As Double
Dim CurrentRow As Integer

    For Each ws In Worksheets
    
        CurrentRow = 2                                                                  'Begin on 2nd row of each sheet
        Total_Volume_of_Stock = 0                                                       'Will maintain sum of unique ticker
        Opening_Price = ws.Cells(2, 3)
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row                                 'determine last row of each sheet
    
            For i = 2 To LastRow                                                            'For each row in the WS
    '
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then                                'IS NEXT TICKER UNIQUE? if so, then TIME TO SUMMARIZE
                Ticker = ws.Cells(i, 1)                                                        'COLLECT CURRENT TICKER
                Total_Volume_of_Stock = Total_Volume_of_Stock + ws.Cells(i, 7)                 'ADD FINAL DAY TO VOLUME SUM
                Closing_Price = ws.Cells(i, 6)                                                 'COLLECT FINAL CLOSING PRICE
                Yearly_Change = Closing_Price - Opening_Price                                  'FINAL CLOSING PRICE - INITIAL OPENING PRICE
    
                   If Opening_Price = 0 Then
                       Percent_Change = 0                                                      'EXCEPTION: CANNOT DIVIDE BY 0
                   Else
                       Percent_Change = 100 * Yearly_Change / Opening_Price                    'CALCULATE PERCENT CHANGE
                   End If
                                                                                               'SUMMARY:
                 ws.Range("I" & CurrentRow).Value = Ticker                                      'TICKER SYMBOL
                 ws.Range("J" & CurrentRow).Value = Yearly_Change                               'CHANGE
                 ws.Range("K" & CurrentRow).Value = (Percent_Change & "%")                      '% CHANGE
                 ws.Range("L" & CurrentRow).Value = Format(Total_Volume_of_Stock, "$#,###")     'VOLUME
    
                Total_Volume_of_Stock = 0                                                        'RESET VOLUME
                Opening_Price = ws.Cells(i + 1, 3)                                               'SET NEW OPENING PRICE FOR NEXT TICKER;
     
                              If ws.Range("J" & CurrentRow).Value >= 0 Then             'positive yearly change --> Green
                                   ws.Range("J" & CurrentRow).Interior.ColorIndex = 4
                              ElseIf ws.Range("J" & CurrentRow).Value < 0 Then
                                   ws.Range("J" & CurrentRow).Interior.ColorIndex = 3    'negative yearly change --> Red
                              End If

                    CurrentRow = CurrentRow + 1
            Else                                                                         'NEXT TICKER IS NOT NEW, SO CONTINUE TO ADD TO CURRENT SUM
                Total_Volume_of_Stock = Total_Volume_of_Stock + ws.Cells(i, 7)
            End If
        Next i

    'Column Name Reset + New Column Names;
    
        ws.Range("A1") = "Ticker Symbol"
        ws.Range("B1") = "Date"
        ws.Range("C1") = "Opening Price"
        ws.Range("D1") = "High"
        ws.Range("E1") = "Low"
        ws.Range("F1") = "Closing Price"
        ws.Range("G1") = "Daily Volume"
        ws.Range("H1") = "-------"
        ws.Range("I1") = "Ticker Symbol"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
    
    Next ws


End Sub


'LOGIC:
'1) go to sheet 1
'2) create new column of unique ticker symbols
'3) create new column of total volume of each symbol
'4) create new column of price change
