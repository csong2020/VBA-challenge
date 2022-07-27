Attribute VB_Name = "Module1"
Sub stock()

Dim ticker As String
Dim greatTicker As String
Dim leastTicker As String
Dim volumeTicker As String
Dim openPrice As Double
Dim closePrice As Double
Dim volume As Double
Dim percentChange As Double
Dim yearChange As Double
Dim tickerCounter As Integer
Dim greatInc As Double
Dim greatDec As Double
Dim greatVol As Double


'ticker needs to match worksheet type (string vs int)

For Each ws In Worksheets
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    tickerCounter = 2                               'initial ticket row

    ws.Cells(1, 9).Value = "Ticker"                 'creates stock calc columns
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(1, 18) = "Ticker"
    ws.Cells(1, 19) = "Value"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    greatInc = 0
    greatDec = 0
    greatVol = 0
    
    openPrice = ws.Cells(2, 3).Value                   'stores initial openPrice in that worksheet
    For i = 2 To LastRow                      'loops through ticker column as long as ticker matches
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then              'if ticker nextline doesn't match, then...
            volume = ws.Cells(i, 7).Value + volume                         'add last volume
            ws.Cells(tickerCounter, 9).Value = ticker                      'writes ticker name
            closePrice = ws.Cells(i, 6).Value                              'stores close price
            ws.Cells(tickerCounter, 10).Value = closePrice - openPrice     'writes yearly change
            
            If ws.Cells(tickerCounter, 10).Value > 0 Then
                ws.Cells(tickerCounter, 10).Interior.ColorIndex = 4        'if positive, green
            Else: ws.Cells(tickerCounter, 10).Interior.ColorIndex = 3      'if negative, red
            End If
                                                                        'vvv writes percent change
            ws.Cells(tickerCounter, 11).Value = (closePrice - openPrice) / openPrice
            ws.Cells(tickerCounter, 12).Value = volume                     'writes total stock volume
            tickerCounter = tickerCounter + 1                           'increments tickerCounter
            closePrice = ws.Cells(i + 1, 6).Value                          'stores new close price
            openPrice = ws.Cells(i + 1, 3).Value                           'stores new open price
            volume = 0                                                  'resets volume
        Else
        'vvv what happens if names match
        volume = volume + ws.Cells(i, 7).Value                             'sums volume
        ticker = ws.Cells(i, 1).Value                                      'stores current ticker name
        End If
    Next i
        
    

    
    For i = 2 To LastRow
        If ws.Cells(i, 11) > 0 And ws.Cells(i, 11) > greatInc Then
        greatInc = ws.Cells(i, 11).Value                                         'sets greatInc value as current value
        ws.Cells(2, 19).Value = greatInc                                         'writes greatest increase value
        greatTicker = ws.Cells(i, 9).Value
        ws.Cells(2, 18).Value = greatTicker                                      'writes greatest increase ticker
        End If
    Next i
        
    For i = 2 To LastRow
        If ws.Cells(i, 11) < 0 And ws.Cells(i, 11) < greatDec Then
        greatDec = ws.Cells(i, 11).Value                                        'sets greatDec value as current value
        ws.Cells(3, 19).Value = greatDec
        leastTicker = ws.Cells(i, 9).Value
        ws.Cells(3, 18).Value = leastTicker                                      'writes greatest decrease ticker
        End If
    Next i
        
    For i = 2 To LastRow
        If ws.Cells(i, 12) > 0 And ws.Cells(i, 12) > greatVol Then
        greatVol = ws.Cells(i, 12).Value                                        'sets greatVol value as current value
        ws.Cells(4, 19).Value = greatVol
        greatVolume = ws.Cells(i, 9).Value
        ws.Cells(4, 18).Value = greatVolume                                      'writes greatest volume ticker
        End If
    Next i
    
    For i = 2 To LastRow
        ws.Cells(i, 11).Style = "Percent"   'puts percent style on percent change column
    Next i
    
    ws.Cells(2, 19).Style = "Percent"   'puts percent style on greatest increase
    ws.Cells(3, 19).Style = "Percent"   'puts percent style on greatest decrease
    
greatInc = 0        'reset greatest increase
greatDec = 0        'reset greatest decrease
greatVol = 0        'reset greatest volume
    
Next ws

End Sub


