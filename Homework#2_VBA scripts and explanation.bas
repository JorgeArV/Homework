Attribute VB_Name = "Module1"
Sub HomeworkFinal():

'We apply our code to all worksheets
For Each ws In Worksheets

Dim TickerCounter As Integer
Dim LastRow As Long

'We will use 3 counters. Each of them begin at 0:
TickerCounter = 0 'TickerCounter goes up by 1 every time we add a new ticker to our ticker column.
Dayscounter = 0 'Dayscounter keeps track of the number of rows between the first datapoint and the last datapoint per ticker.
SumVolume = 0 'Sumvolume adds the number of stocks per ticker

    'We determine the number of rows of the dataset:
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'We write down the labels for the chart we need to create:
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"

    For i = 2 To LastRow
    
        'If ticker in cell (i,1) is different to ticker in the cell below it, then:
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'We add the name of the ticker to our Tickers column:
            ws.Cells(2 + TickerCounter, 10).Value = ws.Cells(i, 1).Value
            
            'We substract the price under <close> on the last day available minus the price under <open> on the first day available.
            'We add the result of this substraction next to the corresponding ticker.
            ws.Cells(2 + TickerCounter, 11).Value = ws.Cells(i, 6).Value - ws.Cells(i - Dayscounter, 3).Value
            
            'We add conditional formating:
                If ws.Cells(2 + TickerCounter, 11).Value < 0 Then
                    ws.Cells(2 + TickerCounter, 11).Interior.ColorIndex = 3
                ElseIf ws.Cells(2 + TickerCounter, 11).Value > 0 Then
                    ws.Cells(2 + TickerCounter, 11).Interior.ColorIndex = 4
                End If
            
            'We determine growth by dividing the values under the 'yearly change' column by the initial price i.e. price under <open> on the first day available per ticker
            'If initial price is 0, finding the percent change is a nonsensical exercise. Hence, if initial price is 0, a "not applicable" message will appear.
            
                If ws.Cells(i - Dayscounter, 3) = 0 Then
                    ws.Cells(2 + TickerCounter, 12).Value = "Not Applicable"
                Else
                    ws.Cells(2 + TickerCounter, 12).Value = FormatPercent((ws.Cells(2 + TickerCounter, 11).Value / ws.Cells(i - Dayscounter, 3)), 2)
                End If
                     
           'We add the final volume datapoint per ticker to our SumVolume tracker. Which has been keeping track of all the stocks purchased.
           'We fill out the Total Stock Volume column with the number of stocks purchased per ticker.
            SumVolume = SumVolume + ws.Cells(i, 7)
            ws.Cells(2 + TickerCounter, 13).Value = SumVolume
            
           'TickerCounter goes up by one every time we find a new ticker
            TickerCounter = TickerCounter + 1
                       
           'Dayscounter and SumVolume counter get resetted to 0 once we add all the information of a ticker
            Dayscounter = 0
            SumVolume = 0
        
        'If ticker in cell (i,1) is the same as ticker in the cell below it, then:
        ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        
            'Dayscounter keeps track of the number of rows between the first and last datapoint per ticker. We need to know the row location of the first datapoint per ticker in order to determine open price.
            'Dayscounter is the way we find this row location.
            Dayscounter = Dayscounter + 1
            
            'SumVolume keeps the count of the number of stocks purchased per ticker.
            SumVolume = SumVolume + ws.Cells(i, 7)
                           
        End If
        
        
    Next i
Next ws
    
End Sub

Sub ChallengeFinal():

'We apply our code to all worksheets
For Each ws In Worksheets

Dim LastRow As Long

    'We determine the number of rows of the dataset:
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
   'We write down the labels for the chart we need to create:
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"

    For i = 2 To LastRow:
    
        'If a ticker does not have a value for the percent change, then it is excluded from our process of identifying greatest % increase
        If ws.Cells(i, 12) = "Not Applicable" Then
    
        'Each datapoint under the "Percent Change" column will be compared with the cell where the "greatest % increase" is showed.
        '"greatest % increase" will continue being updated until there is no higher value.
        ElseIf ws.Cells(i, 12) > ws.Cells(2, 18) Then
            
                ws.Cells(2, 18).Value = FormatPercent(ws.Cells(i, 12).Value, 2)
                ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
        End If
        
        'If a ticker does not have a value for the percent change, then it is excluded from our process of identifying greatest % decrease
        If ws.Cells(i, 12) = "Not Applicable" Then
        
        'Each datapoint under the "Percent Change" column will be compared with the cell where the "greatest % decrease" is showed.
        '"greatest % decrease" will continue being updated until there is no higher value.
        
        ElseIf ws.Cells(i, 12) < ws.Cells(3, 18) Then
            
                ws.Cells(3, 18).Value = FormatPercent(ws.Cells(i, 12).Value, 2)
                ws.Cells(3, 17).Value = ws.Cells(i, 10).Value
        End If
              
        'Each datapoint under the "Total Stock Volume" column will be compared with the cell where the "greatest total volume" is showed.
        '"greatest total volume" will continue being updated until there is no higher value.
        
        If ws.Cells(i, 13) > ws.Cells(4, 18) Then
            
                ws.Cells(4, 18).Value = ws.Cells(i, 13).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 10).Value
        End If
    
    
    Next i
Next ws
           
End Sub
