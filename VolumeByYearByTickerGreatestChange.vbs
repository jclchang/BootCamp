Sub tickerByYearByTickerPctChange()

Dim Current As Worksheet
Dim LastRow As Long

Dim i As Long
Dim tickerVolume As Double
Dim tickerCount As Integer
Dim tickerOpen As Double
Dim tickerClose As Double
Dim pctChange As Double

Dim greatestIncTicker As String
Dim greatestIncPct As Double
Dim greatestDecTicker As String
Dim greatestDecPct As Double
Dim greatestTotlVolTicker As String
Dim greatestTotVolume As Double


' Loop through all of the worksheets in the active workbook.
For Each Current In Worksheets
    'MsgBox Current.Name
    
    LastRow = Current.Cells(Rows.Count, 1).End(xlUp).Row
    
	Current.Cells(1, 9).Value = "Ticker"
    Current.Cells(1, 10).Value = "Yearly Change"
    Current.Columns("j").NumberFormat = "0.00000000"
    Current.Cells(1, 11).Value = "Percent Change"
    Current.Columns("k").NumberFormat = "0.00%"
    Current.Cells(1, 12).Value = "Total Stock Volume"

    Current.Cells(1, 16).Value = "Ticker"
    Current.Cells(1, 17).Value = "Value"
    Current.Cells(2, 15).Value = "Greatest % Increase"
    Current.Cells(3, 15).Value = "Greatest % Decrease"
    Current.Cells(4, 15).Value = "Greatest Total Volume"
   
    tickerVolume = 0
    tickerCount = 1
    tickerOpen = 0
    tickerClose = 0

    greatestIncTicker = ""
    greatestIncPct = 0
    greatestDecTicker = ""
    greatestDecPct = 0
    greatestTotlVolTicker = ""
    greatestTotVolume = 0

    For i = 2 To LastRow

        tickerVolume = tickerVolume + Current.Cells(i, 7).Value
        
        If tickerOpen = 0 Then
            tickerOpen = Current.Cells(i, 3).Value
        End If
               
        If Current.Cells(i, 1).Value <> Current.Cells(i + 1, 1).Value Then
            Current.Cells(tickerCount + 1, 9).Value = Current.Cells(i, 1).Value
            Current.Cells(tickerCount + 1, 12).Value = tickerVolume
            
            tickerClose = Current.Cells(i, 6).Value
            Current.Cells(tickerCount + 1, 10).Value = tickerClose - tickerOpen
            
            If tickerOpen <> 0 Then
                Current.Cells(tickerCount + 1, 11).Value = (tickerClose - tickerOpen) / tickerOpen
            Else
                Current.Cells(tickerCount + 1, 11).Value = 0
            End If
            
            pctChange = Current.Cells(tickerCount + 1, 11).Value
            
            If pctChange < 0 Then
                Current.Cells(tickerCount + 1, 10).Interior.ColorIndex = 3
                If pctChange < greatestDecPct Then
                    greatestDecTicker = Current.Cells(i, 1).Value
                    greatestDecPct = pctChange
                End If
            Else
                Current.Cells(tickerCount + 1, 10).Interior.ColorIndex = 4
                If pctChange > greatestIncPct Then
                    greatestIncTicker = Current.Cells(i, 1).Value
                    greatestIncPct = pctChange
                End If
            End If
            
            If tickerVolume > greatestTotVolume Then
                greatestTotlVolTicker = Current.Cells(i, 1).Value
                greatestTotVolume = tickerVolume
            
            End If
                          

            tickerVolume = 0
            tickerCount = tickerCount + 1
            tickerOpen = Current.Cells(i + 1, 3).Value

        End If
    Next i
    
    Current.Cells(2, 16).Value = greatestIncTicker
    Current.Cells(2, 17).Value = greatestIncPct
    Current.Cells(2, 17).NumberFormat = "0.00%"

    Current.Cells(3, 16).Value = greatestDecTicker
    Current.Cells(3, 17).Value = greatestDecPct
    Current.Cells(3, 17).NumberFormat = "0.00%"

    Current.Cells(4, 16).Value = greatestTotlVolTicker
    Current.Cells(4, 17).Value = greatestTotVolume
    
    Next Current

End Sub
