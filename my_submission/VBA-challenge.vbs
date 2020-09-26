Attribute VB_Name = "Module1"
Sub getStocks()

    '-------------------------------------------------------------------------------------------------------
    'Script to run through all stock sheets and display information based on the guidelines in the README.md
    '-------------------------------------------------------------------------------------------------------
    
    'Row variables
    Dim ticker As String
    
    'yearlyChange = (last close value - first open value)
    Dim oldFirstOpenVal As Double
    Dim firstOpenVal As Double
    Dim lastCloseVal As Double
    
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim percentString As String
    Dim totalStockVolume As Double
    totalStockVolume = 0
    
    'Row tracker
    Dim tickerRow As Integer
    'Always default values to two because headers are on line one.
    tickerRow = 2
    
    Dim lRow As Long
    Dim prevTicker As String
    
    
    Dim i As Long
    
    Dim ws As Worksheet
    
    'Challenge vars
    Dim greatestPerInc As Double
    Dim GPITicker As String
    Dim greatestPerDec As Double
    Dim GPDTicker As String
    Dim greatestTotalVolume As Double
    Dim GTVTicker As String
    Dim challengePerString As String
    
    For Each ws In Sheets
        'Reset tickerRow and totalStockVolume for each sheet
        tickerRow = 2
        totalStockVolume = 0
        'Reset challenge vars
        greatestPerInc = 0
        greatestPerDec = 0
        greatestTotalVolume = 0
        'Titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            If i = 2 Then
                firstOpenVal = ws.Cells(2, 3).Value
            End If
            'Check if the next value is different, if it is this is the last same type ticker...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'This row has lastClose Val and this row has new firstOpenVal
                ticker = ws.Cells(i, 1).Value
                'Update oldFirstOpenVal to what had been stored, and store the next value into firstOpenVal
                lastCloseVal = ws.Cells(i, 6).Value
                oldFirstOpenVal = firstOpenVal
                'Set new firstOpenVal to the next value
                firstOpenVal = ws.Cells(i + 1, 3).Value
                yearlyChange = lastCloseVal - oldFirstOpenVal
                If oldFirstOpenVal = 0 Then
                    'Set percentChange to 0 to avoiding dividing by 0
                    percentChange = 0
                    percentString = Format(percentChange, "Percent")
                Else
                    'Means theres no worries of dividing by 0
                    percentChange = ((lastCloseVal - oldFirstOpenVal) / oldFirstOpenVal)
                    percentString = Format(percentChange, "Percent")
                End If
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
                
                'Update challenge vars
                If greatestPerInc < percentChange Then
                    greatestPerInc = percentChange
                    GPITicker = ticker
                End If
                
                If greatestPerDec > percentChange Then
                    greatestPerDec = percentChange
                    GPDTicker = ticker
                End If
                
                If greatestTotalVolume < totalStockVolume Then
                    greatestTotalVolume = totalStockVolume
                    GTVTicker = ticker
                End If
                
                'Print the values to there respective columns
                ws.Range("I" & tickerRow).Value = ticker

                ws.Range("J" & tickerRow).Value = yearlyChange
            
                ws.Range("K" & tickerRow).Value = percentString
            
                ws.Range("L" & tickerRow).Value = totalStockVolume
            
                ' Add one to the summary table row
                tickerRow = tickerRow + 1
            
                ' Reset the totalStockVolume'
                totalStockVolume = 0
            Else
                ticker = ws.Cells(i, 1).Value
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        'Print challenge Tickers and Values
        ws.Range("P2").Value = GPITicker
        ws.Range("P3").Value = GPDTicker
        ws.Range("P4").Value = GTVTicker
        challengePerString = Format(greatestPerInc, "Percent")
        ws.Range("Q2").Value = challengePerString
        challengePerString = Format(greatestPerDec, "Percent")
        ws.Range("Q3").Value = challengePerString
        ws.Range("Q4").Value = greatestTotalVolume
        
        'Condition formatting yearlyChange
        Dim rng As Range
        Dim condition1 As FormatCondition, condition2 As FormatCondition
        Set rng = ws.Range("J2", "J" & tickerRow)
        rng.FormatConditions.Delete
        Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
        Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
    
        With condition1
        .Interior.Color = vbGreen
        End With

        With condition2
        .Interior.Color = vbRed
        End With
        
        'End of page
    Next ws
End Sub
