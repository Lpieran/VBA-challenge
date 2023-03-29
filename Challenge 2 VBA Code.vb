Sub CalculateYearlyChanges()
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    Dim outputRow As Long

    'Loop through all worksheets
    For Each ws In Worksheets
        'Find last row of data in worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Clear any existing output in columns J to M
        ws.Range("J:M").ClearContents
        
        'Set headers for output
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Initialize the output row
        outputRow = 2
        
        'Loop through all rows of data
        For i = 2 To lastRow
            'Moved to a new ticker symbol
            If ticker <> ws.Cells(i, 1).Value Then
                'Output the previous ticker's results (if applicable)
                If i > 2 Then
                    ws.Cells(outputRow, 9).Value = ticker
                    ws.Cells(outputRow, 10).Value = yearlyChange
                    ws.Cells(outputRow, 11).Value = percentChange
                    ws.Cells(outputRow, 12).Value = totalVolume
                    outputRow = outputRow + 1
                End If
                
                'Set new ticker symbol and reset variables
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                totalVolume = ws.Cells(i, 7).Value
                
            Else
                'Update variables for the same ticker symbol
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
            
            'Check if we've reached the end of the data
            If i = lastRow Then
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = yearlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume
            End If
        Next i
    Next ws
End Sub


Sub FindMinMaxValuesForAllSheets()

    Dim ws As Worksheet
    Dim maxValK As Double
    Dim maxTickerK As String
    Dim minValK As Double
    Dim minTickerK As String
    Dim currValK As Double
    Dim lastRowK As Long
    Dim i As Long
    Dim maxValL As Double
    Dim maxTickerL As String
    Dim currValL As Double
    Dim lastRowL As Long

    For Each ws In ThisWorkbook.Worksheets
        lastRowK = ws.Cells(Rows.Count, "K").End(xlUp).Row
        maxValK = ws.Cells(2, "K").Value
        maxTickerK = ws.Cells(2, "I").Value
        minValK = ws.Cells(2, "K").Value
        minTickerK = ws.Cells(2, "I").Value
    For i = 3 To lastRowK
        currValK = ws.Cells(i, "K").Value
        If currValK > maxValK Then
            maxValK = currValK
            maxTickerK = ws.Cells(i, "I").Value
        ElseIf currValK < minValK Then
            minValK = currValK
            minTickerK = ws.Cells(i, "I").Value
        End If
    Next i
    
    lastRowL = ws.Cells(Rows.Count, "L").End(xlUp).Row
    maxValL = ws.Cells(2, "L").Value
    maxTickerL = ws.Cells(2, "I").Value
    For i = 3 To lastRowL
        currValL = ws.Cells(i, "L").Value
        If currValL > maxValL Then
            maxValL = currValL
            maxTickerL = ws.Cells(i, "I").Value
        End If
    Next i
    
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P2").Value = maxTickerK
        ws.Range("Q2").Value = Format(maxValK, "0.00%")
        ws.Range("P3").Value = minTickerK
        ws.Range("Q3").Value = Format(minValK, "0.00%")
        ws.Range("P4").Value = maxTickerL
        ws.Range("Q4").Value = maxValL
    Next ws

End Sub

Sub MergedCode()
   Call CalculateYearlyChanges
   Call FindMinMaxValuesForAllSheets
End Sub