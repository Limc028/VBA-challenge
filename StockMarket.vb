Sub StockAnalysisAcrossWorksheets()
    ' Define variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Set initial values
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        j = 2
        
        ' Loop through all stocks in the current worksheet
        For i = 2 To lastRow
            ' Check if current ticker symbol is different from previous ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Set opening price for new ticker symbol
                openingPrice = ws.Cells(i, 3).Value
            End If
            
            ' Add to total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if current ticker symbol is different from next ticker symbol or if it's the last row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow Then
                ' Set closing price for current ticker symbol
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate yearly and percent change
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                
                ' Output results
                ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(j, 10).Value = yearlyChange
                ws.Cells(j, 11).Value = percentChange
                ws.Cells(j, 12).Value = totalVolume
                ' Reset variables for next ticker symbol
                ticker = ws.Cells(i, 1).Value
                yearlyChange = 0
                percentChange = 0
                totalVolume = 0
                j = j + 1
            End If
        Next i
    Next ws
End Sub

