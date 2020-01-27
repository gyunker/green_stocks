# green_stocks
Green Stocks Homework

This homework we edited a VBA code set and refactored it to leverage arrays and indexes.  The final VBA code is below.  Data is within the attached xlsm document.

======================================

Sub tickerIndex()
Worksheets("2018").Activate
stockIndex = 0
Dim tickers() As String
Dim startingPrice() As Single
Dim totalVolume() As Double
Dim endingPrice() As Single

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

For j = 2 To RowCount
        If Cells(j - 1, 1).Value <> Cells(j, 1).Value And Cells(j, 1).Value = Cells(j + 1, 1).Value Then
                
                stockIndex = stockIndex + 1
                
                ReDim Preserve tickers(stockIndex + 1), startingPrice(stockIndex + 1), totalVolume(stockIndex), endingPrice(stockIndex)
                
                tickers(stockIndex) = Cells(j, 1).Value
                
                startingPrice(stockIndex) = Cells(j, 6).Value
                
                totalVolume(stockIndex) = totalVolume(stockIndex) + Cells(j, 8).Value
            
            End If
            
            If Cells(j - 1, 1).Value = Cells(j, 1).Value And Cells(j, 1).Value = Cells(j + 1, 1).Value Then
                
                totalVolume(stockIndex) = totalVolume(stockIndex) + Cells(j, 8).Value
            
            End If
            
            If Cells(j - 1, 1).Value = Cells(j, 1).Value And Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
                
                endingPrice(stockIndex) = Cells(j, 6).Value
                
                totalVolume(stockIndex) = totalVolume(stockIndex) + Cells(j, 8).Value
            
            End If
        Next j
     

    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("My Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Start Price"
    Cells(3, 4).Value = "End Price"
    Cells(3, 5).Value = "Return"
    
    'Populate the table with the array values via a loop
    For i = 1 To stockIndex
        Cells(3 + i, 1) = tickers(i)
        Cells(3 + i, 2) = totalVolume(i)
        Cells(3 + i, 3) = startingPrice(i)
        Cells(3 + i, 4) = endingPrice(i)
        Cells(3 + i, 5) = endingPrice(i) / startingPrice(i) - 1
    Next i

    'Format the data grid
    Range("A3:E3").Font.FontStyle = "Bold"
    Range("A3:E3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Columns("B").NumberFormat = "#,##0"
    Columns("C").NumberFormat = "$#,##0.00"
    Columns("D").NumberFormat = "$#,##0.00"
    Columns("E").NumberFormat = "0.0%"
    Columns("B").AutoFit
    Columns("C").AutoFit
    Columns("D").AutoFit
    Columns("E").AutoFit
    dataRowStart = 4
    dataRowEnd = stockIndex + 3
    For i = dataRowStart To dataRowEnd

        If Cells(i, 5) > 0 Then

            Cells(i, 5).Interior.Color = vbGreen

        Else

            Cells(i, 5).Interior.Color = vbRed

        End If

    Next i

End Sub


