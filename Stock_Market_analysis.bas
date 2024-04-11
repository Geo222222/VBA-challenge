Attribute VB_Name = "Module1"
Sub CalculateYearlyChanges()

    ' The declaring of all variables being used.
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerColumn As Long
    Dim dateColumn As Long
    Dim openColumn As Long
    Dim closeColumn As Long
    Dim volColumn As Long
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim outputRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
    
    ' Initializing the greatest for increase, decrease and volume values to be used in comparison
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0

    ' This is a for loop assigned to each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
    
        ' Finds the last row of data in the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Set the column numbers for each data type
        tickerColumn = 1    ' 1 For column A
        dateColumn = 2      ' 2 For column B
        openColumn = 3      ' 3 For column C
        closeColumn = 6     ' 6 For column F
        volColumn = 7       ' 7 For column G
        
        ' Setting the headers for the output columns (ticker, yearly change, percentage change, and total stock volume
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Declaring the output row as 2 or 2nd row
        outputRow = 2
        
        ' Another for loop defining what happens for each row in the spreadsheet starting with 2
        For i = 2 To lastRow
        
            ' First we check if the ticker symbol changes with the worksheet.cells built in
            ' If ticker changes, calculate and output the yearly change, percentage change, and total stock volume in output row and chosen columns, 9, 10, 11, and 12
                
            If ws.Cells(i, tickerColumn).Value <> ws.Cells(i - 1, tickerColumn).Value Then
                
                ' Output for each ticker
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = yearlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume
                
                ' Moves to the next output row
                outputRow = outputRow + 1
                
                ' Updates greatest increase, decrease, and volume and asigns to that ticker/symbol
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    tickerGreatestIncrease = ticker
                End If
                
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    tickerGreatestDecrease = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    tickerGreatestVolume = ticker
                End If
                
                ' Reset variables for the new ticker
                ticker = ws.Cells(i, tickerColumn).Value
                yearlyChange = 0
                percentChange = 0
                totalVolume = 0
                openingPrice = ws.Cells(i, openColumn).Value ' Initialize opening price
            End If
            
            ' Next If statement for the first row: If it's the first row, initialize the opening price
            If i = 2 Then
                openingPrice = ws.Cells(i, openColumn).Value
            End If
            
            ' Calculate the yearly change and total stock volume
            closingPrice = ws.Cells(i, closeColumn).Value
            yearlyChange = closingPrice - openingPrice
            totalVolume = totalVolume + ws.Cells(i, volColumn).Value
            
            ' Calculate the percentage change <> is not equal to
            If openingPrice <> 0 Then
                percentChange = (yearlyChange / openingPrice) * 100
            Else
                percentChange = 0
            End If
            
        Next i
        
        ' Output row and column for the ticker's information
        ws.Cells(outputRow, 9).Value = ticker
        ws.Cells(outputRow, 10).Value = yearlyChange
        ws.Cells(outputRow, 11).Value = percentChange
        ws.Cells(outputRow, 12).Value = totalVolume
        
        ' Output row and column for the greatest increase, decrease, and total volume along with their respective tickers
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 16).Value = tickerGreatestIncrease
        ws.Cells(3, 16).Value = tickerGreatestDecrease
        ws.Cells(4, 16).Value = tickerGreatestVolume
        
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 17).Value = greatestVolume
        
        ' Apply conditional formatting to column J (Yearly Change)
        With ws.Range(ws.Cells(3, 10), ws.Cells(lastRow, 10))
            .FormatConditions.Delete ' Delete existing formatting
            ' Add new formatting for positive and negative numbers
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red
        End With
        
    Next ws

End Sub

