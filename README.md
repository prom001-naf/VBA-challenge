# VBA-challenge
Sub CalculateQuarterlyStockChangeAllSheets()
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Call the function that calculates and formats the data for each worksheet
        CalculateQuarterlyStockChange ws
    Next ws
End Sub

Sub CalculateQuarterlyStockChange(ws As Worksheet)
    Dim lastRow As Long, i As Long
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double
    Dim startDate As Date, endDate As Date
    Dim quarterlyChange As Double, percentChange As Double
    Dim totalVolume As Double
    Dim currentQuarter As Integer, newQuarter As Integer
    
    ' Variables to track the greatest increases, decreases, and volumes
    Dim maxPercentIncrease As Double, maxPercentIncreaseTicker As String
    Dim maxPercentDecrease As Double, maxPercentDecreaseTicker As String
    Dim maxTotalVolume As Double, maxTotalVolumeTicker As String
    
    ' Initialize tracking variables
    maxPercentIncrease = -100000
    maxPercentDecrease = 100000
    maxTotalVolume = 0
    
    ' Find the last row of data in the sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Output headers in the result columns
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Quarterly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    
    Dim outputRow As Long
    outputRow = 2
    
    For i = 2 To lastRow
        ' Get ticker symbol
        ticker = ws.Cells(i, 1).Value
        
        ' Get date and quarter for the current row
        startDate = ws.Cells(i, 2).Value
        currentQuarter = DatePart("q", startDate)
        
        ' Initialize variables for the current quarter
        openPrice = ws.Cells(i, 3).Value
        totalVolume = 0
        
        ' Loop through the quarter
        Do While ticker = ws.Cells(i, 1).Value And currentQuarter = DatePart("q", ws.Cells(i, 2).Value)
            ' Keep updating the close price and sum up the volume
            closePrice = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            i = i + 1
            If i > lastRow Then Exit Do
        Loop
        
        ' Calculate the quarterly change and percent change
        quarterlyChange = closePrice - openPrice
        If openPrice <> 0 Then
            percentChange = (quarterlyChange / openPrice) * 100
        Else
            percentChange = 0
        End If
        
        ' Output the results
        ws.Cells(outputRow, 10).Value = ticker
        ws.Cells(outputRow, 11).Value = Format(quarterlyChange, "0.00")
        ws.Cells(outputRow, 12).Value = Format(percentChange, "0.00") & "%"
        ws.Cells(outputRow, 13).Value = totalVolume
        
        ' Check for greatest percentage increase
        If percentChange > maxPercentIncrease Then
            maxPercentIncrease = percentChange
            maxPercentIncreaseTicker = ticker
        End If
        
        ' Check for greatest percentage decrease
        If percentChange < maxPercentDecrease Then
            maxPercentDecrease = percentChange
            maxPercentDecreaseTicker = ticker
        End If
        
        ' Check for greatest total volume
        If totalVolume > maxTotalVolume Then
            maxTotalVolume = totalVolume
            maxTotalVolumeTicker = ticker
        End If
        
        ' Increment the output row
        outputRow = outputRow + 1
        
        ' Decrement i to process the row again after exiting the loop
        i = i - 1
    Next i
    ' Insert Table Labels Ticker and Value
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    
    ' Output the stock with the greatest percentage increase
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(2, 17).Value = maxPercentIncreaseTicker
    ws.Cells(2, 18).Value = Format(maxPercentIncrease, "0.00") & "%"
    
    ' Output the stock with the greatest percentage decrease
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(3, 17).Value = maxPercentDecreaseTicker
    ws.Cells(3, 18).Value = Format(maxPercentDecrease, "0.00") & "%"
    
    ' Output the stock with the greatest total volume
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    ws.Cells(4, 17).Value = maxTotalVolumeTicker
    ws.Cells(4, 18).Value = maxTotalVolume
    
    ' Apply conditional formatting for quarterly changes (Column K)
    ApplyConditionalFormatting ws, 11, outputRow - 1
End Sub

Sub ApplyConditionalFormatting(ws As Worksheet, col As Long, lastRow As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))
    
    ' Clear any existing conditional formatting
    rng.FormatConditions.Delete
    
    ' Apply conditional formatting for positive values (green)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Interior.Color = RGB(144, 238, 144) ' Light green color
    End With
    
    ' Apply conditional formatting for negative values (red)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        .Interior.Color = RGB(255, 99, 71) ' Red color
    End With
End Sub
