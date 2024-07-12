Attribute VB_Name = "Module1"
Sub AddTickerSymbol()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tickerSymbol As String
    Dim tickerRow As Long ' Track the row in column H to place the ticker symbol
    
    ' Loop through each sheet named Q1, Q2, Q3, and Q4
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            ' Initialize tickerRow for column H starting from row 2
            tickerRow = 2
            
            ' Find the last row with data in column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Loop through each row in the current worksheet
            For i = 2 To lastRow
                tickerSymbol = ws.Cells(i, "A").Value
                
                ' Check if the current ticker symbol is different from the previous one
                If i = 2 Or tickerSymbol <> ws.Cells(i - 1, "A").Value Then
                    ' Assign the ticker symbol to column H in the current tickerRow
                    ws.Cells(tickerRow, "H").Value = tickerSymbol
                    ' Move to the next row in column H
                    tickerRow = tickerRow + 1
                End If
            Next i
        End If
    Next ws
End Sub

Sub CalculateQuarterlyChange()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim tickerRow As Long ' Track the row in column I to place the quarterly change
    Dim startRow As Long ' Track the row where each ticker's data starts in each sheet
    Dim quarterEndDate As Date
    Dim openPrice As Double
    Dim closePrice As Double
    Dim change As Double
    
    ' Array of worksheet names
    Dim quarters As Variant
    quarters = Array("Q1", "Q2", "Q3", "Q4")
    
    ' Quarter end dates
    Dim quarterEndDates As Variant
    quarterEndDates = Array("3/31/2022", "6/30/2022", "9/30/2022", "12/31/2022")
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not IsError(Application.Match(ws.Name, quarters, 0)) Then
            Debug.Print "Processing worksheet: " & ws.Name
            
            ' Find the last row with data in Column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            Debug.Print "Last Row: " & lastRow
            
            ' Initialize variables for ticker and starting row
            ticker = ""
            tickerRow = 2 ' Start from row 2 in column I for each ticker
            startRow = 2 ' Assuming data starts from row 2
            
            ' Determine quarter end date
            quarterEndDate = DateValue(quarterEndDates(Application.Match(ws.Name, quarters, 0) - 1))
            Debug.Print "Quarter End Date: " & quarterEndDate
            
            ' Loop through each row
            For i = 2 To lastRow
                ' Check if we've reached the end of the quarter
                If ws.Cells(i, "B").Value = quarterEndDate Then
                    Debug.Print "Found end of quarter at row: " & i
                    
                    ' Calculate the change
                    closePrice = ws.Cells(i, "F").Value
                    openPrice = ws.Cells(startRow, "C").Value
                    
                    ' Ensure open price is valid (not zero)
                    If openPrice <> 0 Then
                        change = closePrice - openPrice
                    Else
                        change = 0
                    End If
                    
                    ' Write the change to column I for the current tickerRow
                    ws.Cells(tickerRow, "I").Value = Round(change, 3)
                    Debug.Print "Writing change to cell: I" & tickerRow
                    
                    ' Apply color based on the change
                    If change > 0 Then
                        ws.Cells(tickerRow, "I").Interior.Color = RGB(144, 238, 144) ' Light Green
                    ElseIf change < 0 Then
                        ws.Cells(tickerRow, "I").Interior.Color = RGB(255, 99, 71) ' Light Red
                    Else
                        ws.Cells(tickerRow, "I").Interior.ColorIndex = xlNone ' No color
                    End If
                    
                    ' Move to the next row in column I for the next ticker symbol
                    tickerRow = tickerRow + 1
                    
                    ' Update startRow for the next ticker symbol's data
                    startRow = i + 1
                End If
            Next i
        End If
    Next ws
End Sub

Sub CalculateQuarterlyPercentChange()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim tickerRow As Long ' Track the row in column J to place the percent change
    Dim startRow As Long ' Track the row where each ticker's data starts in each sheet
    Dim quarterEndDate As Date
    Dim openPrice As Double
    Dim closePrice As Double
    Dim percentChange As Double
    
    ' Array of worksheet names
    Dim quarters As Variant
    quarters = Array("Q1", "Q2", "Q3", "Q4")
    
    ' Quarter end dates
    Dim quarterEndDates As Variant
    quarterEndDates = Array("3/31/2022", "6/30/2022", "9/30/2022", "12/31/2022")
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not IsError(Application.Match(ws.Name, quarters, 0)) Then
            Debug.Print "Processing worksheet: " & ws.Name
            
            ' Find the last row with data in Column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            Debug.Print "Last Row: " & lastRow
            
            ' Initialize variables for ticker and starting row
            ticker = ""
            tickerRow = 2 ' Start from row 2 in column J for each ticker
            startRow = 2 ' Assuming data starts from row 2
            
            ' Determine quarter end date
            quarterEndDate = DateValue(quarterEndDates(Application.Match(ws.Name, quarters, 0) - 1))
            Debug.Print "Quarter End Date: " & quarterEndDate
            
            ' Loop through each row
            For i = 2 To lastRow
                ' Check if we've reached the end of the quarter
                If ws.Cells(i, "B").Value = quarterEndDate Then
                    Debug.Print "Found end of quarter at row: " & i
                    
                    ' Calculate the percent change
                    closePrice = ws.Cells(i, "F").Value
                    openPrice = ws.Cells(startRow, "C").Value
                    
                    ' Ensure open price is valid (not zero)
                    If openPrice <> 0 Then
                        percentChange = ((closePrice - openPrice) / openPrice)
                    Else
                        percentChange = 0
                    End If
                    
                    ' Write the percent change as percentage to column J for the current tickerRow
                    ws.Cells(tickerRow, "J").NumberFormat = "0.00%;-0.00%"
                    ws.Cells(tickerRow, "J").Value = percentChange
                    Debug.Print "Writing percent change to cell: J" & tickerRow
                    
                    ' Move to the next row in column J for the next ticker symbol
                    tickerRow = tickerRow + 1
                    
                    ' Update startRow for the next ticker symbol's data
                    startRow = i + 1
                End If
            Next i
        End If
    Next ws
End Sub

Sub CalculateTotalVolumePerTicker()
    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim currentRow As Long
    Dim totalVolume As Double
    Dim outputRow As Long
    
    ' Loop through each worksheet (Q1, Q2, Q3, Q4)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q?" Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            currentRow = 2 ' Start after the headers
            outputRow = 2 ' Output starts at row 2 in column K
            
            Do While currentRow <= lastRow
                ' Get the ticker name
                ticker = ws.Cells(currentRow, 1).Value
                totalVolume = 0
                
                ' Sum the volumes for the current ticker
                Do While ws.Cells(currentRow, 1).Value = ticker And currentRow <= lastRow
                    totalVolume = totalVolume + ws.Cells(currentRow, 7).Value
                    currentRow = currentRow + 1
                Loop
                
                ' Output the result to column K
                ws.Cells(outputRow, 11).Value = totalVolume
                outputRow = outputRow + 1
            Loop
        End If
    Next ws
    
    MsgBox "Total volume calculation is complete!"
End Sub

Sub CalculateGreatestValues()
    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim currentRow As Long
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestTotalVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestTotalVolumeTicker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim percentChange As Double
    
    greatestIncrease = -1
    greatestDecrease = 1
    greatestTotalVolume = 0
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in Column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        currentRow = 2 ' Start after the headers
        
        ' Loop through each row in the worksheet
        Do While currentRow <= lastRow
            ' Get the ticker name
            ticker = ws.Cells(currentRow, 1).Value
            totalVolume = 0
            openPrice = ws.Cells(currentRow, 3).Value
            
            ' Sum the volumes for the current ticker and find close price
            Do While ws.Cells(currentRow, 1).Value = ticker And currentRow <= lastRow
                totalVolume = totalVolume + ws.Cells(currentRow, 7).Value
                closePrice = ws.Cells(currentRow, 6).Value ' Column F for closing price
                currentRow = currentRow + 1
            Loop
            
            ' Calculate percentage change
            If openPrice <> 0 Then
                percentChange = ((closePrice - openPrice) / openPrice) * 100
            Else
                percentChange = 0
            End If
            
            ' Find greatest increase
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            ' Find greatest decrease
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            ' Find greatest total volume
            If totalVolume > greatestTotalVolume Then
                greatestTotalVolume = totalVolume
                greatestTotalVolumeTicker = ticker
            End If
        Loop
    Next ws
    
    ' Output results to column N, O, P
    With ThisWorkbook.Sheets(1)
        .Cells(2, 14).Value = "Greatest % Increase"
        .Cells(3, 14).Value = "Greatest % Decrease"
        .Cells(4, 14).Value = "Greatest Total Volume"
        
        .Cells(2, 15).Value = greatestIncreaseTicker
        .Cells(3, 15).Value = greatestDecreaseTicker
        .Cells(4, 15).Value = greatestTotalVolumeTicker
        
        .Cells(2, 16).Value = Round(greatestIncrease, 2) & "%" ' Already a percentage, so no multiplication needed
        .Cells(3, 16).Value = Round(greatestDecrease, 2) & "%" ' Already a percentage, so no multiplication needed
        .Cells(4, 16).Value = Format(greatestTotalVolume, "0.00E+00")
    End With
    
    MsgBox "Greatest values calculation is complete!"
End Sub


