Sub StockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentTicker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim rowIndex As Long

    ' Loop through all sheets
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "QuarterlySummary" Then
            ' Find last row of data
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            ' Clear previous summary data
            ws.Range("I1:L1").Value = Array("Ticker Symbol", "Quarterly Change ($)", "Percentage Change (%)", "Total Stock Volume")
            ws.Range("O2:P4").Value = Array("Ticker", "Value", "Ticker", "Value", "Ticker", "Value")

            ' Initialize variables
            greatestIncrease = -1E+30
            greatestDecrease = 1E+30
            greatestVolume = 0
            rowIndex = 2

            ' Loop through each row of data
            For i = 2 To lastRow
                currentTicker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                closePrice = ws.Cells(i, 6).Value
                totalVolume = ws.Cells(i, 7).Value

                ' Calculate changes
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If

                ' Find greatest increase, decrease, and total volume
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = currentTicker
                End If

                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = currentTicker
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = currentTicker
                End If

                ' Output data to the worksheet
                ws.Cells(rowIndex, 9).Value = currentTicker
                ws.Cells(rowIndex, 10).Value = quarterlyChange
                ws.Cells(rowIndex, 11).Value = percentChange / 100
                ws.Cells(rowIndex, 12).Value = totalVolume

                rowIndex = rowIndex + 1
            Next i

            ' Output greatest values
            ws.Cells(1, 15).Value = "Ticker"
            ws.Cells(1, 16).Value = "Value"
            ws.Cells(2, 14).Value = "Greatest % Increase"
            ws.Cells(2, 15).Value = greatestIncreaseTicker
            ws.Cells(2, 16).Value = greatestIncrease
            
            ws.Cells(3, 14).Value = "Greatest % Decrease"
            ws.Cells(3, 15).Value = greatestDecreaseTicker
            ws.Cells(3, 16).Value = greatestDecrease
            
            ws.Cells(4, 14).Value = "Greatest Total Volume"
            ws.Cells(4, 15).Value = greatestVolumeTicker
            ws.Cells(4, 16).Value = greatestVolume

            ' Apply conditional formatting to the Quarterly Change column
            With ws.Range("J2:J" & rowIndex - 1)
                .FormatConditions.Delete ' Clear any existing formatting
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                .FormatConditions(1).Interior.Color = RGB(144, 238, 144) ' Light green for positive change
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                .FormatConditions(2).Interior.Color = RGB(255, 160, 122) ' Light coral for negative change
            End With

            ' Format Column K (Percentage Change) as percentage
            With ws.Range("K2:K" & rowIndex - 1)
                .NumberFormat = "0.00%"
            End With

        End If
    Next ws

    MsgBox "Quarterly data processing is complete!"
End Sub
