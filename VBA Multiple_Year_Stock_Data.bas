Attribute VB_Name = "Module1"
Sub ProcessFinancialData()

    Dim worksheetData As Worksheet
    Dim stockID As String
    Dim currentCellRange As Range
    Dim symbolDataRange As Range
    Dim lastDataRow As Long
    Dim dataIndex As Long
    Dim initialPrice As Double
    Dim closingPrice As Double
    Dim changeQuarter As Double
    Dim percentageChange As Double
    Dim summaryRowPosition As Integer
    Dim totalTradingVolume As Double
    Dim tableCounterA As Integer
    Dim tableCounterB As Integer
    Dim tableCounterC As Integer
    Dim tableCounterD As Integer
    Dim seriesIndicator As Integer
    Dim minPercentage As Double
    Dim maxPercentage As Double
    Dim highestVol As Double
    Dim topVolumeStockID As String
    Dim topGainStockID As String
    Dim topLossStockID As String

    For Each worksheetData In ThisWorkbook.Worksheets
        If WorksheetFunction.Count(worksheetData.Cells) > 0 Then
            totalTradingVolume = 0
            tableCounterA = 2
            seriesIndicator = 0
            tableCounterB = 2
            tableCounterC = 3
            tableCounterD = 4
            
            minPercentage = 999999999
            maxPercentage = -999999999
            highestVol = 0
            
            summaryRowPosition = 1
            worksheetData.Range("I" & summaryRowPosition).Value = "Stock ID"
            worksheetData.Range("J" & summaryRowPosition).Value = "Quarterly Variation"
            worksheetData.Range("K" & summaryRowPosition).Value = "Percentage Variation"
            worksheetData.Range("L" & summaryRowPosition).Value = "Total Trading Volume"
            
            worksheetData.Range("P" & summaryRowPosition).Value = "Stock ID"
            worksheetData.Range("Q" & summaryRowPosition).Value = "Value"
            
            worksheetData.Range("O" & 2).Value = "Largest % Gain"
            worksheetData.Range("O" & 3).Value = "Largest % Loss"
            worksheetData.Range("O" & 4).Value = "Highest Volume"
            
            lastDataRow = worksheetData.Cells(worksheetData.Rows.Count, "A").End(xlUp).Row
            
            For dataIndex = 2 To lastDataRow
                If worksheetData.Cells(dataIndex + 1, 1).Value = worksheetData.Cells(dataIndex, 1).Value Then
                    If seriesIndicator = 0 Then
                        seriesIndicator = 1
                        initialPrice = worksheetData.Cells(dataIndex, 3).Value
                    End If
                    
                    totalTradingVolume = totalTradingVolume + worksheetData.Cells(dataIndex, 7).Value
                Else
                    closingPrice = worksheetData.Cells(dataIndex, 6).Value
                    totalTradingVolume = totalTradingVolume + worksheetData.Cells(dataIndex, 7).Value
                    changeQuarter = closingPrice - initialPrice
                    worksheetData.Cells(tableCounterA, 10).Value = changeQuarter
                    percentageChange = changeQuarter / initialPrice
                    worksheetData.Cells(tableCounterA, 11).Value = percentageChange
                    worksheetData.Cells(tableCounterA, 11).NumberFormat = "0.00%"
                    worksheetData.Cells(tableCounterA, 12).Value = totalTradingVolume
                    stockID = worksheetData.Cells(dataIndex, 1).Value
                    worksheetData.Cells(tableCounterA, 9).Value = stockID
                    tableCounterA = tableCounterA + 1
                    
                    If percentageChange > maxPercentage Then
                        maxPercentage = percentageChange
                        topGainStockID = stockID
                    End If
                    If percentageChange < minPercentage Then
                        minPercentage = percentageChange
                        topLossStockID = stockID
                    End If
                    If totalTradingVolume > highestVol Then
                        highestVol = totalTradingVolume
                        topVolumeStockID = stockID
                    End If
                    
                    totalTradingVolume = 0
                    seriesIndicator = 0
                    
                    If stockID <> "" Then
                        If WorksheetFunction.CountIf(worksheetData.Range("I:I"), stockID) = 0 Then
                            worksheetData.Cells(worksheetData.Rows.Count, "I").End(xlUp).Offset(1, 0).Value = stockID
                        End If
                    End If
                End If
            Next dataIndex
            
            worksheetData.Cells(tableCounterB, 16).Value = topGainStockID
            worksheetData.Cells(tableCounterB, 17).Value = maxPercentage
            worksheetData.Cells(tableCounterB, 17).NumberFormat = "0.00%"
            
            worksheetData.Cells(tableCounterC, 16).Value = topLossStockID
            worksheetData.Cells(tableCounterC, 17).Value = minPercentage
            worksheetData.Cells(tableCounterC, 17).NumberFormat = "0.00%"
            
            worksheetData.Cells(tableCounterD, 16).Value = topVolumeStockID
            worksheetData.Cells(tableCounterD, 17).Value = highestVol
        End If
    Next worksheetData

End Sub
