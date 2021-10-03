Attribute VB_Name = "Module1"
Sub GetYearlyStockSummary():
    Dim InputWorkSheet As String
    Dim lastrow As Long
    Dim i As Long
    Dim NextTickerStartRow As Long
    Dim NewTickerRow As Long
    Dim TotalStockVolume As Double
    Dim YearOpenStock As Double
    Dim YearCloseStock As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim PercentageChangeFormat As String
    Dim GreatestIncrease As Double
    Dim GreatestIncreaseFormat As String
    Dim GIticker As String
    Dim GreatestDecrease As Double
    Dim GreatestDecreaseFormat As String
    Dim GDticker As String
    Dim GreatestTotalVolume As Double
    Dim GTVticker As String
   
 
    
    
    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        InputWorkSheet = ws.Name
        NewTickerRow = 2
        TotalStockVolume = 0
        YearOpenStock = ws.Cells(2, 3)
        YearCloseStock = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalVolume = 0
        For i = 2 To lastrow
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                ws.Cells(NewTickerRow, 10).Value = ws.Cells(i, 1).Value
                
                If (YearOpenStock = 0) Then
                    For j = NextTickerStartRow To lastrow
                     YearOpenStock = ws.Cells(j, 3)
                        If (YearOpenStock <> 0) Then
                            Exit For
                        End If
                    Next j
                End If
                YearCloseStock = ws.Cells(i, 6).Value
                YearlyChange = YearCloseStock - YearOpenStock
                ws.Cells(NewTickerRow, 11).Value = YearlyChange
                If (YearlyChange < 0) Then
                    ws.Cells(NewTickerRow, 11).Font.ColorIndex = 1
                    ws.Cells(NewTickerRow, 11).Interior.ColorIndex = 3
                ElseIf (YearlyChange >= 0) Then
                    ws.Cells(NewTickerRow, 11).Font.ColorIndex = 1
                    ws.Cells(NewTickerRow, 11).Interior.ColorIndex = 4
                End If
 
                
                
                PercentageChange = YearlyChange / YearOpenStock
                
                If GreatestIncrease < PercentageChange Then
                    GreatestIncrease = PercentageChange
                    GIticker = ws.Cells(i, 1).Value
                End If

                If GreatestDecrease > PercentageChange Then
                    GreatestDecrease = PercentageChange
                    GDticker = ws.Cells(i, 1).Value
                End If


                PercentageChangeFormat = FormatPercent(PercentageChange, 0)
                ws.Cells(NewTickerRow, 12).Value = PercentageChangeFormat
 
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                ws.Cells(NewTickerRow, 13).Value = TotalStockVolume
                
                
                
                If GreatestTotalVolume < TotalStockVolume Then
                    GreatestTotalVolume = TotalStockVolume
                    GTVticker = ws.Cells(i, 1).Value
                End If
                
                
                
                NewTickerRow = NewTickerRow + 1
                TotalStockVolume = 0
                YearOpenStock = ws.Cells(i + 1, 3).Value
                NextTickerStartRow = i + 1
                YearCloseStock = 0
            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i
        GreatestIncreaseFormat = FormatPercent(GreatestIncrease, 0)
        GreatestDecreaseFormat = FormatPercent(GreatestDecrease, 0)
        ws.Cells(2, 17).Value = GIticker
        ws.Cells(2, 18).Value = GreatestIncreaseFormat

        ws.Cells(3, 17).Value = GDticker
        ws.Cells(3, 18).Value = GreatestDecreaseFormat

        ws.Cells(4, 17).Value = GTVticker
        ws.Cells(4, 18).Value = GreatestTotalVolume
    Next ws
End Sub

