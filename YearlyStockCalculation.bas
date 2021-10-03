Attribute VB_Name = "Module11"
'Subroutine to get the yearly stock summary
Sub GetYearlyStockSummary():

    'Declaring Variables.
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
   
 
    
    'For Loop which loops thorough all the work sheets.
    For Each ws In Worksheets
        'Vatiable to get the last row of the current sheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        InputWorkSheet = ws.Name
        NewTickerRow = 2
        TotalSto0ckVolume = 0
        YearOpenStock = ws.Cells(2, 3)
        YearCloseStock = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalVolume = 0
        'For loop for looping the current sheet.
        For i = 2 To lastrow
            'If condition to check for change of ticker.
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                ws.Cells(NewTickerRow, 10).Value = ws.Cells(i, 1).Value
                'If condition to check for any stock which is having zero values in the stock open coloumn.
                If (YearOpenStock = 0) Then
                    For j = NextTickerStartRow To lastrow
                     YearOpenStock = ws.Cells(j, 3)
                        If (YearOpenStock <> 0) Then
                        'Exciting the loop once it found the first no zero value.
                            Exit For
                        End If
                    Next j
                End If

                YearCloseStock = ws.Cells(i, 6).Value
                'Yearly CHange calculation.
                YearlyChange = YearCloseStock - YearOpenStock
                ws.Cells(NewTickerRow, 11).Value = YearlyChange
                'Assigning colors depending on the +ve or -ve values.
                If (YearlyChange < 0) Then
                    'Black Font/Red Cell
                    ws.Cells(NewTickerRow, 11).Font.ColorIndex = 1
                    ws.Cells(NewTickerRow, 11).Interior.ColorIndex = 3
                ElseIf (YearlyChange >= 0) Then
                    'Black Font/Green Cell
                    ws.Cells(NewTickerRow, 11).Font.ColorIndex = 1
                    ws.Cells(NewTickerRow, 11).Interior.ColorIndex = 4
                End If
 
                
                'Percentage Change calculation
                PercentageChange = YearlyChange / YearOpenStock
                'Condition to check the Greatest Increase in Stock for Bounus.
                If GreatestIncrease < PercentageChange Then
                    GreatestIncrease = PercentageChange
                    GIticker = ws.Cells(i, 1).Value
                End If
                'Condition to check the Greatest Decrease in Stock for Bounus.
                If GreatestDecrease > PercentageChange Then
                    GreatestDecrease = PercentageChange
                    GDticker = ws.Cells(i, 1).Value
                End If

                'Changing the format to percentage.
                PercentageChangeFormat = FormatPercent(PercentageChange, 0)
                ws.Cells(NewTickerRow, 12).Value = PercentageChangeFormat

                'Calculating Total Stock Volume.
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                ws.Cells(NewTickerRow, 13).Value = TotalStockVolume
                
                'Condition to check the Greatest Total volume of Stock of current ticker for Bounus.
                If GreatestTotalVolume < TotalStockVolume Then
                    GreatestTotalVolume = TotalStockVolume
                    GTVticker = ws.Cells(i, 1).Value
                End If
                
                
                'operation to get where the values of next ticker has to be.
                NewTickerRow = NewTickerRow + 1
                TotalStockVolume = 0
                YearOpenStock = ws.Cells(i + 1, 3).Value
                NextTickerStartRow = i + 1
                YearCloseStock = 0
            Else
                'Totaling all the stock volume traded for the current ticker.
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i
        'Operation to change format to percentages as Bounus.
        GreatestIncreaseFormat = FormatPercent(GreatestIncrease, 0)
        GreatestDecreaseFormat = FormatPercent(GreatestDecrease, 0)
        'Entering the values for Bounus.
        ws.Cells(2, 17).Value = GIticker
        ws.Cells(2, 18).Value = GreatestIncreaseFormat

        ws.Cells(3, 17).Value = GDticker
        ws.Cells(3, 18).Value = GreatestDecreaseFormat

        ws.Cells(4, 17).Value = GTVticker
        ws.Cells(4, 18).Value = GreatestTotalVolume
    Next ws
End Sub

