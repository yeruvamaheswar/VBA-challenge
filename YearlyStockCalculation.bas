Attribute VB_Name = "Module1"
Sub GetYearlyStockSummary():
    Dim InputWorkSheet As String
    Dim lastrow As Long
    Dim i As Long
    Dim NextTickerStartRow As Long
    Dim NewTickerRow As Long
    Dim TotalStockVolume As Long
    Dim YearOpenStock As Double
    Dim YearCloseStock As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim PercentageChangeFormat As String
    
    
    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        InputWorkSheet = ws.Name
        NewTickerRow = 2
        TotalStockVolume = 0
        YearOpenStock = ws.Cells(2, 3)
        YearCloseStock = 0
        'MsgBox (ws.Name())
        For i = 2 To lastrow
        'For i = 2 To 300
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                ws.Cells(NewTickerRow, 10).Value = ws.Cells(i, 1).Value
                
                If (YearOpenStock = 0) Then
                    'MsgBox (i)
                    'MsgBox (ws.Cells(i, 1).Value)
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
                PercentageChangeFormat = FormatPercent(PercentageChange, 0)
                ws.Cells(NewTickerRow, 12).Value = PercentageChangeFormat
                
                'TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                'ws.Cells(NewTickerRow, 13).Value = TotalStockVolume
                
                
                NewTickerRow = NewTickerRow + 1
                TotalStockVolume = 0
                YearOpenStock = ws.Cells(i + 1, 3).Value
                NextTickerStartRow = i + 1
                YearCloseStock = 0
            Else
                        
                'TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
    Next ws
End Sub
