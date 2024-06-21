Attribute VB_Name = "Module1"
Sub QuarterStock()

Dim Ticker As String
Dim NextTicker As String
Dim PreviousTicker As String
Dim StockDate As Date
Dim StockVolume As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim List As Variant
Dim Unique As Integer
Dim PercentChange As Double
Dim QtrChange As Double
Dim Qtr As Integer
Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim MaxTotalStockVol As Double
Dim MaxTickerIncrease As Integer
Dim MaxTickerDecrease As Integer
Dim MaxTickerTotalStockVol As Integer

For Qtr = 1 To 4

    Worksheets("Q" & Qtr).Select

    List = Range("A:A")
    Unique = 1

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quarterly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    For i = 2 To UBound(List) - 1

        Ticker = Cells(i, 1).Value
        NextTicker = Cells(i + 1, 1).Value
        PreviousTicker = Cells(i - 1, 1).Value

        If Ticker = PreviousTicker Then
            StockVolume = Cells(i, 7).Value + StockVolume
        Else
            OpenPrice = Cells(i, 3).Value
            StockVolume = Cells(i, 7).Value
        End If

        If Ticker <> NextTicker Then
            ClosePrice = Cells(i, 6).Value
            Unique = Unique + 1
            Cells(Unique, 9) = Ticker
            QtrChange = ClosePrice - OpenPrice
            Cells(Unique, 10) = QtrChange

            If QtrChange > 0 Then
                Cells(Unique, 10).Interior.ColorIndex = 4
            ElseIf QtrChange < 0 Then
                Cells(Unique, 10).Interior.ColorIndex = 3
            End If

            PercentChange = (QtrChange / OpenPrice)
            
            Cells(Unique, 11) = PercentChange
                       
            If PercentChange > 0 Then
                Cells(Unique, 11).Interior.ColorIndex = 4
            ElseIf PercentChange < 0 Then
                Cells(Unique, 11).Interior.ColorIndex = 3
            End If
            
            Range("K1:K5000").NumberFormat = "0.00%"
            
            Cells(Unique, 12) = StockVolume
            StockVolume = 0
            
        End If

    Next i

    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    MaxIncrease = Application.WorksheetFunction.Max(Range("K:K"))
    MaxDecrease = Application.WorksheetFunction.Min(Range("K:K"))
    MaxTotalStockVol = Application.WorksheetFunction.Max(Range("L:L"))
       
    Cells(2, 17).Value = MaxIncrease
    Cells(3, 17).Value = MaxDecrease
    Cells(4, 17).Value = MaxTotalStockVol
    
    MaxTickerIncrease = Application.WorksheetFunction.Match(MaxIncrease, Range("K:K"), 0)
    MaxTickerDecrease = Application.WorksheetFunction.Match(MaxDecrease, Range("K:K"), 0)
    MaxTickerTotalStockVol = Application.WorksheetFunction.Match(MaxTotalStockVol, Range("L:L"), 0)

    Cells(2, 16).Value = Cells(MaxTickerIncrease, 9).Value
    Cells(3, 16).Value = Cells(MaxTickerDecrease, 9).Value
    Cells(4, 16).Value = Cells(MaxTickerTotalStockVol, 9).Value
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    
Next Qtr

End Sub
