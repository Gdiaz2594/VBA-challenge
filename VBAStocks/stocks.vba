Sub stocks()

    For a = 1 To Worksheets.Count
    Sheets(a).Select

        Application.ScreenUpdating = False
        
        Dim letter, tickerInc, tickerDec, tickerVol As String
        Dim pos As Integer
        Dim openVal, closeVal, stockVol, maxPercent, minPercent, maxVol As Double
        
        Range("I1").Value = "Ticker"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        pos = 2
        stockVol = 0
        x = Cells(1, 1).End(xlDown).Row
        
        For i = 2 To x
            letter = Cells(i, 1).Value
            stockVol = stockVol + Cells(i, 7).Value
            If letter <> Cells(i - 1, 1).Value Then
                Cells(pos, 9).Value = letter
                openVal = Cells(i, 3).Value
                pos = pos + 1
            End If
            If letter <> Cells(i + 1, 1).Value Then
                closeVal = Cells(i, 6).Value
                Cells(pos - 1, 10).Value = closeVal - openVal
                'Determine color
                If Cells(pos - 1, 10).Value >= 0 Then
                    Cells(pos - 1, 10).Interior.ColorIndex = 4
                Else
                    Cells(pos - 1, 10).Interior.ColorIndex = 3
                End If
                'Percentage change
                If openVal = 0 Then
                    Cells(pos - 1, 11).Value = 0
                Else
                Cells(pos - 1, 11).Value = (closeVal / openVal) - 1
                End If
                    Cells(pos - 1, 11).Select
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                Cells(pos - 1, 12).Value = stockVol
                stockVol = 0
            End If
        Next i
        
        y = Cells(1, 11).End(xlDown).Row
        maxPercent = 0
        minPercent = Cells(2, 11).Value
        maxVol = 0
        'Greatest % increase, Greatest % Decrease and Greatest Total Volume
        For k = 2 To y
            If Cells(k, 11).Value > maxPercent Then
                maxPercent = Cells(k, 11).Value
                tickerInc = Cells(k, 9).Value
            End If
            If Cells(k, 11).Value < minPercent Then
                minPercent = Cells(k, 11).Value
                tickerDec = Cells(k, 9).Value
            End If
            If Cells(k, 12).Value > maxVol Then
                maxVol = Cells(k, 12).Value
                tickerVol = Cells(k, 9).Value
            End If
            
        Next k
        Cells(2, 16).Value = tickerInc
        Cells(2, 17).Value = maxPercent
        Cells(3, 16).Value = tickerDec
        Cells(3, 17).Value = minPercent
        Cells(4, 16).Value = tickerVol
        Cells(4, 17).Value = maxVol
        
    Next a
    
End Sub
