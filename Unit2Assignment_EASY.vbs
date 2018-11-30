Sub checkStocks()

    Dim counter, i As Long
    Dim ticker As String
    Dim totalVolume As Double

    counter = 1

    Cells(counter, 9) = "Ticker"
    Cells(counter, 10) = "Total Volume"

    For i = 2 To 800000

        totalVolume = totalVolume + Cells(i, 7).Value
        
        ' Searches for when the value of the next cell is different than that of the current cell
        ' Computes total volume, then displays in summary table
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ticker = Cells(i, 1).Value
            counter = counter + 1

            Cells(counter, 9).Value = ticker
            Cells(counter, 10).Value = totalVolume

            totalVolume = 0

        End If
        
    Next i

End Sub