Sub Hardpart()
For Each ws In Worksheets
    Dim Ticker As String
    Dim Total_Volume As Double
    Dim Tickerrow As Integer
    Dim Openvalue As String
    Dim Openvaluenext As Double
    Dim Closevalue As String
    Dim Percent_change As Double
    Dim Greatest_increase As Double
    Dim Greatest_decrease As Double
    Dim Greatest_total_volume As Double
    Dim Greatest_ticker As String
    Dim Lowest_ticker As String
    Dim Greatest_volume_ticker As String
    Tickerrow = 2
    Openvalue = ws.Cells(2, 3).Value
    Total_Volume = ws.Cells(2, 7).Value
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To Lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            Closevalue = ws.Cells(i, 6).Value
            Openvaluenext = ws.Cells(i + 1, 3).Value
            ws.Range("I" & Tickerrow).Value = Ticker
            ws.Range("J" & Tickerrow).Value = Total_Volume
            ws.Range("K" & Tickerrow).Value = Closevalue - Openvalue
            If Openvalue > 0 Then
                Percent_change = ws.Range("K" & Tickerrow).Value / (Openvalue) * 100
            Else: Percent_change = 0
            End If
            ws.Range("L" & Tickerrow).Value = Percent_change
            Tickerrow = Tickerrow + 1
            Ticker = 0
            Total_Volume = 0
            Openvalue = Openvaluenext
            Closevalue = 0
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        End If
    Next i
    For i = 2 To Tickerrow
        If ws.Cells(i, 12).Value > 0 Then
            ws.Range("K" & i).Interior.ColorIndex = 4
        Else: ws.Range("K" & i).Interior.ColorIndex = 3
        End If
    Next i
    Greatest_increase = 0
    Greatest_ticker = 0
    For i = 2 To Tickerrow
        If ws.Range("L" & i).Value > Greatest_increase Then
            Greatest_increase = ws.Range("L" & i).Value
            Greatest_ticker = ws.Range("I" & i).Value
        End If
    Next i
    ws.Range("P2").Value = Greatest_increase
    ws.Range("O2").Value = Greatest_ticker
    Greatest_decrease = 0
    Lowest_ticker = 0
    For i = 2 To Tickerrow
        If ws.Range("L" & i).Value < Greatest_decrease Then
            Greatest_decrease = ws.Range("L" & i).Value
            Lowest_ticker = ws.Range("I" & i).Value
        End If
    Next i
    ws.Range("P3").Value = Greatest_decrease
    ws.Range("O3").Value = Lowest_ticker
    Greatest_total_volume = 0
    Greatest_volume_ticker = 0
    For i = 2 To Tickerrow
        If ws.Range("J" & i).Value > Greatest_total_volume Then
            Greatest_total_volume = ws.Range("J" & i).Value
            Greatest_volume_ticker = ws.Range("I" & i).Value
        End If
    Next i
    ws.Range("P4").Value = Greatest_total_volume
    ws.Range("O4").Value = Greatest_volume_ticker
Next ws
End Sub

