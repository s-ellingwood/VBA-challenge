Sub stocksorting():
    Dim n As Integer
    Dim c As Integer
    
    c = ActiveWorkbook.Worksheets.Count
    For n = 1 To c Step 1
        Worksheets(n).Activate
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Value"
        
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ticker = Cells(2, 1).Value
        Range("I2").Value = ticker
        
        Dim opendifference As Double
        Dim daydifference As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim totalstockvalue As Variant
        
        opendifference = 0#
        daydifference = 0#
        yearlychange = 0#
        totalstockvalue = CDec(0)
        
        For i = 2 To lastrow
            If ticker = Cells(i, 1).Value Then
                If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                    opendifference = 0
                Else
                    opendifference = opendifference + (Cells(i, 3).Value - Cells(i - 1, 3).Value)
                    daydifference = Cells(i, 6).Value - Cells(i, 3).Value
                    yearlychange = daydifference + opendifference
                    percentchange = (Cells(i, 6) / (Cells(i, 6) - yearlychange)) - 1
                    totalstockvalue = CDec(totalstockvalue + Cells(i, 7).Value)
                End If
            ElseIf ticker <> Cells(i, 1).Value Then
                NextRow = Range("J" & Rows.Count).End(xlUp).Row + 1
                CurrentRow = Range("I" & Rows.Count).End(xlUp).Row
                Cells(NextRow, 10).Value = yearlychange
                If Cells(CurrentRow, 10) >= 0 Then
                    Cells(CurrentRow, 10).Interior.ColorIndex = 4
                ElseIf Cells(CurrentRow, 10) < 0 Then
                    Cells(CurrentRow, 10).Interior.ColorIndex = 3
                End If
                Cells(NextRow, 11).Value = FormatPercent(percentchange, 2)
                If Cells(CurrentRow, 11) >= 0 Then
                    Cells(CurrentRow, 11).Interior.ColorIndex = 4
                ElseIf Cells(CurrentRow, 11) < 0 Then
                    Cells(CurrentRow, 11).Interior.ColorIndex = 3
                End If
                Cells(NextRow, 12).Value = totalstockvalue
                ticker = Cells(i, 1).Value
                NextRow = Range("I" & Rows.Count).End(xlUp).Row + 1
                Cells(NextRow, 9).Value = ticker
                opendifference = 0
                daydifference = 0
                yearlychange = 0
                percentchange = 0
                totalstockvalue = 0
            End If
        Next i
        NextRow = Range("J" & Rows.Count).End(xlUp).Row + 1
        CurrentRow = Range("I" & Rows.Count).End(xlUp).Row
        Cells(NextRow, 10).Value = yearlychange
        If Cells(CurrentRow, 10) >= 0 Then
            Cells(CurrentRow, 10).Interior.ColorIndex = 4
        ElseIf Cells(CurrentRow, 10) < 0 Then
            Cells(CurrentRow, 10).Interior.ColorIndex = 3
        End If
        Cells(NextRow, 11).Value = FormatPercent(percentchange, 2)
        If Cells(CurrentRow, 11) >= 0 Then
            Cells(CurrentRow, 11).Interior.ColorIndex = 4
        ElseIf Cells(CurrentRow, 11) < 0 Then
            Cells(CurrentRow, 11).Interior.ColorIndex = 3
        End If
        Cells(NextRow, 12).Value = totalstockvalue
        
        Dim pctincrease As Double
        Dim tickerinc As String
        Dim pctdecrease As Double
        Dim tickerdec As String
        Dim totalvolume As Variant
        Dim tickervol As String
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        lastrow = Cells(Rows.Count, 9).End(xlUp).Row
        pctincrease = 0
        pctdecrease = 0
        totalvolume = CDec(0)
        tickerinc = ""
        tickerdec = ""
        tickervol = ""
        
        For i = 2 To lastrow
            If Cells(i, 11).Value > pctincrease Then
                pctincrease = Cells(i, 11).Value
                tickerinc = Cells(i, 9).Value
            End If
            If Cells(i, 11).Value < pctdecrease Then
                pctdecrease = Cells(i, 11).Value
                tickerdec = Cells(i, 9).Value
            End If
            If Cells(i, 12).Value > totalvolume Then
                totalvolume = CDec(Cells(i, 12).Value)
                tickervol = Cells(i, 9).Value
            End If
        Next i
        
        Range("P2") = tickerinc
        Range("Q2") = FormatPercent(pctincrease)
        Range("P3") = tickerdec
        Range("Q3") = FormatPercent(pctdecrease)
        Range("P4") = tickervol
        Range("Q4") = totalvolume
    Next n
End Sub
