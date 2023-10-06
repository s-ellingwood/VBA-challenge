Sub stocksorting():

    ' Defining variables to store number of current worksheet and total number of worksheets in the workbook, respectively
    Dim n As Integer
    Dim c As Integer
    
    ' Finding and storing the total number of worksheets in this workbook
    c = ActiveWorkbook.Worksheets.Count
    
    ' Making a For loop to loop through all worksheets in the workbook
    For n = 1 To c Step 1
        ' Setting the active worksheet to the current worksheet (n)
        Worksheets(n).Activate
        
        ' Setting column titles
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Value"
        
        ' Finding and storing the total number of rows that need to be iterated through
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Defining a variable to store the ticker value and storing the first ticker value (AAB) in the first row of our Ticker column
        ticker = Cells(2, 1).Value
        Range("I2").Value = ticker
        
        ' Defining the variables to be used during iteration
        Dim opendifference As Double
        Dim daydifference As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim totalstockvalue As Variant
        
        ' Setting new variables to 0
        opendifference = 0#
        daydifference = 0#
        yearlychange = 0#
        percentchange = 0#
        totalstockvalue = CDec(0)
        
        ' Creating For loop to iterate through every row in the <ticker> column
        For i = 2 To lastrow
            ' Creating If statement to store and calculate values if the value in <ticker> matches the value in our Ticker column
            If ticker = Cells(i, 1).Value Then
                ' Creating If statement that will ignore the value in the cell before the current cell if their <ticker> values don't match
                If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                    opendifference = 0
                ' The Else statement will store and calculate values only if the <ticker> values of the current cell and previous cell match
                ' Assures that the numbers are only calculated if they're part of the same stock/ticker
                Else
                    ' Calculating the yearly change:
                    ' opendifference is used to calculate the difference between the current open value and the previous open value
                    opendifference = opendifference + (Cells(i, 3).Value - Cells(i - 1, 3).Value)
                    ' daydifference calculate the difference between the current open value and the current closing value
                    daydifference = Cells(i, 6).Value - Cells(i, 3).Value
                    ' yearlychange then uses the values of daydifference and open difference to calculate the difference between the first open value and the last closing value
                    yearlychange = daydifference + opendifference
                    
                    ' Calculting the percent change:
                    ' percentchange subtracts 1 from value of the last closing value over the first open value to calcualte percentage increase or decrease
                    percentchange = (Cells(i, 6) / (Cells(i, 6) - yearlychange)) - 1
                    
                    ' Calculating totalstock value:
                    ' totalstockvalue adds all the values in the <volume> column for the current ticker value
                    totalstockvalue = CDec(totalstockvalue + Cells(i, 7).Value)
                End If
            
            ' The ElseIf statement is applied when the ticker variable's value doesn't match the current rows <ticker> value (the ticker has changed)
            ElseIf ticker <> Cells(i, 1).Value Then
                ' Finds the row number of the next empty row in the J column
                NextRow = Range("J" & Rows.Count).End(xlUp).Row + 1
                
                ' Finds the row number of the last row with text in it in the I column
                CurrentRow = Range("I" & Rows.Count).End(xlUp).Row
                
                ' Sets value the next empty row in the J column equal to yearlychange variable
                Cells(NextRow, 10).Value = yearlychange
                ' Conditional color coding for the yearlychange value; green if positive and red if negative
                If Cells(CurrentRow, 10) >= 0 Then
                    Cells(CurrentRow, 10).Interior.ColorIndex = 4
                ElseIf Cells(CurrentRow, 10) < 0 Then
                    Cells(CurrentRow, 10).Interior.ColorIndex = 3
                End If
                
                ' Sets the value of the next empty row in the K column equal to percentchange variable
                Cells(NextRow, 11).Value = FormatPercent(percentchange, 2)
                ' Conditional color coding for the percentchange value; green if positive and red if negative
                If Cells(CurrentRow, 11) >= 0 Then
                    Cells(CurrentRow, 11).Interior.ColorIndex = 4
                ElseIf Cells(CurrentRow, 11) < 0 Then
                    Cells(CurrentRow, 11).Interior.ColorIndex = 3
                End If
                
                ' Sets the value of the next empty row in the L column equal to totalstock variable
                Cells(NextRow, 12).Value = totalstockvalue
                
                ' Changes the ticker variable to match the value of the current rows <ticker> value
                ticker = Cells(i, 1).Value
                
                ' Updates NextRow to be equal to that of the next empty row in the I (Ticker) column
                NextRow = Range("I" & Rows.Count).End(xlUp).Row + 1
                
                ' Sets the value of the next empty row in the I column equal to the ticker variable
                Cells(NextRow, 9).Value = ticker
                
                ' Resets all variables to be equal to 0
                opendifference = 0
                daydifference = 0
                yearlychange = 0
                percentchange = 0
                totalstockvalue = 0
            End If
        ' Increase i value by 1
        Next i
        
        ' Repeats the code from above for the last ticker value in the loop
        ' The above For loop closes without populating these onto the worksheet because there's no other ticker value to trigger the ElseIf statement
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
        
        
        ' Defining the new variable we will use in calculating Greatest % Increase, % Decrease, and Total Volume
        Dim pctincrease As Double
        Dim tickerinc As String
        Dim pctdecrease As Double
        Dim tickerdec As String
        Dim totalvolume As Variant
        Dim tickervol As String
        
        ' Setting our Row titles
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        ' Setting our column titles
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        ' Updating lastrow to be equal to the number of rows with values in our Ticker column
        lastrow = Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Setting all of our variables to 0 or, for string variables, to empty
        pctincrease = 0
        pctdecrease = 0
        totalvolume = CDec(0)
        tickerinc = ""
        tickerdec = ""
        tickervol = ""
        
        ' Creating a For loop to iteratre through all of the rows in our newly created table (Columns I through L)
        For i = 2 To lastrow
            ' If the value in the Percent Change column is greater than the current value of the pctincrease variable, then
            If Cells(i, 11).Value > pctincrease Then
                ' Updates pctincrease variable's value to the new greatest Percent Change value
                pctincrease = Cells(i, 11).Value
                ' Updates the ticker value for the greatest increase to the current row's ticker value
                tickerinc = Cells(i, 9).Value
            End If
            
            ' If the value in the Percent Change column is lower than the current value of the pctdecrease variable, then
            If Cells(i, 11).Value < pctdecrease Then
                ' Updates pctdecrease variable's value to the new lowest Percent Change value
                pctdecrease = Cells(i, 11).Value
                ' Updates the ticker value for the greatest decrease to the current row's ticker value
                tickerdec = Cells(i, 9).Value
            End If
            
            ' If the value in Total Stock Value is greater than the current value of the totalvolume variable, then
            If Cells(i, 12).Value > totalvolume Then
                ' Updates totalvolume variable's value to the new greatest Total Stock Value value
                totalvolume = CDec(Cells(i, 12).Value)
                ' Updates the ticker value for the greatest total volume to the current row's ticker value
                tickervol = Cells(i, 9).Value
            End If
        ' Increases the value of i by 1
        Next i
        
        ' Populates all of the values onto the worksheet in the proper cells
        Range("P2") = tickerinc
        Range("Q2") = FormatPercent(pctincrease)
        Range("P3") = tickerdec
        Range("Q3") = FormatPercent(pctdecrease)
        Range("P4") = tickervol
        Range("Q4") = totalvolume
    ' Increases the value of n by 1 to iterate to the next worksheet
    Next n
End Sub