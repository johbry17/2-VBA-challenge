Sub main()
    ' iterates over every worksheet, calls DoEverything
    ' in hindsight, I should've broken DoEverything up into multiple sub functions...
    For Each Worksheet In ThisWorkbook.Sheets
        Worksheet.Activate
        Call DoEverything
    Next Worksheet
End Sub
Sub DoEverything()
    ' Format spreadsheet / Assign names to cells
    [I1] = "Ticker"
    [J1] = "Yearly Change"
    [K1] = "Percentage Change"
    [L1] = "Total Stock Volume"
    [O2] = "Greatest % Increase"
    [O3] = "Greatest % Decrease"
    [O4] = "Greatest Total Volume"
    [P1] = "Ticker"
    [Q1] = "Value"
    
    ' Format column width for aesthetics
    Columns("A:Q").AutoFit
    
    ' Declare variables
    Dim ticker As String: ticker = Cells(2, 1).Value
    Dim sheet_length As Double: sheet_length = Cells(Rows.Count, "A").End(xlUp).Row
    ' row_counter is used to create each new row
    Dim row_counter As Long: row_counter = 2
    Dim opening_price As Double: opening_price = Cells(2, 3).Value
    Dim closing_price As Double:
    Dim opening_volume As Range: Set opening_volume = Cells(2, 7)
    Dim closing_volume As Range

    ' loop to populate totals for table of ticker, yearly change, percent change, and total stock volume columns
    ' sets first value in ticker column
    Cells(row_counter, 9).Value = ticker
    ' Loop over <ticker> column
    ' iterate to (sheet_length + 1) to populate the final row of the final ticker
    For i = 2 To (sheet_length + 1)
        ' sets cell to the value in <ticker>
        Dim cell As String: cell = Cells(i, 1).Value
        
        ' compares value in <ticker> to value in ticker column
        ' if not equal, finishes the current row, sets new opening variables, starts a new row
        If Not cell = ticker Then
            ' set closing_price, closing volume variables, populates yearly change, percent change, total stock volume cells
            closing_price = Cells(i - 1, 6).Value
            Set closing_volume = Cells(i - 1, 7)
            Cells(row_counter, 10).Value = closing_price - opening_price
            Cells(row_counter, 11).Value = (closing_price - opening_price) / opening_price
            Cells(row_counter, 12).Value = WorksheetFunction.Sum(Range(opening_volume.Address(), closing_volume.Address()))
            
            ' sets new opening_price, opening_volume
            opening_price = Cells(i, 3).Value
            Set opening_volume = Cells(i, 7)
            
            ' increases row_counter, sets new ticker variable, starts new row with new ticker
            row_counter = row_counter + 1
            ticker = cell
            Cells(row_counter, 9).Value = ticker
        End If
    Next i
    
    ' set variables for formatting table
    Dim table_length As Double: table_length = Cells(Rows.Count, 9).End(xlUp).Row
    Dim column_range As Range
    
    ' formats yearly change column
    Set column_range = Range("J2:J" & table_length)
    column_range.NumberFormat = "0.00"
    
    ' formats percentage change column
    Set column_range = Range("K2:K" & table_length)
    column_range.NumberFormat = "0.00%"
    
    ' formats yearly change column as red and green for decrease / increase
    ' i do not like the colors, but am probably too lazy to change them
    For i = 2 To table_length
        If Cells(i, 10).Value >= 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    ' declare variables for table of greatest increase, decrease, and total volume
    Dim increase As Double: increase = Cells(2, 11).Value
    Dim decrease As Double: decrease = Cells(2, 11).Value
    Dim volume As Double: volume = Cells(2, 12).Value
    ' initially populate ticker and value columns (P2:Q4) for greatest increase, decrease, and total volume
    Cells(2, 16).Value = Cells(2, 9).Value
    Cells(3, 16).Value = Cells(2, 9).Value
    Cells(4, 16).Value = Cells(2, 9).Value
    Cells(2, 17).Value = increase
    Cells(3, 17).Value = decrease
    Cells(4, 17).Value = volume
    
    ' iterate down rows, uses a conditional to check against our variable, updates P2:Q4
    For i = 2 To table_length
        If Cells(i, 11) > increase Then
            increase = Cells(i, 11).Value
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = increase
        ElseIf Cells(i, 11).Value < decrease Then
            decrease = Cells(i, 11).Value
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = decrease
        ElseIf Cells(i, 12).Value > volume Then
            volume = Cells(i, 12).Value
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = volume
        End If
    Next i
    
    ' formats percentages for greatest increase and decrease as a percentage
    Set column_range = Range("Q2:Q3")
    column_range.NumberFormat = "0.00%"

End Sub
