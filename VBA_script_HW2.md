Sub stock_tracker()

'Define stock name
    Dim stock_name As String

'Define entry and exit price
    Dim entry_price As Double
    Dim entry_price_next As Double
    Dim exit_price As Double

'Keep track of row in summary table
    Dim summary_table_row As Double
    summary_table_row = 2

'Keep track of total volume of a stock
    Dim total_vol As Double
    total_vol = 0
    Dim lastRow As Double
    


'To loop through the last row
'For loop for all rows
'    For Each ws In Worksheets
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row + 1
    For i = 2 To lastRow
    
        entry_price = Cells(2, 3).Value
    
        

'Check if we are still in correct credit card
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                stock_name = Cells(i, 1).Value
                entry_price_next = Cells(i + 1, 3).Value
'

                exit_price = Cells(i, 6).Value

                total_vol = total_vol + Cells(i, 7).Value

'Get the difference from entry and exit
                Change = exit_price - entry_price
                If entry_price = 0 Then
                    entry_price = NA
                End If
                percent_change = (exit_price - entry_price) / entry_price


                Range("I" & summary_table_row).Value = stock_name

                Range("J" & summary_table_row).Value = Change
                Range("K" & summary_table_row).Value = percent_change
                Range("L" & summary_table_row).Value = total_vol
        

                summary_table_row = summary_table_row + 1

                total_vol = 0
                entry_price = next_entry_price
            Else

                total_vol = total_vol + Cells(i, 7).Value
            End If
        Next i
End Sub



