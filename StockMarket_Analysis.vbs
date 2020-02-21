Sub StockMarket_Analysis()
' Create a script that will loop through all the stocks for one year for each run and take the following information.
' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.
' You should also have conditional formatting that will highlight positive change in green and negative change in red.

    ' Declare Variables
    Dim ws As Worksheet
    Dim LastRow As Long ' Long is a large or "long" integer or counting number
    Dim i As Long ' I will use this to refer to rows in my worksheet
    Dim summary_row As Long ' counter for summary table
    Dim total_stock_volume
    Dim yearly_change ' the yearly change is the closeing price the year - opening price for the year
    Dim start
    Dim percent_change ' this is the yearly change divided by the opening price * 100
' ------------------------------------------------------------------------------------
    ' Interate through each worksheet (ws)
    For Each ws in worksheets

    ' Initalize variables for each worksheet
    LastRow = 0 ' LastRow will find the last row in the worksheet (want it to be zero for the start of each ws)

    summary_row = 2

    total_stock_volume = 0

    yearly_change = 0

    start = 2

    percent_change = 0
   
    ' Identify the last row
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).row ' this code is from an activity in Day 3 called Wells Fargo
    ' test message to identify the macro is moving to the next worksheet
    'MsgBox (LastRow)

        ' Create Summary Table Headers
        ' On the same worksheet as the rawdata,
        ' allcolumns were correctly created for:
        ' ✓​ticker symbol✓​total stock volume✓​yearly change ($)✓​percent change

        ws.Range("I1") = "ticker symbol"
        ws.Range("J1") = "yearly change($)"
        ws.Range("K1") = "percent change(%)"
        ws.Range("L1") = "total stock volume"

    '---------------------------------------------------------------------------------
        For i = 2 to LastRow

            ' Check for stock with no trading for the year
            ' This is case where total_stock_volume is zero
            If ws.Cells(lastrow, 7) = 0 then
            'msgBox("total stock volume was zero for the year")

                    ' Record summary table results for zero change for the year
                    ws.Range("I" & summary_row) = ws.Cells(i, 1)
                    ws.Range("J" & summary_row) = 0
                    ws.Range("K" & summary_row) = 0 & "%"
                    ws.Range("L" & summary_row) = 0     

            Else

                ' If the ticker symbol changes then record findings in summary table
                ' Note column 1 is the location of the ticker symbol
                If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) then

                    ' Store total_stock_volume in summary table
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7)

                    ' Check if total stock volume is zero
                    If total_stock_volume = 0 then
                    
                
                    Else

                    ' Record ticker symbol in summary table (column I)
                    ' Note: need a counter to keep track of the summary table row.
                    ' Use summary_row starting on row 2
                    ws.Range("I" & summary_row) = ws.Cells(i, 1)


                    ' Calculate the yearly_change where yearly_change = closing price - opening price
                    ' Note the closing price is in cells(i, 6)
                    ' Note start is "first" cells(2, 3) and then moves to cells(start, 3)
                    ' Where start is the first opening price for the new ticker symbol cells(i + 1, 3)

                        ' Need to check for an opening price that is not zero
                        ' If opening price is zero then ...
                        If cells(start, 3) = 0 then

                            ' Interate through rows until the first non-zero is found
                            For new_start = start to LastRow ' new_start and start are just row counters
                                'If new open price is not zero then ...
                                If Cells(new_start, 3) <> 0 then
                                    start = new_start
                                    ' Exit the for loop
                                    Exit for
                                'If there is no non-zero starting price for that ticker for the year
                                'This is also the case when stock volume is zero for the year
                                'This is already handled above
                                End If
                            Next new_start
                        
                        ' If opening price is not zero then ...
                        End If
                    ' Calculate the yearly_change where yearly_change = closing price - opening price
                    yearly_change = ws.cells(i, 6) - ws.cells(start, 3)

                    ' Record yearly_change in summary table (column J)
                    ws.Range("J" & summary_row) = yearly_change

                    ' Calculate percent_change
                    ' Note percent_change = (yearly_change / opening price) * 100
                    percent_change = (yearly_change / ws.cells(start, 3)) * 100 ' note div by zero error

                    ' Record percent_change in summary table (column K)
                    ws.Range("K" & summary_row) = percent_change & "%"

                    End If

                    '-------------------------------------------------------------------
                    ' Advance summary_row to next row
                    summary_row = summary_row + 1

                    ' Reset yearly_change for next ticker symbol
                    yearly_change = 0

                    ' Advance start to the first row of the new ticker symbol
                    start = i + 1

                ' If ticker is the same ...
                Else
                    ' Add vol to total stock volume
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7)

                End if

            End if

        Next i 

    '---------------------------------------------------------------------------------
    ' Format Worksheet before going to next worksheet
    ws.columns.autofit


    Next ws

End Sub
