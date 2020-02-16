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

' ------------------------------------------------------------------------------------
    ' Interate through each worksheet (ws)
    For Each ws in worksheets

 

    ' Initalize variables for each worksheet
    LastRow = 0 ' LastRow will find the last row in the worksheet (want it to be zero for the start of each ws)

    summary_row = 2

    total_stock_volume = 0
   
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

            ' If the ticker symbol changes then record findings in summary table
            ' Note column 1 is the location of the ticker symbol
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) then

                ' Record ticker symbol in summary table (column I)
                ' Note: I need a counter to keep track of the summary table row.
                ' I will use summary_row starting on row 2
                ws.Range("I" & summary_row) = ws.Cells(i, 1)

                ' Record total stock volumne in summary table (column L)
                ws.Range("L" & summary_row) = total_stock_volume


                ' Advance summary_row to next row
                summary_row = summary_row + 1

            ' If ticker is the same ...
            Else
                ' Add vol to total stock volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7)

            End if

        Next i 

    '---------------------------------------------------------------------------------
    ' Format Worksheet before going to next worksheet
    ws.columns.autofit


    Next ws

End Sub
