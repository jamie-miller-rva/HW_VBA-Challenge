Sub Reset()
   ' Reset all worksheets (ws) in worksheets
   ' Declare Variables
   Dim ws As Worksheet
   Dim LastRow As Long ' Long is a large or "long" integer or counting number
   Dim row As Long ' I will use this to refer to rows in my worksheet
   
   ' -----------------------------------------------------------------
   ' Integrate through all worksheets
   ' -----------------------------------------------------------------
   
   For Each ws In Worksheets
   
   ' Initalize variables for each worksheet
   LastRow = 0 ' LastRow will find the last row in the worksheet (want it to be zero for the start of each ws)
   
   ' Identify the last row
   LastRow = ws.Cells(Rows.Count, "A").End(xlUp).row ' this code is from an activity in Day 3 called Wells Fargo
   ' test message to identify the macro is moving to the next worksheet
   'MsgBox (LastRow)
   
   'Reset Headers in summary table
   ws.Range("I1:I" & LastRow) = ""
   ws.Range("J1:J" & LastRow) = ""
   ws.Range("K1:K" & LastRow) = ""
   ws.Range("L1:L" & LastRow) = ""
   
   ws.Range("J1:J" & LastRow).Interior.ColorIndex = 0
   
      
   Next ws
   

End Sub