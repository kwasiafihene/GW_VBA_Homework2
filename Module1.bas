Attribute VB_Name = "Module1"
Sub Stock_Data()

    ' Loop through all of the worksheets in the active workbook.
    For Each ws In Worksheets


         ' Set an initial variable for holding the Ticker
         Dim Ticker As String
        
         ' Set an initial variable for holding the total volume per Ticker
         Dim Stock_Volume As Double
         Stock_Volume = 0
        
         ' Keep track of the location for each Ticker in the summary table
         Dim Summary_Table_Row As Integer
         Summary_Table_Row = 2
         
         'Count number of rows
         'Dim LastRow As Integer
          LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
         ' Loop through all Stock transactions
         For i = 2 To LastRow
        
           ' Check if we are still within the same Ticker, if it is not...
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
             ' Set the Ticker
             Ticker = ws.Cells(i, 1).Value
        
             ' Add to the Stock Volume
             Sock_Volume = Stock_Volume + ws.Cells(i, 7).Value
             ' Print column headings for summary table
             ws.Range("I1").Value = "Ticker"
             ws.Range("J1").Value = "Total Stock Volume"
             ' Print the Ticker in the Summary Table
             ws.Range("I" & Summary_Table_Row).Value = Ticker
        
             ' Print the Ticker Amount to the Summary Table
             ws.Range("J" & Summary_Table_Row).Value = Stock_Volume
             ' Add one to the summary table row
             Summary_Table_Row = Summary_Table_Row + 1
        
             ' Reset the Stock Volume
             Stock_Volume = 0
        
           ' If the cell immediately following a row is the same brand...
           Else
        
             ' Add to the Stock Volume
             Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        
           End If
        
         Next i
         
    Next ws
    
End Sub


