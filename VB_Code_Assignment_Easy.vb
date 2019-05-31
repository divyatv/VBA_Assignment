' Homework easy - get the ticket symbol and total stock volume for all the etf
' Sample code structure or logic
'go over all the sheets
'go over all the rows in each sheets using last row
'set a ticker symbol - to begin with the first row
'set the total volume - to start with zero
'check the cell with next cell and if they are not equal - use the <> sign
' if they are same stock then add the total
'If they are not same stock then reset total to fist row of new ticker, reset summry table row to next one, reset ticker to next stock.
' Do the steps for all the worksheets.

Sub sumtotal()

'Initialize a variable to get the count of active worksheets
'Dim ws_count As Integer

'initialize a counter to go over all the worksheets.
'Dim ws As Integer

' Get the count of all the worksheets in the excel file.
'ws_count = ActiveWorkbook.Worksheets.Count
'MsgBox ("active sheet in this file=" & ws_count)

Dim wk As Worksheet

'Loop through each work sheet
For Each wk In ThisWorkbook.Worksheets
wk.Activate
          

'loop through all the work sheets and do the same calculations.
'For ws = 1 To ws_count -- this does it three times in the same sheet.

    ' Initialise a variable for  totalvolume
    Dim totals As Double
    
    ' Initializing total to the first stock volume
    totals = Cells(2, 7).Value
    'MsgBox ("totals=" & totals)
    
    Dim ticker As String
    'start ticket with the first symbol
    ticker = Cells(2, 1).Value
    
    ' set and get the last row count to loop
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'set an integer for summary table row
    Dim summarytablerow As Integer
    summarytablerow = 2

            'for loop to lopp through all the rows until empty
                For i = 2 To lastrow
            
                                    
                         ' Check if the next row is same stock as the current one -if its not equal to stop adding the total stock volume.
                                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                                'Write the ticker and total volume in summary table
                                Cells(summarytablerow, 10).Value = ticker
                                Cells(summarytablerow, 11).Value = totals
                                
                                'reset the total volume to the first total volume of new stock.
                                totals = Cells(i + 1, 7).Value
                                'Set the summary table to write to next row making sure summary table doesnt get overwritten.
                                 summarytablerow = summarytablerow + 1
                                ticker = Cells(i + 1, 1).Value
                         
                         ' Dont miss the last row - to add the stock volume of last row.
                                ElseIf i = lastrow Then
                                   totals = totals + Cells(i, 7).Value
                                   Cells(summarytablerow, 10).Value = ticker
                                   Cells(summarytablerow, 11).Value = totals
                          
                          'Add the total volume if the nth row and n+1th row are the same stock.
                          
                                Else
                                   totals = totals + Cells(i + 1, 7).Value
                                End If
                        
                          Next i

    'Next ws
     ' Insert your code here.
            ' This line displays the worksheet name in a message box.
            MsgBox wk.Name
  Next wk

End Sub

