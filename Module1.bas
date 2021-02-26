Attribute VB_Name = "Module1"
Sub StockSummary()

 
    Dim lastrow As Long
    Dim lastcolumn As Long
    Dim row As Long
    Dim total As Double
    Dim summaryRow As Long
    Dim ws As Worksheet
    
    For Each ws In Sheets 'loop through each worksheet
        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yealy Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(1, 19) = "open"
        ws.Cells(1, 20) = "close"
        ws.Cells(2, 15) = "Great % Increase"
        ws.Cells(3, 15) = "Great % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim OpeningPriceForFirstTicker As Double
        Dim PercentChange As Double
        
        
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row 'Find the last row of the sheet
        lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column 'Find the last column of the sheet
        
        summaryRow = 2
        total = 0
        
        For row = 2 To lastrow 'Start to look at row 2,3,4...till lastrow
        
            OpeningPriceForFirstTicker = ws.Cells(2, 3).Value 'OpeningPriceForFirstTicker is not captured by If loop below so need to add it here as a special
            ws.Cells(2, 19).Value = OpeningPriceForFirstTicker
            total = ws.Cells(row, 7).Value + total
            
            
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then 'If value in row below is different than current row, the values in the column cotinues to be the same until reaches this breakpoint
               
               ws.Cells(summaryRow, 9).Value = ws.Cells(row, 1).Value 'Print the last row value of the unique value to Cells(2, 9)
               ws.Cells(summaryRow, 12).Value = total                 'Print the total calculated above to Cells(2, 12)
               
               ClosingPrice = ws.Cells(row, 6).Value    'Grab the the last row value of the unique value and then print to Cells(2, 20)
               ws.Cells(summaryRow, 20).Value = ClosingPrice
        
               ws.Cells(summaryRow, 10).Value = ws.Cells(summaryRow, 20).Value - ws.Cells(summaryRow, 19).Value 'calculate the Yearly change which has been printed to summaryRow
               
               If ws.Cells(summaryRow, 19).Value <> 0 Then
                  PercentChange = Round(ws.Cells(summaryRow, 10).Value / ws.Cells(summaryRow, 19).Value, 4) 'calculate the precentchange based on values which have been printed to summaryRow
               Else
                  PercentChange = 0
               End If
               
               ws.Cells(summaryRow, 11).Value = PercentChange
               ws.Cells(summaryRow, 11).Style = "Percent"
               ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
               
        
               summaryRow = summaryRow + 1                      'Finished current summary row and come down to summarize next ticker type
               total = 0                                        'Reset the total value to 0 for next new ticker type
               
               OpeningPrice = ws.Cells(row + 1, 3).Value        'Capture the first value of opening price for next new ticker type
               ws.Cells(summaryRow, 19).Value = OpeningPrice
               
            End If
        Next row
        
        
        
     If IsEmpty(ws.Cells(summaryRow, 9).Value) Then
        ws.Cells(summaryRow, 10).ClearContents
        ws.Cells(summaryRow, 11).ClearContents
        ws.Cells(summaryRow, 19).ClearContents
        ws.Cells(summaryRow, 20).ClearContents
     End If
        
    
     lastsummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).row
     
        'Find the greatest total volume and its ticker
        For summaryRow = 2 To lastsummaryRow
            If ws.Cells(summaryRow, 12).Value > ws.Cells(4, 17).Value Then
               ws.Cells(4, 16).Value = ws.Cells(summaryRow, 9).Value
               ws.Cells(4, 17).Value = ws.Cells(summaryRow, 12).Value
            End If
        Next summaryRow
        
        'Find the greatest %increase and its ticker
        For summaryRow = 2 To lastsummaryRow
            If ws.Cells(summaryRow, 11).Value > ws.Cells(2, 17).Value Then
               ws.Cells(2, 16).Value = ws.Cells(summaryRow, 9).Value
               ws.Cells(2, 17).Value = ws.Cells(summaryRow, 11).Value
               ws.Cells(2, 17).Style = "Percent"
               ws.Cells(2, 17).NumberFormat = "0.00%"
            End If
        Next summaryRow
        
        
       'Find the greatest %decrease and its ticker
        For summaryRow = 2 To lastsummaryRow
            If ws.Cells(summaryRow, 11).Value < ws.Cells(3, 17).Value Then
               ws.Cells(3, 16).Value = ws.Cells(summaryRow, 9).Value
               ws.Cells(3, 17).Value = ws.Cells(summaryRow, 11).Value
               ws.Cells(3, 17).Style = "Percent"
               ws.Cells(3, 17).NumberFormat = "0.00%"
            End If
        Next summaryRow
        
        'Highlight cell colors based on yearly change values
        For summaryRow = 2 To lastsummaryRow
            If ws.Cells(summaryRow, 10).Value < 0 Then
               ws.Cells(summaryRow, 10).Interior.Color = vbRed
            ElseIf ws.Cells(summaryRow, 10).Value > 0 Then
                   ws.Cells(summaryRow, 10).Interior.Color = vbGreen
            End If
        Next summaryRow
        
        
     Next ws
End Sub
