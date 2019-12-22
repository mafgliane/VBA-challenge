Attribute VB_Name = "Module1"
Sub Stock_Report():

Dim lastRow, currentRow, outputRow As Long
Dim stockTotal, maxVolume, openValue, closeValue As Double
Dim yDelta, maxIncr, minIncr As Double
Dim maxIncTicker, minIncTicker, maxVolTicker As String

   
    For Each ws In Worksheets
    
        'Create column headers
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        outputRow = 2
        currrow = 2
        maxIncr = 0
        minIncr = 0
        maxVolume = 0
   
       
        
        'Get needed data per ticker tape name
        Do While currrow <= lastRow
            
            If openValue = 0 Then
                openValue = ws.Cells(currrow, 3).Value
            End If
            
            stockTotal = stockTotal + ws.Cells(currrow, 7).Value

            If ws.Cells(currrow, 1).Value <> ws.Cells(currrow + 1, 1).Value Or currrow > lastRow Then
            
                'Closing Price
                closeValue = ws.Cells(currrow, 6).Value
            
                'Ticker
                ws.Cells(outputRow, 9).Value = ws.Cells(currrow, 1).Value
            
                'Yearly Change
                ws.Cells(outputRow, 10).Value = closeValue - openValue
                If (closeValue - openValue) < 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                End If
            
                'Percent Change
                If openValue = 0 Or closeValue = 0 Then
                    yDelta = 0
                Else
                    yDelta = (closeValue / openValue) - 1
                End If
                
                ws.Cells(outputRow, 11).Value = yDelta
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                

                If maxIncr < yDelta Then
                    maxIncr = yDelta
                    maxIncTicker = ws.Cells(currrow, 1).Value
                End If
                
                If yDelta < minIncr Then
                    minIncr = yDelta
                    minIncTicker = ws.Cells(currrow, 1).Value
                End If

                           
                'Total Stock Volume
                ws.Cells(outputRow, 12).Value = stockTotal
                If maxVolume < stockTotal Then
                    maxVolume = stockTotal
                    maxVolTicker = ws.Cells(currrow, 1).Value
                End If
                               
                            
                'Reset varialbles for next ticker
                stockTotal = 0
                openValue = 0
                closeValue = 0
                outputRow = outputRow + 1
            End If
            
            'Advance to next row
            currrow = currrow + 1
        Loop
        
    'Write greatest values
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(2, 17).Value = maxIncTicker
    ws.Cells(2, 18).Value = maxIncr
    ws.Cells(2, 18).NumberFormat = "0.00%"
    
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(3, 17).Value = minIncTicker
    ws.Cells(3, 18).Value = minIncr
    ws.Cells(3, 18).NumberFormat = "0.00%"
    
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    ws.Cells(4, 17).Value = maxVolTicker
    ws.Cells(4, 18).Value = maxVolume
    
    ws.Columns("I:R").AutoFit
    
    Next ws
End Sub
