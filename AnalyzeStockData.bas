Sub AnalyzeStockData()

    'Loop through all worksheets
    For Each ws In Worksheets
    
    
        'Declare needed variables
        Dim yearlyChange As Double
        Dim closeAtLastDay As Double
        Dim openAtFirstDay As Double
        Dim yearlyPercentChange As Double
        Dim totalVolume As Double
        Dim lastrow As Double
        Dim ticker As String
        Dim summaryTableRow As Double
        Dim hundred As Double

  
        'Get the last row number
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Initialize the Summary Table row
        summaryTableRow = 2
    
        'Initialize Summary Variables
        totalVolume = 0
        openAtFirstDay = 0
        yearlyChange = 0
        yearlyPercentChange = 0
        hundred = 100
    
        'Print Headers for Summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'Format Column K as Percentage - didn't really work..
        'Range("K" & lastrow).NumberFormat = "0.00%"
    
        'Either assume the columns <ticker> and <date> are sorted in ascending order or write code to sort it
    
        'Loop through the cells of the column until the last row and for each new ticker
        For I = 2 To lastrow
    
            'If we find a new ticker, save what we counted so far and print it in the summary table
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
    
                'Since we are at the last day for a given ticker, get the
                closeAtLastDay = ws.Cells(I, 6).Value
    
                'Store the new current ticker value
                ticker = ws.Cells(I, 1).Value
    
                'Accumulate the Stock Volume
                totalVolume = totalVolume + ws.Cells(I, 7)
    
                'Calculate the Yearly Change
                yearlyChange = closeAtLastDay - openAtFirstDay
    
                'Calculate the Percent Change and catch null values
                If openAtFirstDay = 0 Then
                    yearlyPercentChange = 0
                Else
                    yearlyPercentChange = yearlyChange / openAtFirstDay
                End If
    
                'Print new ticker
                ws.Range("I" & summaryTableRow).Value = ticker
    
                'Print Yearly Change
                ws.Range("J" & summaryTableRow).Value = yearlyChange
    
                'Print Percent Change
                ws.Range("K" & summaryTableRow).Value = yearlyPercentChange
    
                'Print the total volume
                ws.Range("L" & summaryTableRow).Value = totalVolume
    
                'Increment the summary table row counter
                summaryTableRow = summaryTableRow + 1
    
                'Empty the buckets
                totalVolume = 0
                openAtFirstDay = 0
                closeAtLastDay = 0
                yearlyChange = 0
                yearlyPercentChange = 0
    
            'If we are still seeing the same ticker
            Else
    
                'The first day is the first time we set the open at first day
                If openAtFirstDay = 0 Then
    
                    openAtFirstDay = ws.Cells(I, 3).Value
    
                End If
    
                'Accumulate the total volume
                totalVolume = totalVolume + ws.Cells(I, 7).Value
    
            End If
    
        Next I
        
        'Once the Summary Table is populated, loop through its rows and use conditional formatting to highlight positive or negative change
        For j = 2 To summaryTableRow
        
            'If the value of Yearly Change is positive, color the cell in green
            If ws.Cells(j, 10) >= 0 Then
            
                ws.Cells(j, 10).Interior.ColorIndex = 4
            
           'Otherwise in red
            Else
            
                ws.Cells(j, 10).Interior.ColorIndex = 3
                
            End If
            
        Next j
        
        'Reinitialize the Summary Table row counter
        summaryTableRow = 0
        'MsgBox ws.Name

    Next ws
    
End Sub


