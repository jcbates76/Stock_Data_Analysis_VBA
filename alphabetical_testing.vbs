Sub StockAnalysis()
    
    'Add the instructions for the assignment up here.
    
    'Program assumptions
       '- The data always has headers in the first row

    Dim ws As Worksheet
    
    'Declare variables for performing analysis on the raw data tables
    Dim RowsOfData As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim n As Long
    Dim p As Long
    Dim q As Long
    Dim YearStartPrice As Currency
    Dim YearEndPrice As Currency
    Dim VolumeSubTotal As Double
    
    'Declare variables for performing analysis on the summary table
    Dim t As Long
    Dim GreatestVolume As Double
    
    
    '----------------------------------------------------
    'Loop through all of the worksheets in the workbook.
    'Each worksheet is a year of data
    '----------------------------------------------------
    For Each ws In Worksheets
    
    
        'Sort columns in base data table to ensure stock ticker is filtered first, then date.
        ws.Range("A:G").Sort Key1:=ws.Range("A:A"), Order1:=xlAscending, Key2:=ws.Range("B:B"), Order1:=xlAscending, Header:=xlYes
        
        
        'Ensure the appropriate data set in the current worksheet is active.
        Range("A1:A1").Select
    
        'Determine the number of rows of data in the table
        RowsOfData = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Set variables here instead of hard-coding
        'Column where stock ticker is stored
        j = 1
        'Column where stock open price is stored
        q = 3
        'Column where stock close price is stored
        n = 6
        'Column where stock volume traded for the day is stored
        r = 7
        'Row where raw data starts
        p = 2
        'Row where summarized data will start
        k = 2
        'Column where the summarized data for each ticker symbol will start
        m = 10
    
        'Set the Summary Table Headers
        ws.Cells(k - 1, m).Value = "Ticker"
        ws.Cells(k - 1, m + 1).Value = "Yearly Change"
        ws.Cells(k - 1, m + 2).Value = "Percent Change"
        ws.Cells(k - 1, m + 3).Value = "Total Stock Volume"
    
        'Set the YearStartPrice for the first stock symbol and reset the Year End Price and Total Stock Volume
        YearStartPrice = ws.Cells(p, q).Value
        YearEndPrice = 0
        TotalStockVolume = 0
    
        'Check all of the rows to determine when the stock symbol changes
        For i = p To RowsOfData
            
            'Check to see if the stock symbol on the row below is the same as the current row.  If not, this would be the last row in that subset.
            If ws.Cells(i + 1, j).Value <> ws.Cells(i, j).Value Then
                
                'Record the stock symbol at the end of that subset
                ws.Cells(k, m).Value = ws.Cells(i, j).Value
                'Record the closing price of that stock
                YearEndPrice = ws.Cells(i, n).Value
                'Determine the change in stock price for the year and record in the summary table
                ws.Cells(k, m + 1).Value = YearEndPrice - YearStartPrice
                ws.Cells(k, m + 1).NumberFormat = "0.00"
                
                'Calculate the percent change for the year and record in the summary table
                'If statement to avoid division by zero
                If YearStartPrice <> 0 Then
                    ws.Cells(k, m + 2).Value = (YearEndPrice - YearStartPrice) / YearStartPrice
                Else
                    ws.Cells(k, m + 2).Value = 0
                End If
                
                ws.Cells(k, m + 2).NumberFormat = "0.00%"
                
                'Add conditional formatting for percent change.  Red for loss for the year, green for gain for the year.
                If ws.Cells(k, m + 2).Value > 0 Then
                    ws.Cells(k, m + 2).Interior.ColorIndex = 4
                ElseIf ws.Cells(k, m + 2).Value < 0 Then
                    ws.Cells(k, m + 2).Interior.ColorIndex = 3
                End If
                
                'Add the current row volume stock traded and record the subtotal into the summary table
                ws.Cells(k, m + 3) = VolumeSubTotal + ws.Cells(i, r).Value
                ws.Cells(k, m + 3).NumberFormat = "#,##0"
                
                'Reset the variables for the next stock symbol
                YearStartPrice = ws.Cells(i + 1, q).Value
                YearEndPrice = 0
                VolumeSubTotal = 0
                
                'Move to the next row in the summary table
                k = k + 1
            
            Else
                
                'The stock symnbotl on the row below is the same as the current row, and continue to add the daily volume of stock traded for that symbol.
                VolumeSubTotal = VolumeSubTotal + ws.Cells(i, r).Value
    
            End If
        
        Next i
    
        'This code is for the bonus section
        'Return the greatest % increase, greatest % decrease, and greatest total volume traded
        'Make the summary table the active table
        'Do a look to check the percent change value compared to the lowest and highest to that point as well as total volume.
        'If it is higher or lower, store the stock symbol and the value.
    
        'Set the GreatestVolume variable to 0
        GreatestVolume = 0
    
        'Set the summary values to zero
        ws.Cells(2, 17).Value = 0
        ws.Cells(3, 17).Value = 0
        ws.Cells(4, 17).Value = 0
    
        
        
        
        'Set a cell in the summary table to active so that the analysis is done on the summary table data
        Cells(1, 10).Select
    
        'Loop through all rows of data in the summary table to compare stock change and value to find the highest and lowest change and highest volume.
        For t = 2 To k
    
            If ws.Cells(t, 12).Value > ws.Cells(2, 17).Value Then
                ws.Cells(2, 17).Value = ws.Cells(t, 12).Value
                ws.Cells(2, 16).Value = ws.Cells(t, 10).Value
            ElseIf ws.Cells(t, 12).Value < ws.Cells(3, 17).Value Then
                ws.Cells(3, 17).Value = ws.Cells(t, 12).Value
                ws.Cells(3, 16).Value = ws.Cells(t, 10).Value
            End If
    
            If ws.Cells(t, 13).Value > GreatestVolume Then
                GreatestVolume = ws.Cells(t, 13).Value
                ws.Cells(4, 17).Value = GreatestVolume
                ws.Cells(4, 16).Value = ws.Cells(t, 10).Value
            End If
    
        Next t
    
        'Add column headers and formatting to the bonus summary table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).NumberFormat = "#,##0"
    
        'Adjust the width of all summary columns to fit all text and data to make visible for the reader.
        ws.Columns("J:Q").AutoFit
        
        MsgBox ("Macro Complete for " & ws.Name)
    Next ws
    
  

End Sub