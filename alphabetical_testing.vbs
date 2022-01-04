Sub StockAnalysis()

    'Program assumptions
        '- The stock name is sorted alphabetically
        '- The date is sorted earliest to latest
        '- The data always has headers in the first row
    
    'Declare variables
    Dim RowsOfData As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim n As Long
    Dim p As Long
    Dim q As Long
    Dim YearStartPrice As Long
    Dim YearEndPrice As Long
    Dim VolumeSubTotal As Double
'    Dim Stock_Name As String
'    Dim Stock_Date As Long
'    Dim Yearly_Open As Long
'    Dim Yearly_Close As Long
'    Dim Volume As Double
'    Dim Yearly_Change As Long
'    Dim Yearly_Percent_Change As Long
    
    'Determine the number of rows of data in the table
    RowsOfData = Cells(Rows.Count, 1).End(xlUp).Row
        
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
    'Column where the summarized stock symbol will be stored
    m = 10
    
    'Set the YearStartPrice for the first stock symbol
    YearStartPrice = Cells(p, q).Value
    
    'Check all of the rows to determine when the stock symbol changes
    For i = p To RowsOfData
        'Check to see if the stock symbol on the row below is the same as the current row.  If not, this would be the last row in that subset.
        If Cells(i + 1, j).Value <> Cells(i, j).Value Then
            'Record the stock symbol at the end of that subset
            Cells(k, m).Value = Cells(i, j).Value
            'Record the closing price of that stock
            YearEndPrice = Cells(i, n).Value
            'Determine the change in stock price for the year and record in the summary table
            Cells(k, m + 1).Value = YearEndPrice - YearStartPrice
            'Calculate the percent change for the year and record in the summary table
            Cells(k, m + 2).Value = (YearEndPrice - YearStartPrice) / YearStartPrice
            'Add the current row volume stock traded and record the subtotal into the summary table
            Cells(k, m + 3) = VolumeSubTotal + Cells(i, r).Value
            'Reset the variables for the next stock symbol
            YearStartPrice = Cells(i + 1, q).Value
            VolumeSubTotal = 0
            'Move to the next row in the summary table
            k = k + 1
        Else
            'The stock symnbotl on the row below is the same as the current row, and continue to add the daily volume of stock traded for that symbol.
            VolumeSubTotal = VolumeSubTotal + Cells(i, r).Value
        
        End If
    Next i
            
            
            
            
            
    
    
    
    
    
    
    
    
    'Stock_Name = Cells(2, 1).Value
                
    'Set j as the column where data entry will happen for the summary
    'j = 1
    
    'Set Volume to 0 initially
    'Volume = 0
    
    'Set Stock_Date to initial data cell
    'Stock_Date = Cells(2, 2).Value
                   
    'Loop through the stock names to see if the stock row has changed.
    'Accummulate the volume for that stock name throughout the course of the year.
    'For i = 2 To Rows
        'If Cells(i, 1) <> Stock_Name Then
            'Cells(j + 1, 10).Value = Stock_Name
            'Cells(j + 1, 13).Value = Volume
            'Cells(j + 1, 11).Value = Yearly_Close - Yearly_Open
            'Cells(j + 1, 12).Value = (Yearly_Close - Yearly_Open) / Yearly_Open
            
            ' Reset variables
            'Stock_Name = Cells(i, 1)
            'j = j + 1
            'Volume = Cells(i, 7).Value
            'Yearly_Open = Cells(i, 3).Value
            'Yearly_Close = Cells(i, 6).Value
        'Else
            'Volume = Volume + Cells(i, 7).Value
            'If Stock_Date < Cells(i, 2).Value Then
                'Yearly_Open = Cells(i, 3).Value
            'ElseIf Stock_Date > Cells(i, 2).Value Then
                'Yearly_Close = Cells(i, 6).Value
            'End If
        'End If
    'Next i
    
    'Format(Range("K:K"),"Standard")
    
    
    'For data analysis only
    'Cells(1, 10).Value = Rows
    
End Sub
