Sub StockAnalysis()
    
    'Add the instructions for the assignment up here.
    
    'Program assumptions
       '- The data always has headers in the first row

    'Sort columns in base data table to ensure stock ticker is filtered first, then date.
    Range("A:G").Sort Key1:=Range("A:A"), Order1:=xlAscending, Key2:=Range("B:B"), Order1:=xlAscending, Header:=xlYes
    
    'Declare variables
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

    'Ensure the appropriate data set in the current worksheet is active.
    Cells(1, 1).Activate

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

    'Set the Summary Table Headers
    Cells(k - 1, m).Value = "Ticker"
    Cells(k - 1, m + 1).Value = "Yearly Change"
    Cells(k - 1, m + 2).Value = "Percent Change"
    Cells(k - 1, m + 3).Value = "Total Stock Volume"

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
            Cells(k, m + 1).NumberFormat = "0.00"
            'Calculate the percent change for the year and record in the summary table
            Cells(k, m + 2).Value = (YearEndPrice - YearStartPrice) / YearStartPrice
            Cells(k, m + 2).NumberFormat = "0.00%"
            If Cells(k, m + 2).Value > 0 Then
                Cells(k, m + 2).Interior.ColorIndex = 4
            ElseIf Cells(k, m + 2).Value < 0 Then
                Cells(k, m + 2).Interior.ColorIndex = 3
            End If
            'Add the current row volume stock traded and record the subtotal into the summary table
            Cells(k, m + 3) = VolumeSubTotal + Cells(i, r).Value
            Cells(k, m + 3).NumberFormat = "#,##0"
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

    'This code is for the bonus section
    'Return the greatest % increase, greatest % decrease, and greatest total volume traded
    'Make the summary table the active table
    'Do a look to check the percent change value compared to the lowest and highest to that point as well as total volume.
    'If it is higher or lower, store the stock symbol and the value.

    'Declare variables
    Dim t As Long
    Dim GreatestVolume As Double

    'Set the summary values to zero
    Cells(2, 17).Value = 0
    Cells(3, 17).Value = 0
    Cells(4, 17).Value = 0

    'Set a cell in the summary table to active so that the analysis is done on the summary table data
    Cells(1, 10).Activate

    'Loop through all rows of data in the summary table to compare stock change and value to find the highest and lowest change and highest volume.
    For t = 2 To k

        If Cells(t, 12).Value > Cells(2, 17).Value Then
            Cells(2, 17).Value = Cells(t, 12).Value
            Cells(2, 16).Value = Cells(t, 10).Value
        ElseIf Cells(t, 12).Value < Cells(3, 17).Value Then
            Cells(3, 17).Value = Cells(t, 12).Value
            Cells(3, 16).Value = Cells(t, 10).Value
        End If

        If Cells(t, 13).Value > GreatestVolume Then
            GreatestVolume = Cells(t, 13).Value
            Cells(4, 17).Value = GreatestVolume
            Cells(4, 16).Value = Cells(t, 10).Value
        End If

    Next t

    'Add column headers and formatting to the bonus summary table
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Volume"
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 17).NumberFormat = "#,##0"

    'Adjust the width of all summary columns to fit all text and data to make visible for the reader.
    Columns("J:Q").AutoFit

End Sub
