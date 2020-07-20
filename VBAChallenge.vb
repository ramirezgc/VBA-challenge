Sub VBAChallenge ()

Dim CurrentWs as Worksheet

'Define the worksheets that will be used in the for each loop - this will make the script work on all the worksheets in the workbook
For Each CurrentWs In Worksheets

'Define the variable types needed for the formulas
Dim Ticker as String
Dim OpenPrice as Double
OpenPrice = 0

Dim ClosePrice as Double
ClosePrice = 0

Dim YearlyChange as Double
YearlyChange =0

Dim PercentChange as Double
PercentChange =0

Dim TotalVolume as Double
TotalVolume = 0

'Define the starting row of the table that will house the summary of the ticker data
Dim Table As Integer
Table = 2

'Row count is needed per worksheet for the for loop to know how many times it will loop
Dim RowCount as Double
     
    'Row count is literaly a count of the rows in the current worksheet
    RowCount = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
    'Open price needs to be preserved for each ticker to be the open price on the first day of the year
    OpenPrice = CurrentWs.Cells(2,3).Value

    'This will define the headers in each worksheet at Row 1 in the corresponding columns
    CurrentWs.Cells(1, "J").Value = "Ticker"
    CurrentWs.Cells(1, "K").Value = "Yearly Change"
    CurrentWs.Cells(1, "L").Value = "Percent Change"
    CurrentWs.Cells(1, "M").Value = "Total Stock Volume"

    'The loop will start at 2 because the first row of data are just headers. This will loop for as many rows are in each worksheet
    For i = 2 to RowCount

        'When the ticker name in column A changes, the following actions will take place 
        If CurrentWs.Cells(i,"A") <> CurrentWs.Cells(i+1,"A").Value Then
            ' Ticker name is set to the value at row i column A
            Ticker = CurrentWs.Cells(i, 1).Value
            ' ClosePrice is set to the close price in row i column F 
            ClosePrice = CurrentWs.Cells(i,"F").Value
             'YearlyChange is the difference of the close price on the lasy day of the year and the open price on the first day of the year
            YearlyChange = ClosePrice - OpenPrice
             
             'This was required to ensure I was not getting an error for dividing by zero
                IF OpenPrice <> 0 Then
                    'Percent change is the YearlyChange divided by the OpenPrice and multiplied by 100 to get a percentage
                    PercentChange = (YearlyChange/OpenPrice)*100
                End If

            'Total volume will add the last volume amount for the ticker in the ith row to the volume that has been added in the else statement before the tickers no longer matched
            TotalVolume = TotalVolume + CurrentWs.Cells(i,"G").Value
            
            'The following calculations fill in the values above, to the summary table 
            CurrentWs.Range("J" & Table).Value = Ticker
            CurrentWs.Range("K" & Table).Value = YearlyChange
            'This if statment will conditionally format the yearly change to red or green
                If YearlyChange > 0 Then
                    CurrentWs.Range("K" & Table).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 THEN
                    CurrentWs.Range("K" & Table).Interior.ColorIndex = 3
                End IF

            CurrentWs.Range("L" & Table).Value = (Cstr(PercentChange)&"%")
            CurrentWs.Range("M" & Table).Value = TotalVolume

            'Moves the table row up by one and will reset the values for next ticker to calculate accurately
            Table = Table + 1
            YearlyChange = 0
            ClosePrice = 0
            OpenPrice = CurrentWs.Cells(i+1,3).Value
            TotalVolume = 0

            'While the ticker is the same in each row, the volume columns will be adding together
            Else
            
            TotalVolume = TotalVolume + CurrentWs.Cells(i,"G").Value
        'Ends the if loop
        End If
    'Moves on to the next i
    Next i
'Moves to the next worksheet 
Next CurrentWs

End Sub
