Attribute VB_Name = "Module1"
Sub Wall_Street()

 Dim TestCounter As Integer
 Dim CurrentTickerValue As String
 Dim NextTickerValue As String
 Dim TickerRowCounter As Long
 Dim OpenValue As Currency
 Dim CloseValue As Currency
 Dim TotalStockVolume As Double
 Dim GreatestPercentIncrease As Double
 Dim GreatestPercentDecrease As Double
 
 TestCounter = 0
 GreatestPercentIncrease = 0
 GreatestPercentDecrease = 0
'Loop through all worksheets in the workbook
For Each ws In Worksheets

        ' Determine the Last Row
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

         'Write the header column labels for the output of summary values
          ws.Cells(1, 9).Value = "Ticker"
          ws.Cells(1, 10).Value = "Yearly Change"
          ws.Cells(1, 11).Value = "Percent Change"
          ws.Cells(1, 12).Value = "Total Stock Volume"
        ' Challenge assignment
        
        'MsgBox "ws:" & ws.Index
        If ws.Index = 1 Then
          ws.Cells(1, 16).Value = "Ticker"
          ws.Cells(1, 17).Value = "Value"
          ws.Cells(2, 15).Value = "Greatest % Increase"
          ws.Cells(3, 15).Value = "Greatest % Decrease"
          ws.Cells(4, 15).Value = "Greatest Total Volume"
        End If
        
         'Loop through the <ticker> column on each worksheet, starting at row 2, since row 1 is the header row
         CurrentTickerValue = ws.Cells(2, 1).Value
         ws.Cells(2, 9).Value = CurrentTickerValue
         TickerRowCounter = 2

         OpenValue = ws.Cells(2, 3).Value
         'TotalStockVolume = ws.Cells(2, 7)

        For i = 2 To Lastrow
           'Find each unique ticker value by detecting when it changes value
           'Fetch the next ticker value from the next cell on the worksheet for comparison
           CurrentTickerValue = ws.Cells(i, 1).Value
           NextTickerValue = ws.Cells(i + 1, 1).Value
           
           'Compare the next cell ticker value with whatever was held in the current (previous) ticker value variable
           'If the value has not changed then lets continue adding to the total stock volumne
           If CurrentTickerValue = NextTickerValue Then
              
              TotalStockVolume = TotalStockVolume + ws.Cells(i + 1, 7)
             ' ws.Cells(TickerRowCounter, 12).Value = TotalStockVolume
           
           'If the value has changed indicating a new ticker value has been detected then
           ElseIf NextTickerValue <> CurrentTickerValue Then
              
             'MsgBox "Next Ticker Value: " & NextTickerValue
             
             'if you've reached the next ticker value then save the totalstockvolume to the worksheet and reset the
             'TotalStockVolume variable to 0 for the next ticker
              ws.Cells(TickerRowCounter, 12).Value = TotalStockVolume
              
              If GreatestTotalVolume < TotalStockVolume Then
                   GreatestTotalVolume = TotalStockVolume
                   GreatestTotalVolumeTicker = CurrentTickerValue
              End If
              
              TotalStockVolume = 0
             
             'if you've reached the next ticker value then capture the year end closing value of the previous ticker
             CloseValue = ws.Cells(i, 6).Value
             
             'Calculate the annual change in value and assign it to a variable and to the cell for "Yearly Change"
             GetYearlyChange = CloseValue - OpenValue
             ws.Cells(TickerRowCounter, 10).Value = GetYearlyChange
             
             'Check for positive change or negative change, and color the cell accordingly as (Green) or (Red)
             
             If GetYearlyChange >= 0 Then
                ' Set the Cell Colors to Green : positive change
                ws.Cells(TickerRowCounter, 10).Interior.ColorIndex = 4
             Else
                ' Set the Cell Colors to Red : Negative change
                 ws.Cells(TickerRowCounter, 10).Interior.ColorIndex = 3
             End If
             
             'Calculate the percentage change in value and assign it to the cell adjacent to the ticker value, and Yearly change in value
             'Check for divide by zero error by ensuring that OpenValue is not = 0, if it is 0 calculate % change to account for a 0 open value
             'Format cell as a percentage with 2 decimal places
             
             ws.Cells(TickerRowCounter, 11).NumberFormat = "0.00%"
             
             If OpenValue <> 0 Then
                
                ws.Cells(TickerRowCounter, 11).Value = (((CloseValue - OpenValue) / OpenValue))
                
                'Capture values for greatest percent increase and greatest percent decrease as we loop though all sheets
                'and save to variables for later reference
                If GreatestPercentIncrease < (((CloseValue - OpenValue) / OpenValue)) Then
                   GreatestPercentIncrease = (((CloseValue - OpenValue) / OpenValue))
                   GreatestPercentIncreaseTicker = CurrentTickerValue
                End If
                
                If GreatestPercentDecrease > (((CloseValue - OpenValue) / OpenValue)) Then
                   GreatestPercentDecrease = (((CloseValue - OpenValue) / OpenValue))
                   GreatestPercentDecreaseTicker = CurrentTickerValue
                End If
                   
             Else
                ws.Cells(TickerRowCounter, 11).Value = 0
             End If
             
             'Set CurrentTickerValue to the new NextTickerValue
             CurrentTickerValue = NextTickerValue
             
             'Increment the row counter for the next entry of the next ticker
             TickerRowCounter = TickerRowCounter + 1
             
             'Populate the next ticker value detected in the next row of the worksheet
             ws.Cells(TickerRowCounter, 9).Value = CurrentTickerValue
             
             'Capture the openvalue of the next ticker
             OpenValue = ws.Cells(i + 1, 3).Value
             
           End If
        Next i
     ws.Columns("I:R").AutoFit

 Next ws
      Sheet1.Cells(2, 16).Value = GreatestPercentIncreaseTicker
      Sheet1.Cells(2, 17).NumberFormat = "0.00%"
      Sheet1.Cells(2, 17).Value = GreatestPercentIncrease
      
      Sheet1.Cells(3, 16).Value = GreatestPercentDecreaseTicker
      Sheet1.Cells(3, 17).NumberFormat = "0.00%"
      Sheet1.Cells(3, 17).Value = GreatestPercentDecrease
      
      Sheet1.Cells(4, 16).Value = GreatestTotalVolumeTicker
      Sheet1.Cells(4, 17).Value = GreatestTotalVolume
      
      Sheet1.Columns("O:Q").AutoFit
End Sub


