Sub test_for_year_stock_data():


'setting all the dimensions over here and declaring the variables

Dim ws As Worksheet
Dim total As Double
Dim i As Long
Dim YearlyChange As Double
Dim j As Integer
Dim start As Long
Dim LastRow As Long
Dim PercentageChange As Double
Dim days As Integer

Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestTotalTicker As String
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotal As Double

   
   
   'this code provides the function of looping through all the worksheets
   
   For Each ws In Worksheets
      ws.Activate
  
   
   'adding the column names for the requested data
   Range("I1").Value = "Ticker"
   Range("J1").Value = "Yearly Change"
   Range("K1").Value = "Percentage Change"
   Range("L1").Value = "Total Stock Volume"
   Range("P1").Value = "Ticker"
   Range("Q1").Value = "Value"
   Range("O2").Value = "Greatest % Increase"
   Range("O3").Value = "Greatest % Decrease"
   Range("O4").Value = "Greatest Total Volume"
   
   
   'overhere we are stating the intial values
   j = 0
   total = 0
   YearlyChange = 0
   start = 2
   
   
   'determining the last row to chek through all the rows
   LastRow = Cells(Rows.Count, "A").End(xlUp).Row
   For i = 2 To LastRow
   
   
   'see if we are still within the same ticker sign
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           total = total + Cells(i, 7).Value
           If total = 0 Then
               
               
               'put the resaults gathered in the field
               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = 0
               Range("K" & 2 + j).Value = "%" & 0
               Range("L" & 2 + j).Value = 0
            Else
           
           
           
           'finding ticker when value is not 0
              If Cells(start, 3) = 0 Then
               For find_value = start To i
                       If Cells(find_value, 3).Value <> 0 Then
                           start = find_value
                           Exit For
                       End If
                Next find_value
              End If
               
            
            'calculating change by dividing yearly change by the opening price and then storing the output in percentage coloumn
            YearlyChange = (Cells(i, 6) - Cells(start, 3))
            PercentageChange = YearlyChange / Cells(start, 3)
            
            'start of the next stock ticker and record results
            
            start = i + 1
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = YearlyChange
            Range("J" & 2 + j).NumberFormat = "0.00"
            Range("K" & 2 + j).Value = PercentageChange
            Range("K" & 2 + j).NumberFormat = "0.00%"
            Range("L" & 2 + j).Value = total
            
                'add color conditional formatting to yearly change and percentage change
                
                If (YearlyChange > 0) Then
                Range("J" & 2 + j).Interior.ColorIndex = 4
                ElseIf (YearlyChange <= 0) Then
                Range("J" & 2 + j).Interior.ColorIndex = 3
                End If
                If (PercentageChange > 0) Then
                Range("K" & 2 + j).Interior.ColorIndex = 4
                ElseIf (PercentageChange <= 0) Then
                Range("K" & 2 + j).Interior.ColorIndex = 3
                End If
               
               
            End If
                         
           'resets the value of the current ticker if new ticker is found
           total = 0
           YearlyChange = 0
           j = j + 1
           days = 0
           
           'if same ticker then add the value to the ticker box
       Else
           total = total + Cells(i, 7).Value
    End If
    
          
    Next i
    
      'Find the last row of data in column I
      LastRow = Cells(Rows.Count, "I").End(xlUp).Row

      'find the ticker symbol with the greatest percentage increase, percentage decrease, and total volume
      GreatestIncrease = Application.WorksheetFunction.Max(Range("K2:K" & LastRow))
      GreatestDecrease = Application.WorksheetFunction.Min(Range("K2:K" & LastRow))
      GreatestTotal = Application.WorksheetFunction.Max(Range("L2:L" & LastRow))
      GreatestIncreaseTicker = Cells(Application.WorksheetFunction.Match(GreatestIncrease, Range("K2:K" & LastRow), 0) + 1, 9).Value
      GreatestDecreaseTicker = Cells(Application.WorksheetFunction.Match(GreatestDecrease, Range("K2:K" & LastRow), 0) + 1, 9).Value
      GreatestTotalTicker = Cells(Application.WorksheetFunction.Match(GreatestTotal, Range("L2:L" & LastRow), 0) + 1, 9).Value

      'write the results to the summary table
      Range("P2:P4").Value = Array(GreatestIncreaseTicker, GreatestDecreaseTicker, GreatestTotalTicker)
      Range("Q2").Value = GreatestIncrease
      Range("Q3").Value = GreatestDecrease
      Range("Q2:Q3").NumberFormat = "0.00%"
      Range("Q4").Value = GreatestTotal
      Range("Q4").NumberFormat = "#"

   
    Next ws

End Sub