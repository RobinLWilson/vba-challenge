Sub tickerfinal()
'set variables
Dim open_price As Variant
Dim close_price As Variant
Dim total_volume As String
Dim yearly_change As Variant

'Go through worksheets
    For Each ws In Worksheets
        ws.Activate
            
        'Create Row titles
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        'Set total volume & count of each ticker symbol
        total_volume = 0
        summary_pointer = 2
        
        'count how many rows are in each sheet
        row_count = Cells(Rows.Count, "A").End(xlUp).Row
        
        'set initial open price
        open_price = Cells(2, 3).Value
                   
        'Create Greatest Summary Table Headings
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
                  
        'Begin looping to find groups of ticker symbols
        For i = 2 To row_count
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                close_price = Cells(i, 6).Value
                
                'enter ticker symbol in summary table
                Range("I" & summary_pointer).Value = Cells(i, 1).Value
                
                'enter Yearly change in column j & format cell
                yearly_change = close_price - open_price
                Range("J" & summary_pointer).Value = yearly_change
                Range("J" & summary_pointer).NumberFormat = "$#,##0.00"
                
                'enter percent change in column k & format cell
                Range("K" & summary_pointer).Value = (yearly_change / open_price)
                Range("K" & summary_pointer).NumberFormat = "0.00%"
                
                'enter total stock volume
                Range("L" & summary_pointer).Value = total_volume
                
                summary_pointer = summary_pointer + 1
                
                open_price = Cells(i + 1, 3).Value
                total_volume = 0
            Else
                total_volume = total_volume + Cells(i, 7).Value
                
            End If
        Next i
        
'analyze and format the summary table

    'find the number of columns in the summary table
    Dim summary_count As Variant
    summary_count = Cells(Rows.Count, "K").End(xlUp).Row
    
'format conditional cells from VBA 2.3 Grader exercise
        'set counter for conditional formatting
        j = 0
        
        'loop through summary table
        For j = 2 To summary_count
            
            'if statement for conditional formatting in column J
            If Range("J" & j).Value >= 0 Then
                Range("J" & j).Interior.ColorIndex = 4
            
            Else
                Range("J" & j).Interior.ColorIndex = 3
            
            End If
            
            'if statement for conditional formatting in column k
             If Range("K" & j).Value >= 0 Then
                Range("K" & j).Interior.ColorIndex = 4
            
            Else
                Range("K" & j).Interior.ColorIndex = 3
            
            End If
                       
            
        Next j
        
'Set variables for Greatest Summary Table Values
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_total As String
    Dim total_ticker As Range

    'find greatest % increase in column k and place in the table
    greatest_increase = WorksheetFunction.Max(Range("K2", "K" & summary_count))
    Range("Q2") = greatest_increase
    Range("Q2").NumberFormat = "0.00%"
    
    'find ticker for greatest % increase in column I and place in table
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & summary_count)), Range("K2:K" & summary_count), 0)
    Range("P2") = Cells(increase_number + 1, 9)
    
    'find greatest % decrease in column k and place in the table
    greatest_decrease = WorksheetFunction.Min(Range("K2", "K" & summary_count))
    Range("Q3") = greatest_decrease
    Range("Q3").NumberFormat = "0.00%"
    
    'find ticker for greatest % decrease in column I and place in table
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & summary_count)), Range("K2:K" & summary_count), 0)
    Range("P3") = Cells(decrease_number + 1, 9)
    
    'find greatest total in column l & place it in the table
    greatest_total = WorksheetFunction.Max(Range("L:L"))
    Range("Q4") = greatest_total
     
    'find ticker for greatest total in column I & place in the table
    Set total_ticker = Range("L:L").Find(What:=greatest_total)
    Range("P4").Value = total_ticker.Offset(, -3).Value
        
        'Format columns L to increase width
        ActiveSheet.UsedRange.EntireColumn.AutoFit
        
    Next ws
End Sub
