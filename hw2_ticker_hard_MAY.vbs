Sub ticker()

    'set iterator for all worksheets
    Dim k As Integer
    
    'variable for number for worksheets in the workbook
    Dim wscount As Integer
    
    'Identify number of worksheets in the workbook
    wscount = ActiveWorkbook.Worksheets.Count
    
    'iterate through all worksheets
    For k = 1 To wscount
    
    'activate current worksheet in interation
    Worksheets(k).Activate
    
    'set iterator for all data points in sheet
    Dim i As Long
    
    
    'set initial ticker volume to zero
    Dim volume As Double
    volume = 0
    
    'set open and closing stock price as variables in addition to percentage change
    Dim stock_close As Double
    Dim stock_open As Double
    Dim stock_change As Double
    Dim perc_change As Double
            
    'variable for stock ticker
    Dim ticker As String
    
    'set initial summary table row to 2
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'count number of total data points on sheet
    Dim lrow As Long
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'cell titles for summary table
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Volume"
    
    'iterator through all data in worksheet
    For i = 2 To lrow
    
        'find individual stock tickers assuming data set is ordered
        'finds when next cell ticker is different and pulls out current ticker value and places in summary table
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            ticker = Cells(i, 1).Value
            
            'calculate cumulative volume
            volume = volume + Cells(i, 7).Value
            
            'assign tickers and volumes to summary table
            Range("I" & Summary_Table_Row) = ticker
            Range("L" & Summary_Table_Row) = volume
            
            'stock close price is also the last closing price for any set of unique tickers
            stock_close = Cells(i, 6).Value
            
            'stock open will be calculated in next if statement.
            'determine change in stock price and place in summary table
            stock_change = stock_close - stock_open
            Range("J" & Summary_Table_Row) = stock_change
            
            'avoids divide by zero scenario
            If stock_open = 0 Then
            perc_change = 0
            
            Else: perc_change = stock_close / stock_open - 1
            
            End If
            
            'places percentage change into summary table
            Range("K" & Summary_Table_Row) = perc_change
            
            'reset volume to zero for each ticker when individual sum is complete
            volume = 0
            
            'advances summary table rows
            Summary_Table_Row = Summary_Table_Row + 1
       
            
        'determines first ticker in each group of data and identifies opening stock value
        ElseIf (Cells(i - 1, 1).Value <> Cells(i, 1).Value) Then
            stock_open = Cells(i, 3).Value
          
        'continues to sum volume
        Else: volume = volume + Cells(i, 7).Value
        
        End If
        
        
    Next i
    
'variable and calculation for number of data points in the summary table
Dim lrow_summary As Long
lrow_summary = Cells(Rows.Count, 9).End(xlUp).Row

'declare variables for max gains, losses and volumes and associated tickers
Dim j As Integer
Dim MaxGain As Double
Dim MaxLoss As Double
Dim MaxVolume As Double
Dim MaxGainTicker As String
Dim MaxLossTicker As String
Dim MaxVolumeTicker As String


'label cells for summary table
Range("O1") = "Ticker"
Range("P1") = "Value"
Range("N2") = "Greatest Percent Increase"
Range("N3") = "Greatest Percent Loss"
Range("N4") = "Greatest Total Volume"

'excel worksheet functions to determine max gain, loss and volume values from the summary table
MaxGain = Application.WorksheetFunction.Max(Range("K2:K" & lrow_summary))
MaxLoss = Application.WorksheetFunction.Min(Range("K2:K" & lrow_summary))
MaxVolume = Application.WorksheetFunction.Max(Range("L2:L" & lrow_summary))

'places values into summary table and formats max gain and max loss as percentage
Range("P2") = MaxGain
Range("P3") = MaxLoss
Range("P4") = MaxVolume
Range("P2").NumberFormat = "0.00%"
Range("P3").NumberFormat = "0.00%"

    'iterator to format cells in summary table for gains/losses
    For j = 2 To lrow_summary
    
        If (Cells(j, 10).Value > 0) Then
            Cells(j, 10).Interior.ColorIndex = 4
        
            Else: Cells(j, 10).Interior.ColorIndex = 3
    
        End If
    
        'logical test to determine which row contains the max gain, max loss, and max volumes and identifies associated ticker
        If (Cells(j, 11).Value = MaxGain) Then
            MaxGainTicker = Cells(j, 9).Value
            
        ElseIf (Cells(j, 11).Value = MaxLoss) Then
            MaxLossTicker = Cells(j, 9).Value
            
        End If
        
        If (Cells(j, 12).Value = MaxVolume) Then
            MaxVolumeTicker = Cells(j, 9).Value
            
        End If
                
    'formats all gains and losses as a percentage in the summary table
    Cells(j, 11).NumberFormat = "0.00%"
    
    Next j
    
'assigns tickers to final summary table
Range("O2") = MaxGainTicker
Range("O3") = MaxLossTicker
Range("O4") = MaxVolumeTicker

'advances iteration to next worksheet
Next k
    
End Sub

    

