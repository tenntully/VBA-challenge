Attribute VB_Name = "Module1"
Sub analyze_stock_data():
    'This script analyzes stock data to find annual changes for each stock ticker in price, percentage, and volume.
    'Jason Hanlin - 2/1/23
    'For Data Analytics GT Boot Camp
    
    '!!!Assumes data is sorted by stock ticker and then by date (earliest to latest)!!!
    
    Dim column As Integer
    Dim row As Long
    Dim resultsRow As Integer
    Dim ticker As String
    Dim PriceStart As Double
    Dim PriceEnd As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim StockVolume As Double
    Dim WorksheetName As String
    Dim GreatPercentIncrease As Double
    Dim GreatPercentDecrease As Double
    Dim GreatStockVolume As Double
    Dim indexRow As Range
    Dim assignTicker As String
    Dim lrow As Long
    
    
    For Each ws In Worksheets       'Apply to each worksheet in workbook
    
        ' This will create the new headers in the sheet and format
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greastest % Increase"
        ws.Cells(3, 15).Value = "Greastest % Decrease"
        ws.Cells(4, 15).Value = "Greastest Total Volume"
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Rows(1).Font.FontStyle = "Bold"
        ws.Columns(15).Font.FontStyle = "Bold"
        
        ' Set number of rows of data, assumes data in first column is complete
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'Initialize Greatest variables
        GreatPercentIncrease = 0
        GreatPercentDecrease = 0
        GreatStockVolume = 0
        
        'Initialize first row
        row = 2
        ticker = ws.Cells(row, 1).Value
        PriceStart = ws.Cells(row, 3).Value
        StockVolume = 0
        resultsRow = 2
        ws.Cells(resultsRow, 9) = ticker
        
        'Cycle through rows to calculate summary ticker data
        For row = 2 To lastRow
            StockVolume = StockVolume + ws.Cells(row, 7).Value
            
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then  'Goes until the ticker symbol changes, then runs delta calculations
                PriceEnd = ws.Cells(row, 6).Value
                VolumeEnd = ws.Cells(row, 7).Value
                YearlyChange = PriceEnd - PriceStart
                PercentChange = YearlyChange / PriceStart
                
                'Record the data
                ws.Cells(resultsRow, 10).Value = YearlyChange
                ws.Cells(resultsRow, 11).Value = PercentChange
                ws.Cells(resultsRow, 12).Value = StockVolume
                
                'Reset with new ticker symbol
                ticker = ws.Cells(row + 1, 1).Value
                StockVolume = 0
                resultsRow = resultsRow + 1
                ws.Cells(resultsRow, 9) = ticker
                
            End If
        Next row
        
        'Conditional format the yearly change column
        'Green if equal to or greater than zero, red if less than
        For row = 2 To lastRow
            If ws.Cells(row, 10) < 0 Then
                ws.Cells(row, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(row, 10).Interior.ColorIndex = 4
            End If
        Next
        
        'Calculate overall ticker summary data for the worksheet, including max increase, max decrease, max stock volume.
        'Finds the value first.  Then iterates through with For To and If to find the corresponding ticker name.
        GreatPercentIncrease = Application.WorksheetFunction.Max(ws.Range("k:k"))
        ws.Range("Q2").Value = GreatPercentIncrease
        For row = 2 To lastRow
            If ws.Cells(row, 11) = GreatPercentIncrease Then
                ws.Range("P2").Value = ws.Cells(row, 9)
                Exit For
            End If
        Next
        
        GreatPercentDecrease = Application.WorksheetFunction.Min(ws.Range("k:k"))
        ws.Range("Q3").Value = GreatPercentDecrease
        For row = 2 To lastRow
            If ws.Cells(row, 11) = GreatPercentDecrease Then
                ws.Range("P3").Value = ws.Cells(row, 9)
                Exit For
            End If
        Next
        
        GreatStockVolume = Application.WorksheetFunction.Max(ws.Range("l:l"))
        ws.Range("Q4").Value = GreatStockVolume
        For row = 2 To lastRow
            If ws.Cells(row, 12) = GreatStockVolume Then
                ws.Range("P4").Value = ws.Cells(row, 9)
                Exit For
            End If
        Next
   
    Next ws
    
End Sub

