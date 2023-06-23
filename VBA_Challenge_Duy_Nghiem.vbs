Sub StockSheets()

'to do list (pseudocode):
'make  new columns for ticker, yearly change, percent change, total stock volume

'Loop through all stocks
    'Read:
        'ticker symbol
        'Yearly change (closing price - opening price)
        'Percent change ((closing price - opening price)/opening price) or yearly change/opening price
        'Total stock volume
    'needs loop for worksheets
    'loops for start to end of worksheet
        'if statement for ticker
            'if they are the same, don't add anything to spreadsheet
                'only add stockvolume
            'if they are not, then:
                'Get Ticker
                'Get ClosePrice
                'Calculate Yearly/percent change
                'add last stock volume for the stock
                'print everything into a row
                'format numbers on cell properly
                'reset stockvolume
                'make new opening price
                'NewRow++

'Format cells (color)
    'loop to end
    '<0 = red
    '>=0 = green
    
'add functionality (Greatest % Increase/Decrease, Greatest Total Volume)
    'loop through new rows
        'if statement for max
        'if statement for min
        'if statement for volumemax
        

'Make VBA worksheet run on all worksheets

Dim WS As Worksheet

'Loop through Worksheets
For Each WS In ActiveWorkbook.Worksheets
WS.Activate
    
    'make variables

    'counter are long in case more than 32k rows
    Dim i As Long
    
    'part 1 variables
    Dim GIncrease As String
    Dim GDecrease As String
    Dim GVolume As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim StockVolume As Double

    Dim Ticker As String
    
    Dim RowCount As Integer
    Dim Columns As Integer
    
    'part 2 variables
    Dim NewRow As Long
    Dim NewRowCount As Integer
    
    'Part 3 variables
    Dim IncreaseMax As Double
    Dim DecreaseMin As Double
    Dim VolumeMax As Double
    Dim MaxTicker As String
    Dim MinTicker As String
    Dim VolumeTicker As String

    'Initialize variables
    GIncrease = "Greatest % Increase"
    GDecrease = "Greatest % Decrease"
    GVolume = "Greatest Total Volume"
    
    RowCount = 0
    'column variable just in case
    Columns = 1
    'starts at 2 since there is headers at 1
    NewRow = 2
    
    OpenPrice = 0
    ClosePrice = 0
    YearlyChange = 0
    PercentChange = 0
    StockVolume = 0
    
    IncreaseMax = 0
    IncreaseMin = 0
    VolumeMax = 0
    
    'set headers
    
    'NewRows
    Range("i1:l1").Value = [{"Ticker", "Yearly Change", "Percent Change", "Total Stock Volume"}]
    
    'Functionality Section
    Range("p1:q1").Value = [{"Ticker", "Value"}]
    Range("o2").Value = GIncrease
    Range("o3").Value = GDecrease
    Range("o4").Value = GVolume
    
    'setting RowCount
    RowCount = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Initial OpenPrice
    OpenPrice = Range("C2").Value
    
    'loop through whole sheet
    For i = 2 To RowCount
      'check for Tickers
        If Cells(i + 1, Columns).Value = Cells(i, 1).Value Then
        'add StockVolumes
            StockVolume = StockVolume + Cells(i, 7)
        'If the tickers are not the same

        Else
            'initialize Ticker
            Ticker = Cells(i, 1).Value
            'print ticker
            Cells(NewRow, 9).Value = Ticker
            'change format to general
            Cells(NewRow, 9).NumberFormat = "General"
            
            'set Close Price for current ticker
            ClosePrice = Cells(i, 6)
            
            'Calculate Yearly Change
            YearlyChange = ClosePrice - OpenPrice
            'print yearly change
            Cells(NewRow, 10).Value = YearlyChange
            'change format to number
            Cells(NewRow, 10).NumberFormat = "0.00"
            
            'Calculate Percent Change
            PercentChange = YearlyChange / OpenPrice
            Cells(NewRow, 11).Value = PercentChange
            'change format to percent
            Cells(NewRow, 11).NumberFormat = "0.00%"
            
            'add last StockVolume
            StockVolume = StockVolume + Cells(i, 7)
            'print StockVolumes
            Cells(NewRow, 12).Value = StockVolume
            'set format to general
            Cells(NewRow, 12).NumberFormat = "General"
            
            'set StockVolume to 0 for next iteration
            StockVolume = 0
            
            'set up values for next iteration
            'Opening Price (next cell, column 3)
            OpeningPrice = Cells(i + 1, 3)
            'NewRow (goes down 1)
            NewRow = NewRow + 1
            End If
    Next i
    
    'Change Yearly Change Cells from Green to Red
    
    NewRowCount = WS.Cells(Rows.Count, Columns + 8).End(xlUp).Row
    
    'loop through NewRow
    For i = 2 To NewRowCount
        'If Positive, turn green
        If Cells(i, 10) >= 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        'If Negative, turn red
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    'Greatest Increase/Decrease/Total Volume
    
    'set initial values for max, min, and volumemax
        'max
        MaxTicker = Cells(2, 9).Value
        IncreaseMax = Cells(2, 11).Value
        
        'min
        MinTicker = Cells(2, 9).Value
        DecreaseMin = Cells(2, 11).Value
        
        'VolumeMax
        VolumeTicker = Cells(2, 9).Value
        VolumeMax = Cells(2, 12).Value
    
    'For loop til the end
    For i = 2 To NewRowCount
        
        'Make if statement to determine max
        'if current max less than cell below, change that to max
        If IncreaseMax < Cells(i + 1, 11).Value Then
            'grab new ticker
            MaxTicker = Cells(i + 1, 9).Value
            'grab new max
            IncreaseMax = Cells(i + 1, 11).Value
        End If
        
        'Make if statement to determine min
        'if current min is greater than cell under, then change min to cell under
        If DecreaseMin > Cells(i + 1, 11) Then
            'grab new ticker
            MinTicker = Cells(i + 1, 9).Value
            'grab new min
            DecreaseMin = Cells(i + 1, 11).Value
        End If
        
        'Make if statement for volumemax
        'if current max less than cell below, get new max
        If VolumeMax < Cells(i + 1, 12).Value Then
            'grab new ticker
            VolumeTicker = Cells(i + 1, 9).Value
            'grab new max
            VolumeMax = Cells(i + 1, 12).Value
        End If
        
    Next i
    
    'print max
    Range("p2").Value = MaxTicker
    Range("q2").Value = IncreaseMax
    Range("q2").NumberFormat = "0.00%"
    
    'print min
    Range("p3").Value = MinTicker
    Range("q3").Value = DecreaseMin
    Range("q3").NumberFormat = "0.00%"
    
    'print volumemax
    Range("p4").Value = VolumeTicker
    Range("q4").Value = VolumeMax
    Range("q4").NumberFormat = "General"
        
Next WS

End Sub