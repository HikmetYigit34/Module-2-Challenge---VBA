'please RUN only stockDataForAllYears(), loop through all 3 sheets
'before hitting run button, make sure the cursor is within stockDataForAllYears() function code, or out of both function codes.
'stockDataForYear() is called from the first function

Sub stockDataForAllYears()
    For yr = 1 To 3                     'we have total of 3 sheets(years)
        Sheets(yr).Activate             'activate sheets one by one and call for function stockDataForYear evry time
        stockDataForYear                'Call for yearly function, instead of this line all stockDataForYear() function can be copied here,
                                        'i think more descriptive this way, for diagnostics, sheets can be run one by one as well
    Next yr
    Sheets(1).Activate                  'return back to sheet 1, original view
End Sub



Sub stockDataForYear()

'variable declarations
Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim stockVolume As Double
Dim yearlyChange As Double
Dim percentChange As Double

Dim tickers(3000) As String             'once per ticker in worksheet
Dim openPrices(3000) As Double          'open price per ticker
Dim closePrices(3000) As Double         'close price per ticker
Dim totalStockVolumes(3000) As Double   'totalized stock value per ticker

Dim n As Integer
Dim tickerVolumeTotalizer As Double
Dim rowCount As Double
Dim greatestPercentIncrease As Double
Dim greatestPercentIncreaseTicker As String
Dim greatestPercentDecrease As Double
Dim greatestPercentDecreaseTicker As String
Dim greatestTotalVolume As Double
Dim greatestTotalVolumeTicker As String

'data collections
'collect tickers and open prices---------------------------------------------------------------------------------------
n = 0                                                                               'starting values
If (ActiveSheet.Name = ("2018")) Then rowCount = 753001
If (ActiveSheet.Name = ("2019")) Then rowCount = 756001
If (ActiveSheet.Name = ("2020")) Then rowCount = 759001

For i = 2 To rowCount                                                               'sheets have different row counts so it is dynamic
    ticker = ActiveSheet.Cells(i, 1).Value                                          'reading value from worksheet and storing to a variable
    openPrice = ActiveSheet.Cells(i, 3).Value                                       'reading value from worksheet and storing to a variable
    If (ticker <> ActiveSheet.Cells(i - 1, 1).Value) Then
        tickers(n) = ticker                                                         'assign ticker name
        openPrices(n) = openPrice                                                   'assign open price
        n = n + 1
    End If
Next i

'collect close prices and totalize stock volumes-----------------------------------------------------------------------
n = 0: tickerVolumeTotalizer = 0                                                    'starting values
For i = 2 To rowCount                                                               'sheets have different row counts so it is dynamic
    ticker = ActiveSheet.Cells(i, 1).Value                                          'reading value from worksheet and storing to a variable
    closePrice = ActiveSheet.Cells(i, 6).Value                                      'reading value from worksheet and storing to a variable
    stockVolume = ActiveSheet.Cells(i, 7).Value                                     'reading value from worksheet and storing to a variable
    tickerVolumeTotalizer = tickerVolumeTotalizer + stockVolume                     'keep adding volumes for the same ticker
    If (ticker <> ActiveSheet.Cells(i + 1, 1).Value) Then
        closePrices(n) = closePrice                                                 'assign close price
        totalStockVolumes(n) = tickerVolumeTotalizer                                'assign ticker total volume if it is the last ticker
        tickerVolumeTotalizer = 0                                                   'reset totalizer for the next ticker
        n = n + 1
    End If
Next i


'return values
'first row, titles col 9, col 10, col 11 ------------------------------------------------------------------------------
ActiveSheet.Cells(1, 9).Value = "Ticker"
ActiveSheet.Cells(1, 10).Value = "Yearly Change"
ActiveSheet.Cells(1, 11).Value = "Percent Change"
ActiveSheet.Cells(1, 12).Value = "Total Stock Volume"
ActiveSheet.Cells(1, 15).Value = "Ticker"
ActiveSheet.Cells(1, 16).Value = "Value"

'other rows
'yearly change---------------------------------------------------------------------------------------------------------
For j = 0 To n - 1                                                                  'n is length of array, is 3000 so, it is 0 to 2999
    ActiveSheet.Cells(j + 2, 9).Value = tickers(j)
    yearlyChange = closePrices(j) - openPrices(j)
    
    'conditional color formatting for yearly change
    If (yearlyChange < 0) Then ActiveSheet.Cells(j + 2, 10).Interior.Color = vbRed Else ActiveSheet.Cells(j + 2, 10).Interior.Color = vbGreen
    ActiveSheet.Cells(j + 2, 10) = yearlyChange                                     'writing output to worksheet
    ActiveSheet.Cells(j + 2, 10).NumberFormat = "$#,##0.00"                         'adding $ sign
    
    'percentage------------------------------------------------------------------------------------------------------------
    percentChange = (closePrices(j) - openPrices(j)) / openPrices(j)
    'format percentage cell for 2 decimal places
    ActiveSheet.Cells(j + 2, 11).Value = Format(percentChange, "#.00%")

    'conditional color formatting for percentage
    If (ActiveSheet.Cells(j + 2, 11).Value < 0) Then ActiveSheet.Cells(j + 2, 11).Interior.Color = vbRed Else ActiveSheet.Cells(j + 2, 11).Interior.Color = vbGreen

    'total stock volumes---------------------------------------------------------------------------------------------------
    ActiveSheet.Cells(j + 2, 12).Value = totalStockVolumes(j)

    'greatest percent increase---------------------------------------------------------------------------------------------
    If (ActiveSheet.Cells(j + 2, 11).Value > greatestPercentIncrease) Then
        greatestPercentIncrease = ActiveSheet.Cells(j + 2, 11)
        greatestPercentIncreaseTicker = ActiveSheet.Cells(j + 2, 9)
        ActiveSheet.Cells(2, 14).Value = "Greatest % Increase"
        ActiveSheet.Cells(2, 15).Value = greatestPercentIncreaseTicker
        ActiveSheet.Cells(2, 16).Value = Format(greatestPercentIncrease, "#.00%")
    End If

    'greatest percent decrease---------------------------------------------------------------------------------------------
    If (ActiveSheet.Cells(j + 2, 11).Value < greatestPercentDecrease) Then
        greatestPercentDecrease = ActiveSheet.Cells(j + 2, 11)
        greatestPercentDecreaseTicker = ActiveSheet.Cells(j + 2, 9)
        ActiveSheet.Cells(3, 14).Value = "Greatest % Decrease"
        ActiveSheet.Cells(3, 15).Value = greatestPercentDecreaseTicker
        ActiveSheet.Cells(3, 16).Value = Format(greatestPercentDecrease, "#.00%")
    End If

    'greatest total volume-------------------------------------------------------------------------------------------------
    If (ActiveSheet.Cells(j + 2, 12).Value > greatestTotalVolume) Then
        greatestTotalVolume = ActiveSheet.Cells(j + 2, 12)
        greatestTotalVolumeTicker = ActiveSheet.Cells(j + 2, 9)
        ActiveSheet.Cells(4, 14).Value = "Greatest Total Volume"
        ActiveSheet.Cells(4, 15).Value = greatestTotalVolumeTicker
        ActiveSheet.Cells(4, 16).Value = greatestTotalVolume
    End If

Next j




End Sub

