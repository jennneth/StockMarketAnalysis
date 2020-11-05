Attribute VB_Name = "Module1"
Sub StockAnalysis()

'---------------------------------------------
'Create some variables to hold the Summary data
Dim GreatestPercIncr As Double
Dim GreatestPercDecr As Double
Dim GreatestTotalVol As Double
Dim GreatestPercIncr_Ticker As String
Dim GreatestPercDecr_Ticker As String
Dim GreatestTotalVol_Ticker As String

'---------------------------------------------
'-----
'Loop through each worksheet
'-----
For Each ws In Worksheets

'1. Put the headers in to indicate the data being displayed.
'Headers will be Ticker, Yearly Change in Price, % Change from Opening to Closing, Total Stock Volume
'For the sake of debugging, opening and closing price will also be displayed
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change in Price"
ws.Range("K1") = "% Change"
ws.Range("L1") = "Total Stock Volume"
'ws.Range("M1") = "Opening Price"
'ws.Range("N1") = "Closing Price"

'2. Calculate the ticker, opening price and closing price for the year
'Note that worksheets are already organized by year and do not contain data that extends beyond a year
'Note that worksheet data is already organized by year/month/day,
'so the opening price of the year is the first value and the closing price for the year is the last value for that ticker

'Store the ticker symbol, opening price, stock volume
'Create variables
Dim Ticker As String
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim TotalStockVolume As Double

'Initialize the variables
Ticker = ws.Range("A2")
OpeningPrice = ws.Range("C2")
ClosingPrice = ws.Range("F2")
TotalStockVolume = ws.Range("G2")

'Store the location of each ticker symbol in the summary table
Dim SummaryTableRow As Integer
SummaryTableRow = 2

'Store the last row in the sheet
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create a Variable to use to store the PriceDelta
Dim PriceDelta As Double
PriceDelta = 0
'Create a Variable to use to store the Percent Change
Dim DeltaPercent As Double
DeltaPercent = 0

'Loop through all of the ticker data rows
For i = 2 To LastRow
    'if the next cell has the same ticker symbol
    If ws.Cells(i + 1, 1).Value = Ticker Then
        'add the stock volume to total stock volume
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    'otherwise the next cell is a different ticker
    Else
        'add the stock volume to total stock volume
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        'store the closing price
        ClosingPrice = ws.Cells(i, 6).Value
        'calculate the delta in opening price and closing price
        PriceDelta = ClosingPrice - OpeningPrice
        'calculate the percentage change
        'note the error if dividing by 0 of the OpeningPrice is 0
        If OpeningPrice = 0 Then
            DeltaPercent = 0
        Else
            'DeltaPercent = Round(PriceDelta / OpeningPrice, 2)
            DeltaPercent = (PriceDelta / OpeningPrice)
        End If
        'update the table with ticker, opening price(*), closing price(*),delta,  percentage and total stock volume
        ws.Cells(SummaryTableRow, 9) = Ticker
        ws.Cells(SummaryTableRow, 10) = PriceDelta
        ws.Cells(SummaryTableRow, 11) = FormatPercent(DeltaPercent, 2)
        ws.Cells(SummaryTableRow, 12) = TotalStockVolume
        'ws.Cells(SummaryTableRow, 13) = OpeningPrice
        'ws.Cells(SummaryTableRow, 14) = ClosingPrice
    'format the results by highlilghting positive change in green and negative change in red
        If PriceDelta > 0 Then
            ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4 'color 4 is green
        Else
            ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3 'color 3 is red
        End If
'****** BONUS******
    'check whether the percent change is the greatest % increase or greatest % decrease or greatest total volume
   
    'if the DeltaPercent is positive, check if it is larger than what was previously stored.  If yes, update the largest percent and the ticker
    If DeltaPercent > 0 Then
        If DeltaPercent > GreatestPercIncr Then
            GreatestPercIncr = DeltaPercent
            GreatestPercIncr_Ticker = Ticker
        End If
    'if the deltapercent is negative, check if it is more negative than what was previously stored. if yes. update the largest percent deduction and the ticker
    ElseIf DeltaPercent < 0 Then
        If DeltaPercent < GreatestPercDecr Then
            GreatestPercDecr = DeltaPercent
            GreatestPercDecr_Ticker = Ticker
        End If
    End If
    'if the TotalStockVolume is larger than  the GreatestTotalVol, update the GreatestTotalVol with TotalStockVolume and Ticker
    If TotalStockVolume > GreatestTotalVol Then
        GreatestTotalVol = TotalStockVolume
        GreatestTotalVol_Ticker = Ticker
    End If
'****END BONUS******
    'reset the variables
    TotalStockVolume = 0
    Ticker = ws.Cells(i + 1, 1).Value
    OpeningPrice = ws.Cells(i + 1, 3).Value
    PriceDelta = 0
    DeltaPercent = 0
       
    'increment the row counter for where to write the final table
    SummaryTableRow = SummaryTableRow + 1
    End If
Next i

'********************
'BONUS
'********************
' display the summary values for the worksheet
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"

ws.Cells(2, 17).Value = GreatestPercIncr_Ticker
ws.Cells(3, 17).Value = GreatestPercDecr_Ticker
ws.Cells(4, 17).Value = GreatestTotalVol_Ticker
ws.Cells(2, 18).Value = FormatPercent(GreatestPercIncr, 2)
ws.Cells(3, 18).Value = FormatPercent(GreatestPercDecr, 2)
ws.Cells(4, 18).Value = GreatestTotalVol

'Reset Greatest variables
GreatestPercIncr_Ticker = ""
GreatestPercDecr_Ticker = ""
GreatestTotalVol_Ticker = ""
GreatestPercIncr = 0
GreatestPercDecr = 0
GreatestTotalVol = 0

'go to the next worksheetloop
Next ws

End Sub
