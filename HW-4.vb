

Public Sub GetData()


' Declare variables for iterator, opening and closing ticker prices
' ticker and next ticker names, current report table row, total volume of ticker
' yearly value change and yearly value change in percent of each stock.
Dim i As Long
Dim lastDataRow As Long
Dim openingPrice As Double
Dim closingPrice As Double
Dim tickerName As String
Dim nextTickerName As String
Dim reportTableRow As Integer
Dim totalVolume As Currency
Dim yearlyChangeValue As Double
Dim yearlyChangePercent As Double
Dim greatestPercentIncrease As Double
Dim greatestPercentDecrease As Double
Dim greatestTotalVolume As Currency

' Start report at row 2 to allow for headers in row 1.
reportTableRow = 2

' Insert headers for Summary Tables
range("I1").Value = "Ticker"
range("J1").Value = "Yearly Change"
range("K1").Value = "Percent Change"
range("L1").Value = "Total Stock Volume"

Columns("I:L").AutoFit

' Loop through all entries, beginning at row 2
' TODO: change lower boundary to end of colum's data (xlDown)
lastDataRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastDataRow

    tickerName = Cells(i, 1).Value
    nextTickerName = Cells(i + 1, 1).Value
       
    ' Check whether current ticker is same as first ticker
    If (nextTickerName = tickerName) Then
    
        ' Add current row's volume to totalVolume variable
        totalVolume = totalVolume + Cells(i, 7).Value
        currentTickerCounter = currentTickerCounter + 1
                    
    Else
        totalVolume = totalVolume + Cells(i, 7).Value
        
        range("I" & reportTableRow).Value = tickerName
        range("L" & reportTableRow).Value = totalVolume
        range("L" & reportTableRow).NumberFormat = "0"
        
        ' Set closingPrice variable to current cell
        closingPrice = Cells(i, 6).Value
        openingPrice = Cells(i - currentTickerCounter, 3).Value
      
        yearlyChangeValue = closingPrice - openingPrice
        yearlyChangePercent = (closingPrice / openingPrice) - 1
        
        range("J" & reportTableRow).Value = yearlyChangeValue
        range("J" & reportTableRow).NumberFormat = "0.00"
        
        If yearlyChangeValue < 0 Then
            range("J" & reportTableRow).Interior.Color = vbRed
            
        Else
            range("J" & reportTableRow).Interior.Color = vbGreen
        End If
        
        range("K" & reportTableRow).Value = yearlyChangePercent
        range("K" & reportTableRow).NumberFormat = "0.00%"
        
        reportTableRow = reportTableRow + 1
        
        totalVolume = 0
        currentTickerCounter = 0
        
        
        ' Range("Q2") = Max(Range("K2:K6"))
        ' Range("Q3") = Min(Range("K2:K6"))
        ' Range("Q4") = Max(Range("L2:L6"))
        
                
    End If
    
    Next i
    
Call BonusSummaryChart

End Sub



Sub BonusSummaryChart()
'
' BonusSummaryChartValues Macro
Dim lastPercentChangeRow As Long
Dim lastTotalVolumeRow As Long
Dim maxValuePercent As Double
Dim minValuePercent As Double
Dim greatestTotalVolume As Currency
Dim maxValuePercentTicker As String
Dim minValuePercentTicker As String
Dim greatestTotalVolumeTicker As String

    
range("O2") = "Greatest % Increase"
range("O3") = "Greatest % Decrease"
range("O4") = "Greatest Total Volume"
range("P1") = "Ticker"
range("Q1") = "Value"

lastPercentChangeRow = Cells(Rows.Count, 11).End(xlUp).Row
lastTotalVolumeRow = Cells(Rows.Count, 12).End(xlUp).Row
    

'Find maximum and minimum % change values
maxValuePercent = Application.WorksheetFunction.Max(range("K2:K" & lastPercentChangeRow))
range("Q2").Select
ActiveCell.Value = maxValuePercent
ActiveCell.NumberFormat = "0.00%"


' maxValuePercentTicker = maxValuePercent.Select.Offset(0, -2).Value
' Range("P2").Select
' ActiveCell.Value = maxValuePercentTicker

 
minValuePercent = Application.WorksheetFunction.Min(range("K2:K" & lastPercentChangeRow))
range("Q3").Select
ActiveCell.Value = minValuePercent
ActiveCell.NumberFormat = "0.00%"

greatestTotalVolume = Application.WorksheetFunction.Max(range("L2:L" & lastTotalVolumeRow))
range("Q4").Select
ActiveCell.Value = greatestTotalVolume
ActiveCell.NumberFormat = "0"

Columns("O:Q").AutoFit
    
End Sub


