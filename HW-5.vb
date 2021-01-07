

Public Sub GenerateSummaryTable()


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


' Insert headers for Summary Table
range("I1").Value = "Ticker"
range("J1").Value = "Yearly Change"
range("K1").Value = "Percent Change"
range("L1").Value = "Total Stock Volume"

' Auto fit above columns
Columns("I:L").AutoFit

' Loop through all entries, beginning at row 2
lastDataRow = Cells(Rows.Count, 1).End(xlUp).Row

' Start Summary Table at row 2 to allow for headers in row 1.
reportTableRow = 2

' Loop through each row containing data to create summary table data
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
        ' Generate Summary Table columns for ticker and total volume
        range("I" & reportTableRow).Value = tickerName
        range("L" & reportTableRow).Value = totalVolume
        range("L" & reportTableRow).NumberFormat = "0"
        
        ' Set closingPrice variable to current row's closing price
        closingPrice = Cells(i, 6).Value
        ' Set openingPrice variable to current row's opening price
        openingPrice = Cells(i - currentTickerCounter, 3).Value
        ' Calculate yearly change in current ticker
        yearlyChangeValue = closingPrice - openingPrice
        ' Calculate yearly change in current ticker as percent
        yearlyChangePercent = (closingPrice / openingPrice) - 1
        ' Generate Summary Table column for yearly change in value
        range("J" & reportTableRow).Value = yearlyChangeValue
        range("J" & reportTableRow).NumberFormat = "0.00"
        ' Format yearly change in value red for negative and green for positive
        If yearlyChangeValue < 0 Then
            range("J" & reportTableRow).Interior.Color = vbRed
            
        Else
            range("J" & reportTableRow).Interior.Color = vbGreen
        End If
        ' Generate Summary Table column for yearly change in percent
        range("K" & reportTableRow).Value = yearlyChangePercent
        range("K" & reportTableRow).NumberFormat = "0.00%"
        ' Reset loop variables for next tickers' iteration
        reportTableRow = reportTableRow + 1
        totalVolume = 0
        currentTickerCounter = 0
          
                
    End If
    
    Next i
    
Call BonusSummaryChart

End Sub



Sub BonusSummaryChart()
'Declare variables to calculate and generate Bonus table
Dim lastPercentChangeRow As Long
Dim lastTotalVolumeRow As Long
Dim maxValuePercent As Double
Dim minValuePercent As Double
Dim greatestTotalVolume As Currency
Dim maxValuePercentTicker As String
Dim minValuePercentTicker As String
Dim greatestTotalVolumeTicker As String

' Insert headers
range("O2") = "Greatest % Increase"
range("O3") = "Greatest % Decrease"
range("O4") = "Greatest Total Volume"
range("P1") = "Ticker"
range("Q1") = "Value"
' Find end of data for both columns
lastPercentChangeRow = Cells(Rows.Count, 11).End(xlUp).Row
lastTotalVolumeRow = Cells(Rows.Count, 12).End(xlUp).Row
    

'Find maximum % change value in column K
maxValuePercent = Application.WorksheetFunction.Max(range("K2:K" & lastPercentChangeRow))
' Select bonus summary table cell to insert value into
range("Q2").Select
' Insert value
ActiveCell.Value = maxValuePercent
' Format number as percent
ActiveCell.NumberFormat = "0.00%"

' Find and enter minimum % change values and format it
minValuePercent = Application.WorksheetFunction.Min(range("K2:K" & lastPercentChangeRow))
range("Q3").Select
ActiveCell.Value = minValuePercent
ActiveCell.NumberFormat = "0.00%"

' Find and enter greatest total volume value and format it
greatestTotalVolume = Application.WorksheetFunction.Max(range("L2:L" & lastTotalVolumeRow))
range("Q4").Select
ActiveCell.Value = greatestTotalVolume
ActiveCell.NumberFormat = "0"

Call FindTickerNames

' Auto fit bonus table
Columns("O:Q").AutoFit
    
End Sub



Sub FindTickerNames()
'Macro to find ticker names for bonus summary table
    
    ' Select greatest % increase value to search for it's ticker name
    range("Q2").Select
    Selection.Copy
    ' Select range to search in
    Columns("K:K").Select
    ' Find value in range and activate found cell
    Selection.Find(What:="107.14%", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ' Select cell of first found instance
    ActiveCell.Select
    ' Move two cells to the left to select ticker name's cell
    ActiveCell.Offset(0, -2).Select
    ' Copy ticker name
    Selection.Copy
    ' Paste ticker name into bonus summary table
    range("P2").Select
    ActiveSheet.Paste
    
    ' Repeat above process for greatest % decrease value
    range("Q3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("K:K").Select
    Selection.Find(What:="-3.01%", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Select
    ActiveCell.Offset(0, -2).Select
    Application.CutCopyMode = False
    Selection.Copy
    range("P3").Select
    ActiveSheet.Paste
    
    ' Repeat above process for greatest total volume value
    range("Q4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("L:L").Select
    Selection.Find(What:="77251236600", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Select
    ActiveCell.Offset(0, -3).Select
    Application.CutCopyMode = False
    Selection.Copy
    range("P4").Select
    ActiveSheet.Paste
End Sub


