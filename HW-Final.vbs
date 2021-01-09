
Public Sub GenerateSummaryTable()

Application.ScreenUpdating = False
Application.Calculation = xlCalculateManual

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
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

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
        Range("I" & reportTableRow).Value = tickerName
        Range("L" & reportTableRow).Value = totalVolume
        Range("L" & reportTableRow).NumberFormat = "0"
        
        ' Set closingPrice variable to current row's closing price
        closingPrice = Cells(i, 6).Value
        ' Set openingPrice variable to current row's opening price
        openingPrice = Cells(i - currentTickerCounter, 3).Value
        ' Calculate yearly change in current ticker
        yearlyChangeValue = closingPrice - openingPrice
        ' Substitute 1 for any zero values to allow calculation below
        If openingPrice = 0 Then
        openingPrice = 1
        End If
        If closingPrice = 0 Then
        closingPrice = 1
        End If
        ' Calculate yearly change in current ticker as percent
        yearlyChangePercent = (closingPrice / openingPrice) - 1
        ' Generate Summary Table column for yearly change in value
        Range("J" & reportTableRow).Value = yearlyChangeValue
        Range("J" & reportTableRow).NumberFormat = "0.00"
        ' Format yearly change in value red for negative and green for positive
        If yearlyChangeValue < 0 Then
            Range("J" & reportTableRow).Interior.Color = vbRed
            
        Else
            Range("J" & reportTableRow).Interior.Color = vbGreen
        End If
        
        yearlyChangePercent = (closingPrice / openingPrice) - 1
        ' Generate Summary Table column for yearly change in value
        Range("J" & reportTableRow).Value = yearlyChangeValue
        Range("J" & reportTableRow).NumberFormat = "0.00"
        ' Format yearly change in value red for negative and green for positive
        If yearlyChangeValue < 0 Then
            Range("J" & reportTableRow).Interior.Color = vbRed
            
        Else
            Range("J" & reportTableRow).Interior.Color = vbGreen
        End If
        ' Generate Summary Table column for yearly change in percent
        Range("K" & reportTableRow).Value = yearlyChangePercent
        Range("K" & reportTableRow).NumberFormat = "0.00%"
        ' Reset loop variables for next tickers' iteration
        reportTableRow = reportTableRow + 1
        totalVolume = 0
        currentTickerCounter = 0
          
                
    End If
    
    Next i
    
Call BonusSummaryChart

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

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
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"
' Find end of data for both columns
lastPercentChangeRow = Cells(Rows.Count, 11).End(xlUp).Row
lastTotalVolumeRow = Cells(Rows.Count, 12).End(xlUp).Row
    

'Find maximum % change value in column K
maxValuePercent = Application.WorksheetFunction.Max(Range("K2:K" & lastPercentChangeRow))
' Select bonus summary table cell to insert value into
Range("Q2").Select
' Insert value
ActiveCell.Value = maxValuePercent
' Format number as percent
ActiveCell.NumberFormat = "0.00%"

' Find and enter minimum % change values and format it
minValuePercent = Application.WorksheetFunction.Min(Range("K2:K" & lastPercentChangeRow))
Range("Q3").Select
ActiveCell.Value = minValuePercent
ActiveCell.NumberFormat = "0.00%"

' Find and enter greatest total volume value and format it
greatestTotalVolume = Application.WorksheetFunction.Max(Range("L2:L" & lastTotalVolumeRow))
Range("Q4").Select
ActiveCell.Value = greatestTotalVolume
ActiveCell.NumberFormat = "0"

Call FindTickerNames

' Auto fit bonus table
Columns("O:Q").AutoFit
    
End Sub



Sub FindTickerNames()
'Macro to find ticker names for bonus summary table
Dim cell As Range
Dim find As String

    ' Select greatest % increase value to search for it's ticker name
    Range("Q2").Select
    Selection.Copy
    find = ActiveSheet.Paste
    ' Select range to search in
    Columns("K:K").Select
    ' Find value in range and activate found cell
    Set cell = Selection.find(What:=" & find & ", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    ' Select cell of first found instance
    If cell Is Nothing Then
    Else
    ActiveCell.Select
    ' Move two cells to the left to select ticker name's cell
    ActiveCell.Offset(0, -2).Select
    ' Copy ticker name
    Selection.Copy
    ' Paste ticker name into bonus summary table
    Range("P2").Select
    ActiveSheet.Paste
    End If
    
    ' Repeat above process for greatest % decrease value
    Range("Q3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("K:K").Select
    Set cell = Selection.find(What:=" & ActiveSheet.Paste & ", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If cell Is Nothing Then
    Else
    ActiveCell.Select
    ActiveCell.Offset(0, -2).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P3").Select
    ActiveSheet.Paste
    End If
    
    ' Repeat above process for greatest total volume value
    Range("Q4").Select
    ' Application.CutCopyMode = False
    Selection.Copy
    Columns("L:L").Select
    Set cell = Selection.find(What:=" & ActiveSheet.Paste & ", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If cell Is Nothing Then
    Else
    ActiveCell.Select
    ActiveCell.Offset(0, -3).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P4").Select
    ActiveSheet.Paste
    End If
    
    Application.CutCopyMode = False
    
End Sub








