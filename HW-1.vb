

Public Sub GetData()


' Declare variable for ticker symbol, iterator, first and last
' ticker cells, ticker name and total volume of ticker

Dim ticker As String
Dim i As Long
Dim openingPrice As Double
Dim closingPrice As Double
Dim tickerName As String
Dim nextTickerName As String
Dim reportTableRow As Integer
Dim totalVolume As Currency


reportTableRow = 2

' Loop through all entries, beginning at row 2
' TODO: change lower boundary to end of colum's data (xlDown)
For i = 2 To 72000

    tickerName = Cells(i, 1).Value
    nextTickerName = Cells(i + 1, 1).Value
       
    ' Check whether current ticker is same as first ticker
    If (nextTickerName = tickerName) Then
    
        ' Add current row's volume to totalVolume variable
        totalVolume = totalVolume + Cells(i, 7).Value
        currentTickerCounter = currentTickerCounter + 1
                    
    Else
        totalVolume = totalVolume + Cells(i, 7).Value

        Range("I" & reportTableRow).Value = tickerName
                
        Range("L" & reportTableRow).Value = totalVolume
        
        ' Set closingPrice variable to current cell
        closingPrice = Cells(i, 6).Value
        openingPrice = Cells(i - currentTickerCounter, 3).Value
       
        Range("J" & reportTableRow).Value = openingPrice
        Range("K" & reportTableRow).Value = closingPrice
        
        reportTableRow = reportTableRow + 1
        
        totalVolume = 0
        currentTickerCounter = 0
        
                
    End If
    
    Next i
    

End Sub
