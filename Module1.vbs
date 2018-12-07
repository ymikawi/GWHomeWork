Attribute VB_Name = "Module1"
Sub HomeworkStockmarket()
Dim Total_Stock_Value As String
Dim Stock_Total As Double
Stock_Total = 0
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For Each ws In Worksheets
    
  ' Loop through all Stock values
        For I = 2 To Lastrow

    ' Check if we are still within the same ticker, if it is not...
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      ' Set the ticker name
        Ticker = Cells(I, 2).Value

      ' Add to the Ticker Total
        Ticker_Total = Ticker_Total + Cells(I, 7).Value
        
        Ticker_Total = Range("I9").Value
        
        Stock_Total = Range("J9").Value
       
      ' Reset the ticker Total
        Ticker_Total = 0

    '   If the cell immediately following a row is the same brand...
        Else

      '     Add to the Brand Total
        Ticker_Total = Ticker_Total + Cells(I, 3).Value
        End If
    Next I
    Next ws

End Sub

