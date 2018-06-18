Sub stock_data()

  ' Set an initial variable for holding the stock name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per each stock Ticker
  Dim Ticker_Total As Double
  Ticker_Total = 0

  ' Keep track of the location for each stock namein the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

    'Set values for Header
    Range("J1").value = "Ticker"
    Range("K1").value = "Total Stock Volume"


  ' Loop through all credit card purchases
  For i = 2 To 800000

    ' Check if we are still within the same stocker Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Brand Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value

      ' Print the stock Ticker in the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Ticker Total to the Summary Table
      Range("K" & Summary_Table_Row).Value = Ticker_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Ticker_Total = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value

    End If

  Next i

End Sub