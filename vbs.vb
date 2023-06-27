Sub homeworkloops()
    'loop though each work sheeet and define it as string
    For Each ws In Worksheets
    Dim worksheetName As String
    
  ' Set an initial variable for holding the brand name
  Dim stock_Name As String

  ' Set an initial variable for holding the total per stock brand
  Dim stock_Total As Double
  stock_Total = 0

  ' Keep track of the location for each stock brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
    
  ' creating a variable for holding yearly change
    Dim yearly_change As Integer
    yearly_change = ws.Cells(i, 6) - ws.Cells(2, 3)

'creating a variable for holding percent change of a stock
    Dim percent_change As Integer
    percent_change = yearly_change / ws.Cells(i, 7).Value
    
    
    
  ' define last_row
  last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
  ' Loop through all stock totals
  For i = 2 To last_row

    ' Check if we are still within the same stock brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the stocks name
      stock_Name = Cells(i, 1).Value

      ' Add to the stock Total
      stock_Total = Brand_Total + Cells(i, 2).Value

      ' Print the stock Brand in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = stock_Name

      ' Print the total Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = stock_Total
      
      'print tthe yearly change to the summary tabel
      ws.Range("K" & Summary_Table_Row).Value = percent_change
      'print the percent change to the summary tabel
      ws.Range("J" & Summary_Table_Row).Value = yearly_Total
         
         
      ' Add one to the summary table row
      ws.Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      stock_Total = 0
      ' reset the yearly_change
      yearly_change = 0
      'reset the percent change
      percent_change = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      stock_Total = stock_Total + Cells(i, 7).Value

    End If

  Next i
Next ws
End Sub

