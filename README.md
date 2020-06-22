# VBA-Challenge

Sub ticker()

' Set an initial variable for holding the brand name
  Dim ticker As String
  Dim summary_table_row As Integer
  Dim Lastrow As Long
  
' Set Row Header
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"
 
  summary_table_row = 2

  ' Loop through all ticker names
  Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To Lastrow
  
    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the TICKER name
      ticker = Cells(i, 1).Value
      
    ' Print the TICKER in the Summary Table
      Range("I" & summary_table_row).Value = ticker
      
    ' Add one to the summary table row
      summary_table_row = summary_table_row + 1

  End If
    Next i
    
End Sub

Sub totalstockvolume()

'Set variable for holding tickername
    Dim ticker As String
'Set an initial variable for holding the total stock vol per ticker
    Dim total_stock_vol As Variant
    total_stock_vol = 0
'Keep track of location for each ticker in summary table
    Dim summary_table_row As Integer
    summary_table_row = 2

'Loop through ticker data
  Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  For i = 2 To Lastrow
    
    ' Check if we are still within theticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    ' Add to the total stock vol
      total_stock_vol = total_stock_vol + Cells(i, 7).Value
      
    ' Print the Total stock vol to the Summary Table
      Range("L" & summary_table_row).Value = total_stock_vol

      ' Add one to the summary table row
      summary_table_row = summary_table_row + 1
      
      ' Reset the total_stock_vol
      total_stock_vol = 0

    ' If the cell immediately following a row is the same brand...
    Else
    
    ' Add to the total_stock_vol
     total_stock_vol = total_stock_vol + Cells(i, 7).Value

    End If
  Next i
  
End Sub

Sub yearcounts()

'Set variable for holding variables
Dim ticker As String
Dim year_change As Variant
Dim year_percentage As Variant
Dim year_open As Variant
Dim year_close As Variant

year_open = 0
year_close = 0
year_change = 0
year_percentage = 0

' Keep track of location for each ticker in summary table
Dim summary_table_row As Integer
summary_table_row = 2

'Loop through ticker data
  Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To Lastrow
  
    ' Check if we are still within theticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    ' Add to the year totals
    year_open = year_open + Cells(i, 3).Value
    year_close = year_close + Cells(i, 6).Value
    
    year_change = year_close - year_open
    year_percentage = (year_change / year_open)
    year_percentage = year_percentage * 100

    ' Print the Total stock vol to the Summary Table
      Range("J" & summary_table_row).Value = year_change
      Range("K" & summary_table_row).Value = year_percentage
      
    ' Add one to the summary table row
      summary_table_row = summary_table_row + 1
      
      ' Reset the values
    year_open = 0
    year_close = 0
    year_change = 0
    year_percentage = 0
         

    ' If the cell immediately following a row is the same brand...
    Else
    
    ' Add to the values
    year_open = year_open + Cells(i, 3).Value
    year_close = year_close + Cells(i, 6).Value


    End If
  Next i
  
  
End Sub

Sub color()

  Lastrow = Cells(Rows.Count, 9).End(xlUp).Row
  
  For i = 2 To Lastrow
  
    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
    Else
        Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i

End Sub

Sub challenge()


Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Greatest % decrease"
Cells(4, 15).Value = "Greatest total volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Dim g_increase As Double
Dim g_decrease As Double
Dim g_tvol As Variant
Dim ticker1 As String
Dim ticker2 As String
Dim ticker3 As String



'Loop for greatest % increase
  Lastrow = Cells(Rows.Count, 11).End(xlUp).Row
  
  For i = 2 To Lastrow
  g_increase = Application.WorksheetFunction.Max(Range("K:K"))
  Cells(2, 17).Value = g_increase
  
'Loop for greatest % decrease
  g_decrease = Application.WorksheetFunction.Min(Range("K:K"))
  Cells(3, 17).Value = g_decrease


'Loop for greatest total volume
  g_tvol = Application.WorksheetFunction.Max(Range("L:L"))
  Cells(4, 17).Value = g_tvol

  
  Next i
  

End Sub


Sub percent_format()

    Range("K:K").NumberFormat = "0.00%"
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"

End Sub

Sub autofitting()

Columns("A:Q").EntireColumn.AutoFit

End Sub
