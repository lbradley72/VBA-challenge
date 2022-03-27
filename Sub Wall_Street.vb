Sub Wall_Street_New()
  
'Create multiple worksheet loop
For Each ws In Worksheets
ws.Activate

'Create worksheet Headers
    
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
 
'Declare variables required for output
Dim ticker As String
Dim open_price As Double
Dim volume As Double
Dim year_total As Double
Dim percent_change As Double
Dim ticker_row As Long
Dim LastDataRow As Long
Dim i As Long

LastDataRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
volume = 0
ticker_row = 2
year_total = 0
  
'Create loop to produce output
For i = 2 To LastDataRow
open_price = ws.Cells(ticker_row, 3).Value

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & ticker_row).Value = ticker
               
        year_total = (-open_price + ws.Cells(i, 6).Value) + year_total
        ws.Range("J" & ticker_row).Value = year_total
    
        percent_change = (year_total / open_price)
        ws.Range("K" & ticker_row).Value = percent_change
        ws.Range("K" & ticker_row).Style = "Percent"
        
        volume = ws.Cells(i, 7).Value + volume
        ws.Range("L" & ticker_row).Value = volume
      
        ticker_row = ticker_row + 1
        year_total = 0
        volume = 0
        open_price = ws.Cells(ticker_row, 3).Value
    Else
        volume = volume + ws.Cells(i, 7).Value
    End If
    
Next i

'Format yearly change column in all spreadsheets

Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    For i = 2 To LastDataRow

    If ws.Cells(i, 10).Value >= 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If

Next i
        
        ws.Range("I1:L1").Columns.AutoFit
    
Next ws


End Sub
