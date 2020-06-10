Sub stocks()
'Loop through all sheets
For Each ws In Worksheets

'Summary Table Ticker Symbol*
Dim ticker As String
Dim yearly_change As Double
Dim percentage_change As Double
Dim total_volume As Double
Dim open_volume As Double

'Count first open stock value
open_volume = ws.Cells(2, 3).Value


'Summary Table Headers
ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total Stock Volume"


Dim summary_table As Integer
summary_table = 2

'Set up lastrow count*
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row



'Loop through all ticker symbols*
For i = 2 To lastrow


'Check if still in same ticker value, if not...*
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'Set Ticker*
    ticker = ws.Cells(i, 1).Value
    
    If open_volume = 0 Then
        percentage_change = 0
    
    
    Else
    
    'Calculate percentage change
    percentage_change = (ws.Cells(i, 6).Value - open_volume) / open_volume
    
    End If

    'yearly change calculation
    yearly_change = ws.Cells(i, 6).Value - open_volume
   
   'Calculate Total volume
    total_volume = total_volume + ws.Cells(i, 7).Value
    
    'Print Ticker symbol, yearly change, percentage change and total volume in Summary table
    ws.Range("I" & summary_table).Value = ticker
    ws.Range("J" & summary_table).Value = yearly_change
    ws.Range("K" & summary_table).Value = percentage_change
    ws.Range("L" & summary_table).Value = total_volume
    
    'Add one to the summary table row
    summary_table = summary_table + 1
    
    'reset total volume
    total_volume = 0
    
    'Reset open volume
    open_volume = ws.Cells(i + 1, 3).Value
    

    
Else
    'Calculate total volume
    total_volume = total_volume + ws.Cells(i, 7).Value
    

End If
Next i

'Formating percentage change to reflect %
For i = 2 To lastrow

  ws.Cells(i, 11).NumberFormat = "0.00%"
  
  Next i

'Conditional format Yearly Change colors
For i = 2 To lastrow
    
    If (ws.Cells(i, 10) > 0) Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf (ws.Cells(i, 10) < 0) Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 0

End If

Next i

'Challenge portion headers
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


Dim pincrease As Double
Dim tickerpinc As String
Dim pdecrease As Double
Dim tickerpdec As String
Dim gvolume As Double
Dim tickergvol As String

'Search for Greatest % Increase, apply to challenge table to format %
pincrease = Application.WorksheetFunction.Max(ws.Range("k:k"))
ws.Cells(2, 17).Value = pincrease
ws.Cells(2, 17).NumberFormat = "0.00%"

'Search for Greatest % Decrease, apply to challenge table to format %
pdecrease = Application.WorksheetFunction.Min(ws.Range("k:k"))
ws.Cells(3, 17).Value = pdecrease
ws.Cells(3, 17).NumberFormat = "0.00%"

'Search for Greatest Total Volume and apply to challenge table
tickergvol = Application.WorksheetFunction.Max(ws.Range("l:l"))
ws.Cells(4, 17).Value = tickergvol


'Calculating challenge ticker
For i = 2 To lastrow

If ws.Cells(i, 11).Value = pincrease Then
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value

ElseIf ws.Cells(i, 11).Value = pdecrease Then
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

ElseIf ws.Cells(i, 12).Value = tickergvol Then
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value

End If
Next i
Next ws

End Sub

