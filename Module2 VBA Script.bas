Attribute VB_Name = "Module1"
Sub stocks()

Dim ws As Worksheet
For Each ws In Worksheets

Dim column As Double
column = 1

Dim row As Double
row = 2

Dim total_volume As Double
total_volume = 0

'close price is the last row for each ticker
Dim close_price As Double


'open price row number, first instance is at line 2
Dim open_row As Double
open_row = 2

Dim open_price As Double

Dim change As Double

Dim perc_change As Double



'outerloop through worksheets

'add heading to each sheet
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'to replace hardcoded row #; end row will differ for each tab
RowCount = ws.Cells(Rows.Count, "A").End(xlUp).row

    For i = 2 To RowCount

        'within the loop distinguish between changing ticker symbol
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
        
                'this puts ticker symbol in column J
                    ws.Cells(row, 10) = ws.Cells(i, column).Value
                    
                'sum total volume
                    total_volume = total_volume + ws.Cells(i, 7).Value
                    
                'this puts total volume in column M
                    ws.Cells(row, 13) = total_volume
                    
                    'close price
                    close_price = ws.Cells(i, 6).Value
                                        
                               
                    'open_price
                    open_price = ws.Cells(open_row, 3).Value
                    
                    
                    'Yearly Change is close - open price
                    change = close_price - open_price
                    ws.Cells(row, 11) = change
                                        
                    'Color positive change green etc
                        If change > 0 Then
                        ws.Cells(row, 11).Interior.ColorIndex = 4
                        ElseIf change < 0 Then
                        ws.Cells(row, 11).Interior.ColorIndex = 3
                        End If
                    
                    '% change is chage/open price
                    perc_change = change / open_price
                    ws.Cells(row, 12) = perc_change
                    ws.Cells(row, 12).NumberFormat = "0.00%"
                    
                    'conditional formatting
                        If perc_change > 0 Then
                        ws.Cells(row, 12).Interior.ColorIndex = 4
                        ElseIf perc_change < 0 Then
                        ws.Cells(row, 12).Interior.ColorIndex = 3
                        End If
                    
                're-initializing variables, no need to reinitialize open or close price since they are not dynamic
                    row = row + 1
                    total_volume = 0
                    open_row = i + 1
                    
        'if ticker is same we do not enter the loop, keep a tally of total volume
        Else
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        
        End If

    Next i
    
'Greatest and Least Change Summary
ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("L2:L" & RowCount))
ws.Range("Q3").NumberFormat = "0.00%"

ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("M2:M" & RowCount))

'ticker label for greatest and least % summary
increase_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
ws.Range("P2") = ws.Cells(increase_index + 1, 10)

increase_index2 = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
ws.Range("P3") = ws.Cells(increase_index2 + 1, 10)

increase_index3 = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("M2:M" & RowCount)), ws.Range("M2:M" & RowCount), 0)
ws.Range("P4") = ws.Cells(increase_index3 + 1, 10)

Next ws
End Sub
