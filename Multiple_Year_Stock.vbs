Sub totalstockvolume()

Dim ticker As String
Dim volume As Long
Dim year_open As Double
Dim year_close As Double
Dim totalvolume As Double
Dim yearly_change As Double
Dim Percent_change As Double


'Declaring worksheet as ws

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly change"
ws.Range("K1").Value = "Percent_change"
ws.Range("L1").Value = "Totalvolume"
 
' find last   row

Dim lastrow As Long
lastrow = cells(Rows.Count, 1).End(xlUp).Row
 
 'Initial value of Total is 0
 
totalvolume = 0
'result row initial value is 2 .

Dim resultRow As Integer
resultRow = 2

 
For i = 2 To lastrow

  If ws.cells(i, 1).Value <> ws.cells(i - 1, 1).Value Then

                year_open = ws.cells(i, 3).Value
                End If
volume = cells(i, 7).Value

If ws.cells(i + 1, 1).Value <> ws.cells(i, 1).Value Then
ticker = ws.cells(i, 1).Value

'adding each volume for ticker symbol
totalvolume = totalvolume + volume
ws.Range("I" & resultRow).Value = ticker
ws.Range("L" & resultRow).Value = totalvolume




'Grab year end price
                year_close = ws.cells(i, 6).Value
                
 yearly_change = year_close - year_open
 ws.Range("J" & resultRow).Value = yearly_change
 ws.Range("J" & resultRow).NumberFormat = "0.000000000"
 
            
                
 If ws.Range("J" & resultRow).Value >= 0 Then
 
                    ws.Range("J" & resultRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & resultRow).Interior.ColorIndex = 3
                End If
                
                
                 'Calculate the percent change
                'Conditional for calculating percent change
                If year_open = 0 And year_close = 0 Then
                 
                    Percent_change = 0
                   ws.Range("K" & resultRow).Value = Percent_change
                    ws.Range("K" & resultRow).NumberFormat = "0.00%"
                    
                ElseIf year_open = 0 Then
                Percent_change = 0
                
                      ws.Range("K" & resultRow).Value = Percent_change
                      
                Else
                
                    Percent_change = yearly_change / year_open
                    ws.Range("K" & resultRow).Value = Percent_change
                    ws.Range("K" & resultRow).NumberFormat = "0.00%"
                    
                End If
                
resultRow = resultRow + 1

                  totalvolume = 0
                year_open = 0
                year_close = 0
                yearly_change = 0
                Percent_change = 0

Else


totalvolume = totalvolume + volume
 
End If
Next i

 'Create a separate table for greatest increase and greatest decrease in percentage and greatest stock volume

        'Titles
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        lastrow = ws.cells(Rows.Count, 9).End(xlUp).Row

         Dim increase_stock As String
        Dim increase_value As Double


        'Set  increase stock percentage value to first value of the column

         increase_value = ws.cells(2, 11).Value

        Dim decrease_stock As String
        Dim decrease_value As Double

        'Set decrease  stock value to the first stock

        decrease_value = ws.cells(2, 11).Value

        Dim highstock_vol_stock  As String

        Dim highstock_vol_value As Double

        'Set highest  volume equal to the first cell value
        highstock_vol_value = ws.cells(2, 12).Value

        'Loop to search through Result table
        For j = 2 To lastrow

            'Conditional to determine increase value
            If ws.cells(j, 11).Value > increase_value Then
                increase_value = ws.cells(j, 11).Value
                increase_stock = ws.cells(j, 9).Value
            End If

            'Conditional to determine decrease value
            If ws.cells(j, 11).Value < decrease_value Then
                decrease_value = ws.cells(j, 11).Value
                decrease_stock = ws.cells(j, 9).Value
            End If

            'Conditional to determine stock with the greatest volume

            If ws.cells(j, 12).Value > highstock_vol_value Then
                highstock_vol_value = ws.cells(j, 12).Value
                highstock_vol_stock = ws.cells(j, 9).Value
            End If

        Next j

        'Move increase, decrease and stock with the most volume items to the result table

        ws.cells(2, 16).Value = increase_stock
        ws.cells(2, 17).Value = increase_value
        ws.cells(2, 17).NumberFormat = "0.00%"
        ws.cells(3, 16).Value = decrease_stock
        ws.cells(3, 17).Value = decrease_value
        ws.cells(3, 17).NumberFormat = "0.00%"
        ws.cells(4, 16).Value = highstock_vol_stock
        ws.cells(4, 17).Value = highstock_vol_value

        'Autofit table columns
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit

    Next ws

End Sub
