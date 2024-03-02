Sub VBA_challenge():

Dim ws As Worksheet
Dim WorksheetName As String
WorksheetName = "Multi_year_stock_data.xlsm"

For Each ws In Worksheets

Dim ticker_sym As String
Dim stock_num As Integer
Dim LR As Long
Dim I As Long
Dim volume_total As Double
Dim open_value As Double
Dim close_value As Double
Dim price_change As Double
Dim percent_change As Double
Dim output_row As Integer
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double
Dim greatest_i_ticker As String
Dim greatest_d_ticker As String
Dim greatest_v_ticker As String
 

LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Columns("K:K").NumberFormat = "0.00%"
ws.Range("L1").Value = "Total Volume"

ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("P3").NumberFormat = "0.00%"
ws.Range("N4").Value = "Greatest Total Volume"

ticker_sym = 0
stock_num = 0
volume_total = 0
open_value = 0
close_value = 0
output_row = 2


'Part1

For I = 2 To LR

    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1) Then
     
        ' Set the ticker_sym
        ticker_sym = ws.Cells(I, 1).Value

        ' Add to the volume_total
        volume_total = volume_total + ws.Cells(I, 7).Value
      
        'Set close_value
        close_value = ws.Cells(I, 6).Value

        'Print the ticker_sym in the output_row
        ws.Range("I" & output_row).Value = ticker_sym
      
        'Open_value, was set on first "else" instance
      
        'Calculate price_change (Yearly Change)
        price_change = close_value - open_value
        
        'Print price_change in output_row
        ws.Range("J" & output_row).Value = price_change
      
        'Calculate percent_change
        percent_change = price_change / open_value
      
        'Print percent_change in output_row
        ws.Range("K" & output_row).Value = percent_change
      
        'Print the volume_tot to the output_row
        ws.Range("L" & output_row).Value = volume_total

        'Add one to the output_row
        output_row = output_row + 1
      
        'Reset the volume_total
        volume_total = 0
      
        'Reset open_value
        open_value = 0
         

    Else

        'Counter adding up the value of the volume while ticker_sym is the same
        volume_total = volume_total + ws.Cells(I, 7).Value
    
            'If statement to only capture the initial value. open value is reset above
            If open_value = 0 Then
    
                open_value = ws.Cells(I, 3).Value
        
            End If
    
End If
Next I

'Color formatting percent change
For I = 2 To LR

    If ws.Cells(I, 10).Value < 0 Then
        ws.Cells(I, 10).Interior.ColorIndex = 3

    ElseIf ws.Cells(I, 10).Value > 0 Then
        ws.Cells(I, 10).Interior.ColorIndex = 4

    End If
Next I

'Determining greatest_volume
greatest_volume = ws.Cells(2, 12)

For I = 2 To LR

    If ws.Cells(I, 12).Value = greatest_volume Then
        greatest_volume = ws.Cells(I, 12).Value
        greatest_v_ticker = ws.Cells(I, 9).Value
        ws.Range("O4").Value = greatest_v_ticker
        ws.Range("P4").Value = greatest_volume
    
    ElseIf ws.Cells(I, 12).Value > greatest_volume Then
        greatest_volume = ws.Cells(I, 12).Value
        greatest_v_ticker = ws.Cells(I, 9).Value
        ws.Range("O4").Value = greatest_v_ticker
        ws.Range("P4").Value = greatest_volume
       
    End If
Next I

'Determining greatest_increase
greatest_increase = ws.Cells(2, 11).Value

For I = 2 To LR

    If ws.Cells(I, 11).Value = greatest_increase Then
        greatest_increase = ws.Cells(I, 11).Value
        greatest_i_ticker = ws.Cells(I, 9).Value
        ws.Range("O2").Value = greatest_i_ticker
        ws.Range("P2").Value = greatest_increase

    ElseIf ws.Cells(I, 11).Value > greatest_increase Then
        greatest_increase = ws.Cells(I, 11).Value
        greatest_i_ticker = ws.Cells(I, 9).Value
        ws.Range("O2").Value = greatest_i_ticker
        ws.Range("P2").Value = greatest_increase
    
    End If
Next I


'Determining greatest_decrease
greatest_decrease = ws.Cells(2, 11).Value

For I = 2 To LR


    If ws.Cells(I, 11).Value = greatest_decrease Then
        greatest_decrease = ws.Cells(I, 11).Value
        greatest_d_ticker = ws.Cells(I, 9).Value
        ws.Range("O3").Value = greatest_d_ticker
        ws.Range("P3").Value = greatest_decrease

    ElseIf ws.Cells(I, 11).Value < greatest_decrease Then
        greatest_decrease = ws.Cells(I, 11).Value
        greatest_d_ticker = ws.Cells(I, 9).Value
        ws.Range("O3").Value = greatest_d_ticker
        ws.Range("P3").Value = greatest_decrease

    End If
Next I

Next ws

End Sub