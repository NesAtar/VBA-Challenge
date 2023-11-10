Sub stockanalysis()

Dim ticker As String
Dim next_ticker As String
Dim lastrow As Long
Dim companies As Integer
Dim yearly_change As Double
Dim close_price As Double
Dim open_price As Double
Dim total_stock_volume As Double
Dim percent_change As Double
Dim Greatest_number As Double
Dim Greatest_number_ticker As String
Dim Minimum_number As Double
Dim Minimum_number_ticker As String
Dim Greatest_volume As Double
Dim Greatest_volume_ticker As String


companies = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
open_price = Cells(2, 3).Value
total_stock_volume = 0
Greatest_number = 0
Mimum_number = 0
Greatest_volume = 0

For i = 2 To lastrow

ticker = Cells(i, 1).Value
next_ticker = Cells(i + 1, 1).Value

    'if ticker changes Then
    If ticker <> next_ticker Then
        
        'yearly price & percentage change from the open-close price
        close_price = Cells(i, 6).Value
        yearly_change = close_price - open_price
        
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        
        Cells(companies, 9).Value = ticker
        Cells(companies, 10).Value = yearly_change
        
        
            If yearly_change < 0 Then
                Cells(companies, 10).Interior.Color = RGB(255, 0, 0)
                
                Else
                Cells(companies, 10).Interior.Color = RGB(0, 255, 0)
                
            End If
        
        
        
        percent_change = yearly_change / open_price
        Cells(companies, 11).Value = percent_change
        Cells(companies, 11).NumberFormat = "0.00%"
            
            If percent_change > Greatest_number Then
                Greatest_number = percent_change
                Greatest_number_ticker = ticker
            End If
            
            If percent_change < Minimum_number Then
                Minimum_number = percent_change
                Minimum_number_ticker = ticker
            End If
            
            If Greatest_volume < total_stock_volume Then
                Greatest_volume = total_stock_volume
                Greatest_volume_ticker = ticker
             End If
        
        Cells(companies, 12).Value = total_stock_volume
         
        companies = companies + 1
        open_price = Cells(i + 1, 3).Value
        
        total_stock_volume = 0
        
        Else
       
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
    End If
    
Next i

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Greatest_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
'Minimum_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)


Range("P2") = Greatest_number_ticker
Range("Q2") = Greatest_number
'Cells(5, 1) = Format(0.56324, "Percent")
Cells(2, 17).NumberFormat = "0.00%"

'Cells(companies, 11).NumberFormat = "0.00%"

Range("P3") = Minimum_number_ticker
Range("Q3") = Minimum_number
Cells(3, 17).NumberFormat = "0.00%"

Range("P4") = Greatest_volume_ticker
Range("Q4") = Greatest_volume

'Range("P3") = Cells(Minimum_number + 1, 9)

'Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
'Range("Q3") = "%" & WorkksheetFunction.Min(Range("K2:K" & lastrow)) * 100


    
End Sub



