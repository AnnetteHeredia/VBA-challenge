Attribute VB_Name = "Module1"
Sub Stock_ticker():
    
    'Column and row variables
    Dim TickerColumn As String
    TickerColumn = "I"
    
    Dim Yearly_change As String
    Yearly_change = "J"
    
    Dim Percent_change As String
    Percent_change = "K"
    
    Dim Total_stock_volume As String
    Total_stock_volume = "L"
    
    Dim SummaryRow As Integer
    SummaryRow = 2
    
    Dim RawTickerColumn As Integer
    RawTickerColumn = 1
    
    Dim RawVolumeColumn As Integer
    RawVolumeColumn = 7
    
    Dim SummaryTickerColumn As Integer
    SummaryTickerColumn = 9
    
    Dim SummaryStockColumn As Integer
    SummaryStockColumn = 12
    
    
    Dim YearlyOpenColumn As Integer
    YearlyOpenColumn = 3
    
    Dim YearlyCloseColumn As Integer
    YearlyCloseColumn = 6
    
    Dim YearlyChangeColumn As Integer
    YearlyChangeColumn = 10
    
    Dim PercentChangeColumn As Integer
    PercentChangeColumn = 11
    
    
    
    'Value Variables
    Dim YearlyChange As Double
    YearlyChange = 0
    
    Dim PercentChange As Double
    PercentChange = 0
    
    Dim Ticker_Name As String
    
    Dim Total_Stock As Double
    Total_Stock = 0
    
    Dim FirstOpen As Double
    FirstOpen = Cells(2, YearlyOpenColumn).Value
    
    Range(TickerColumn & "1").Value = "Ticker"
    Range(Yearly_change & "1").Value = "Yearly Change"
    Range(Percent_change & "1").Value = "Percent Change"
    Range(Total_stock_volume & "1").Value = "Total Stock Volume"
 
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row - 1
    
    
    For i = 2 To lastrow
    Total_Stock = Total_Stock + Cells(i, RawVolumeColumn).Value
        If Cells(i + 1, RawTickerColumn).Value <> Cells(i, RawTickerColumn).Value Then
            
            'Cells(i, 8).Value = "Change"
            Ticker_Name = Cells(i, RawTickerColumn).Value
            Cells(SummaryRow, SummaryTickerColumn).Value = Ticker_Name
            
            Cells(SummaryRow, SummaryStockColumn).Value = Total_Stock
            
            YearlyChange = (Cells(i, YearlyCloseColumn).Value - FirstOpen)
            Cells(SummaryRow, YearlyChangeColumn).Value = YearlyChange
            
            PercentChange = (YearlyChange) / (FirstOpen) * 100
            Cells(SummaryRow, PercentChangeColumn).NumberFormat = "0.00%"
            
            Cells(SummaryRow, PercentChangeColumn).Value = PercentChange
           
            
            If Cells(SummaryRow, YearlyChangeColumn).Value >= 1 Then
                Cells(SummaryRow, YearlyChangeColumn).Interior.ColorIndex = 4
            ElseIf Cells(SummaryRow, YearlyChangeColumn).Value <= 1 Then
                Cells(SummaryRow, YearlyChangeColumn).Interior.ColorIndex = 3
            End If
            
            
            SummaryRow = SummaryRow + 1
       
            'Reset for next Company Set of Data
            Total_Stock = 0
            FirstOpen = Cells(i + 1, YearlyOpenColumn).Value
      
      
        End If
            
            
    Next i
    

End Sub

    

