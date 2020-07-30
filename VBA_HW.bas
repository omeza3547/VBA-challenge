VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockData()
 
 
  For Each ws In Worksheets

        'Declare Variables and initial value
        Dim TickerName As String
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVol As Double
        TotalStockVol = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        Dim PreviousAmount As Long
        PreviousAmount = 2
         
        
        
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim LastRowValue As Long
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0

        
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            
             TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                
                TickerName = ws.Cells(i, 1).Value
                
                 
                ws.Range("I" & SummaryTableRow).Value = TickerName
                
                
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
                
                ws.Range("L" & SummaryTableRow).Value = TotalStockVol
                
                 TotalStockVol = 0


                If YearlyOpen = 0 Then
                   PercentChange = 0
                   
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                    
                End If
                
                
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

               
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
            Next i
   
   

            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            
            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
            
        
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
 
    Next ws

End Sub
