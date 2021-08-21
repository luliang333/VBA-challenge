Attribute VB_Name = "Stock"
Sub Stock():
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    
    
    Dim ticker As String
    Dim NewRow As Integer
    Dim Start As Double
    Dim LastRow As Long
    Dim BeginingPrice As Double
    Dim EndingPrice As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalVolume As Double

    NewRow = 2
    Start = 2
    EndingPrice = 0
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range("H1").Value = "Ticker"
    ws.Range("I1").Value = "Yearly Change"
    ws.Range("J1").Value = "Percent Change"
    ws.Range("K1").Value = "Total Stock Volume"


    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'ticker
            ticker = ws.Cells(i, 1).Value
            ws.Range("H" & NewRow).Value = ticker
        
            'Yearly Change
            BeginingPrice = ws.Cells(Start, 3).Value
            EndingPrice = ws.Cells(i, 6).Value
            YearlyChange = EndingPrice - BeginingPrice
            Start = i + 1
            ws.Range("I" & NewRow).Value = YearlyChange
        
            'PercentageChange
            PercentageChange = YearlyChange / BeginingPrice
            ws.Range("J" & NewRow).Value = PercentageChange
            ws.Range("J" & NewRow).NumberFormat = "0.00%"
            
        
            'Total Volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            ws.Range("k" & NewRow).Value = TotalVolume
    
            NewRow = NewRow + 1
            TotalVolume = 0
        
    
        Else
            YearlyChange = EndingPrice - BeginingPrice
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value

        End If
        
        
            'Format the Color
        If ws.Cells(i, 9).Value > 0 Then
            ws.Cells(i, 9).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 9).Value < 0 Then
            ws.Cells(i, 9).Interior.ColorIndex = 3
        End If
        
        
        
    'Bonus Question
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    
    GreatestIncrease = WorksheetFunction.Max(ws.Range("J:J"))
    GreatestDecrease = WorksheetFunction.Min(ws.Range("J:J"))
    GreatestVolume = WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(2, 16).Value = GreatestIncrease
    ws.Cells(3, 16).Value = GreatestDecrease
    ws.Cells(4, 16).Value = GreatestVolume
    ws.Range("P2:P3").NumberFormat = "0.00%"


   
        If ws.Cells(i, 10).Value = GreatestIncrease Then
        ws.Cells(2, 15).Value = ws.Cells(i, 8)
        
        ElseIf ws.Cells(i, 10).Value = GreatestDecrease Then
        ws.Cells(3, 15).Value = ws.Cells(i, 8)
    
        
        ElseIf ws.Cells(i, 11).Value = GreatestVolume Then
        ws.Cells(4, 15).Value = ws.Cells(i, 8)
        
        End If
    
    
    

    Next i
    Next ws

    
End Sub


