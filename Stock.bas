Attribute VB_Name = "Stock"
Sub Stock():

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
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Range("H1").Value = "Ticker"
    Range("I1").Value = "Yearly Change"
    Range("J1").Value = "Percent Change"
    Range("K1").Value = "Total Stock Volume"


    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'ticker
            ticker = Cells(i, 1).Value
            Range("H" & NewRow).Value = ticker
        
            'Yearly Change
            BeginingPrice = Cells(Start, 3).Value
            EndingPrice = Cells(i, 6).Value
            YearlyChange = EndingPrice - BeginingPrice
            Start = i + 1
            Range("I" & NewRow).Value = YearlyChange
        
            'PercentageChange
            PercentageChange = YearlyChange / BeginingPrice
            Range("J" & NewRow).Value = PercentageChange
        
            'Total Volume
            TotalVolume = TotalVolume + Cells(i, 7).Value
            Range("k" & NewRow).Value = TotalVolume
    
    
            NewRow = NewRow + 1
        
    
        Else
            YearlyChange = EndingPrice - BeginingPrice
            TotalVolume = TotalVolume + Cells(i, 7).Value

        End If
        
        
            'Format the Color
        If Cells(i, 9).Value > 0 Then
            Cells(i, 9).Interior.ColorIndex = 4
        Else
            Cells(i, 9).Interior.ColorIndex = 3
        End If
    
    

    Next i

    Next ws
    
End Sub



