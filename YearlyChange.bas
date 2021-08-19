Attribute VB_Name = "YearlyChange"
Sub YearlyChange():


Dim BeginingPrice As Double
Dim EndingPrice As Double
Dim NewRow As Integer
Dim LastRow As Long
Dim YearlyChange As Double

BeginingPrice = Cells(2, 3).Value
EndingPrice = 0
NewRow = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row



For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        BeginingPrice = Cells(2, 3).Value
        EndingPrice = Cells(i, 6).Value
        YearlyChange = EndingPrice - BeginingPrice

 Range("I" & NewRow).Value = YearlyChange
    NewRow = NewRow + 1
    



End If
Next i


End Sub



