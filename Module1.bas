Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("2014").Select
    Call Ticker
    Sheets("2015").Select
    Call Ticker
    Sheets("2016").Select
    Call Ticker
End Sub

Sub Ticker()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Dim j As Integer
Dim Total As Double
Dim opening As Double
Dim closing As Double
Dim yearly_change As Double

j = 1
Total = 0
opening = Cells(2, 3).Value
closing = 0
yearly_change = 0

For i = 2 To 800000
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        j = j + 1
        Cells(j, 9).Value = Cells(i, 1).Value
        closing = Cells(i, 6).Value
        yearly_change = closing - opening
        Cells(j, 10).Value = yearly_change
        If Cells(j, 10).Value > 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        Else
            Cells(j, 10).Interior.ColorIndex = 3
        End If
        If opening <> 0 Then
        Cells(j, 11).Value = yearly_change / opening
        End If
        Cells(j, 11).NumberFormat = "0.00%"
        Cells(j, 12).Value = Total + Cells(i, 7).Value
    
        Total = 0
        opening = Cells(i + 1, 3).Value
    Else
        Total = Total + Cells(i, 7).Value
    End If

    Next i

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Dim max_per_ticker As String
Dim max_per As Double
Dim min_per_ticker As String
Dim min_per As Double
Dim max_vol_ticker As String
Dim max_vol As Double

max_per_ticker = Cells(2, 9).Value
max_per = Cells(2, 11).Value
min_per_ticker = Cells(2, 9).Value
min_per = Cells(2, 11).Value
max_vol_ticker = Cells(2, 9).Value
max_vol = Cells(2, 12).Value

For i = 2 To 6000
    If max_per < Cells(i, 11).Value Then
        max_per_ticker = Cells(i, 9).Value
        max_per = Cells(i, 11).Value
    End If
    If min_per > Cells(i, 11).Value Then
        min_per_ticker = Cells(i, 9).Value
        min_per = Cells(i, 11).Value
    End If
        If max_vol < Cells(i, 12).Value Then
        max_vol_ticker = Cells(i, 9).Value
        max_vol = Cells(i, 12).Value
    End If
    
    Next i

Cells(2, 16).Value = max_per_ticker
Cells(2, 17).Value = max_per
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 16).Value = min_per_ticker
Cells(3, 17).Value = min_per
Cells(3, 17).NumberFormat = "0.00%"
Cells(4, 16).Value = max_vol_ticker
Cells(4, 17).Value = max_vol
    
End Sub
