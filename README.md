# VBA-challenge
Mod 2 challenge
Sub MYSD():
'variables
Dim Ticker As String
Dim Yearly_Change As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Total As Double
Dim Percent_Change As Double
Dim Max_Ticker As String
Dim Max_Volume As Double
    Total = 0
    Max_Volume = 0
' Loop in all Worksheetts
For Each ws In Worksheets
' Last Row Formula
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Summary_Row = 2
' Column Labels 1
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
' Begin Part 1 For Loop for Ticker, Open, Close, Change
        For I = 2 To LastRow
            If ws.Cells(I + 1, 1) <> ws.Cells(I, 1) Then
                Ticker = ws.Cells(I, 1)
                Total_Volume = Total_Volume + ws.Cells(I, 7)
                    ws.Cells(Summary_Row, 9) = Ticker
                    ws.Cells(Summary_Row, 12) = Total_Volume
                Total_Volume = 0
                Ticker = ""
                Year_Close = ws.Cells(I, 6)
                Yearly_Change = Year_Close - Year_Open
                            ws.Cells(Summary_Row, 10) = Yearly_Change
                Percent_Change = (Yearly_Change / Year_Open)
                            ws.Cells(Summary_Row, 11) = Percent_Change
                Summary_Row = Summary_Row + 1
            Else
                If ws.Cells(I - 1, 1) <> ws.Cells(I, 1) Then
                Year_Open = ws.Cells(I, 3)
                End If
                Total_Volume = Total_Volume + ws.Cells(I, 7)
            End If
        Next I
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
' Begin Part 2 Data Analysis For Loop for Max, Increase, Decrease
        For I = 2 To LastRow2
            If ws.Cells(I, 12) > Max_Volume Then
                Max_Volume = ws.Cells(I, 12)
                Max_Ticker = ws.Cells(I, 9)
            End If
            If ws.Cells(I, 11) > Greatest_Increase Then
                Greatest_Increase = ws.Cells(I, 11)
                Greatest_Increase_Ticker = ws.Cells(I, 9)
            End If
            If ws.Cells(I, 11) < Greatest_Decrease Then
                Greatest_Decrease = ws.Cells(I, 11)
                Greatest_Decrease_Ticker = ws.Cells(I, 9)
            End If
        Next I
' Column Labels 2
    ws.Range("Q4") = Max_Volume
    ws.Range("P4") = Max_Ticker
    ws.Range("Q2") = Greatest_Increase
    ws.Range("P2") = Greatest_Increase_Ticker
    ws.Range("Q3") = Greatest_Decrease
    ws.Range("P3") = Greatest_Decrease_Ticker
' Reset variables for next ws
    Ticker_Row = 2
    Max_Volume = 0
    Max_Ticker = ""
' Format cells
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("I:L").Columns.AutoFit
    ws.Range("O:Q").Columns.AutoFit
        For I = 2 To LastRow2
            If ws.Cells(I, 10) < 0 Then
            ws.Cells(I, 10).Interior.ColorIndex = 3
            Else
            ws.Cells(I, 10).Interior.ColorIndex = 4
            End If
        Next I
Next ws
End Sub

<!-- I worked with Charlotte  Van Dyck, Alex Nguyen,  and Adam Change to complete this challenge
We worked together to build functional code
I also met with Zabur from  BCS Tutoring to make sure my code was not only fucntional, but simplified -->