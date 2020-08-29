Sub StockData()
    
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Add Heading for summary
        Range("I1").Select
        ActiveCell.FormulaR1C1 = "Ticker"
        Range("J1").Select
        ActiveCell.FormulaR1C1 = "Yearly Change"
        Range("K1").Select
        ActiveCell.FormulaR1C1 = "Percent Change"
        Range("L1").Select
        ActiveCell.FormulaR1C1 = "Total Stock Volume"
        
        ' Determine the Last Row
        Lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Create Variable to hold Value
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        'Set Initial Open Price
        Open_Price = Cells(2, Column + 2).Value
         
         ' Loop through all ticker symbol
        
        For i = 2 To Lastrow
         
         ' Check if we are still within the same ticker symbol
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                ' Set Ticker name
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
                ' Set Close Price
                Close_Price = Cells(i, Column + 5).Value
                ' Add Yearly Change
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Column + 9).Value = Yearly_Change
                ' Add Percent Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' Add Total Volume
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                ' Add one to the summary table row
                Row = Row + 1
                ' reset the Open Price
                Open_Price = Cells(i + 1, Column + 2)
                ' reset the Volume Total
                Volume = 0
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Determine the Last Row of Yearly Change per WS
        Lastrow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        ' Set the Cell Colors
        For j = 2 To Lastrow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 4
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        ' Set Greatest % Increase, % Decrease, and Total Volume
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        
        ' Look through each rows to find the greatest value and its associate ticker
        For Y = 2 To Lastrow
            If Cells(Y, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & Lastrow)) Then
                Cells(2, Column + 15).Value = Cells(Y, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Y, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Y, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & Lastrow)) Then
                Cells(3, Column + 15).Value = Cells(Y, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Y, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Y, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & Lastrow)) Then
                Cells(4, Column + 15).Value = Cells(Y, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Y, Column + 11).Value
            End If
        Next Y
        
    Next WS
        
End Sub
