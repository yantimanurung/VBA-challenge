VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock()
Row = 2
beg = 2
great_inc = 0
great_dec = 0
great_vol = 0

For i = 2 To 800000
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
        Cells(Row, 9).Value = Cells(i, 1).Value
        
        Cells(Row, 12) = volume + Cells(i, 7).Value
        
        Cells(Row, 10) = Cells(i, 6).Value - Cells(beg, 3).Value
            If Cells(Row, 10).Value > 0 Then
                Cells(Row, 10).Interior.ColorIndex = 4
            Else: Cells(Row, 10).Interior.ColorIndex = 3
            End If
            
        Cells(Row, 11) = Cells(Row, 10) / Cells(beg, 3).Value
        Cells(Row, 11).NumberFormat = "0.00%"
        
        Row = Row + 1
        volume = 0
        beg = i + 1
        
    Else: volume = volume + Cells(i, 7)
    End If

    If Cells(i, 11).Value > great_inc Then
        great_inc = Cells(i, 11).Value
        ticker1 = Cells(i, 9).Value
     End If
    If Cells(i, 11).Value < great_dec Then
        great_dec = Cells(i, 11).Value
        ticker2 = Cells(i, 9).Value
    End If
    If Cells(i, 12).Value > great_vol Then
        great_vol = Cells(i, 12).Value
        ticker3 = Cells(i, 9).Value
    End If
Next i
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "greates % increase"
Cells(3, 14).Value = "greates % decrease"
Cells(4, 14).Value = "greates total volume"
Cells(2, 16).Value = great_inc
Cells(2, 16).NumberFormat = "0.00%"
Cells(3, 16).Value = great_dec
Cells(3, 16).NumberFormat = "0.00%"
Cells(4, 16).Value = great_vol
Cells(2, 15).Value = ticker1
Cells(3, 15).Value = ticker2
Cells(4, 15).Value = ticker3
Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "Yearly change"
Cells(1, 11).Value = "Percentage"
Cells(1, 12).Value = "Total volume"

End Sub
