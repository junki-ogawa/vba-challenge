Sub greatest()
    Dim WS As Worksheet
    Set WS = Worksheets("A")
    Dim Ticker As String

    WS.Cells(1, "P").Value = "Ticker"
    WS.Cells(1, "Q").Value = "Value"
    WS.Cells(2, "O").Value = "Greatest Percent Increase"
    WS.Cells(3, "O").Value = "Greatest Percent Decrease"
    WS.Cells(4, "O").Value = "Greatest Total Volume"
    
    Start = 2
    LastRow = WS.Cells(Rows.Count, "I").End(xlUp).Row
    For i = 2 To LastRow
    
    WS.Cells(2, "Q").Value = Application.WorksheetFunction.Max(WS.Range("K:K"), 1)
    WS.Cells(2, "Q").NumberFormat = "0.00%"
    
    If WS.Cells(2, "Q").Value = WS.Cells(2, "I").Value Then
        WS.Cells(2, "Q").Value = WS.Cells(2, "I").Value
        End If
        Next i
        
    WS.Cells(3, "Q").Value = Application.WorksheetFunction.Min(WS.Range("K:K"), 1)
    WS.Cells(3, "Q").NumberFormat = "0.00%"
    
    WS.Cells(4, "Q").Value = Application.WorksheetFunction.Max(WS.Range("L:L"), 1)
    
End Sub
