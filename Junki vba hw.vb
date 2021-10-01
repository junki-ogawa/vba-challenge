Sub Stocks()

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    Debug.Print WS.Name
    
    
    
    WS.Cells(1, "I").Value = "Ticker"
    WS.Cells(1, "J").Value = "Yearly Change"
    WS.Cells(1, "K").Value = "Percent Change"
    WS.Cells(1, "L").Value = "Total Stock Volume"
    WS.Cells(1, "O").Value = "Greatest % Increase"
    WS.Cells(1, "P").Value = "Ticker"
    WS.Cells(1, "Q").Value = "Value"
    
    
    Dim Ticker As String
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Volume As Double
    
    lastRowIndex = WS.Cells(WS.Rows.Count, 1).End(xlUp).Row
    
    Volume = 0
        
    Dim Summary_Table As Double
    Summary_Table = 2
    
    Dim i As Long
    Dim Start As Long
       
    Start = 2
        
    For i = 2 To lastRowIndex
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
        Ticker = WS.Cells(i, 1).Value
        WS.Cells(Summary_Table, 9).Value = Ticker

        If WS.Cells(Start, 3).Value = 0 Then
            For FindValue = Start To i
                If WS.Cells(FindValue, 3).Value <> 0 Then
                    Start = FindValue
                    Exit For
                End If
            Next FindValue
        End If
        
        Year_Close = WS.Cells(i, 6).Value
        
        Yearly_Change = WS.Cells(i, 6).Value - WS.Cells(Start, 3).Value
        WS.Cells(Summary_Table, "J").Value = Yearly_Change
        
        Percent_Change = Yearly_Change / WS.Cells(Start, 3).Value
        WS.Cells(Summary_Table, "K").Value = Percent_Change
        WS.Cells(Summary_Table, "K").NumberFormat = "0.00%"
        
        If WS.Cells(Summary_Table, "J").Value >= 0 Then
            WS.Cells(Summary_Table, "J").Interior.ColorIndex = 10
        Else
            WS.Cells(Summary_Table, "J").Interior.ColorIndex = 3
        End If
        
        Start = i + 1
        
        Volume = Volume + WS.Cells(i, 7).Value
        WS.Cells(Summary_Table, "L").Value = Volume
        Summary_Table = Summary_Table + 1
        
        Yearly_Change = 0
        Volume = 0
    Else
        
        Volume = Volume + WS.Cells(i, 7).Value
                
    End If
    Next i
    Next WS
End Sub





