Attribute VB_Name = "Module1"
Sub Stocks()
    
    
    Dim ticke_name As String
    Dim total_sv As LongLong
    Dim op As Double
    Dim cp As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim Summary_Table_Row As Integer
    Dim row_counter As Integer
    
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    row_counter = 0
    Summary_Table_Row = 2
    total_sv = 0
    
    For i = 2 To lastrow
     
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        total_sv = total_sv + Cells(i, 7).Value
        Range("L" & Summary_Table_Row).Value = total_sv
        
        ticker_name = Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = ticker_name
        
        op = Cells(i - row_counter, 3).Value
        cp = Cells(i, 6).Value
        year_change = cp - op
        Range("J" & Summary_Table_Row).Value = year_change
        
            If Range("J" & Summary_Table_Row).Value < 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            
        percent_change = ((cp - op) / op)
        Range("K" & Summary_Table_Row).Value = percent_change
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        Summary_Table_Row = Summary_Table_Row + 1
        total_sv = 0
        row_counter = 0
        
    Else
    
        row_counter = row_counter + 1
        total_sv = total_sv + Cells(i, 7).Value
    
    End If
    
 Next i
 
 Columns("I:L").AutoFit
    
End Sub
