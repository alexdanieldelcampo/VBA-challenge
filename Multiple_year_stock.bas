Attribute VB_Name = "Module1"
Sub Stocks()

    Dim ws As Worksheet
    
  For Each ws In Worksheets

    Dim ticke_name As String
    Dim total_sv As LongLong
    Dim op As Double
    Dim cp As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim Summary_Table_Row As Integer
    Dim row_counter As Integer
    
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    row_counter = 0
    Summary_Table_Row = 2
    total_sv = 0
    
    For i = 2 To lastrow
    
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        total_sv = total_sv + ws.Cells(i, 7).Value
        ws.Range("L" & Summary_Table_Row).Value = total_sv
        
        ticker_name = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_Table_Row).Value = ticker_name
        
        op = ws.Cells(i - row_counter, 3).Value
        cp = ws.Cells(i, 6).Value
        year_change = cp - op
        ws.Range("J" & Summary_Table_Row).Value = year_change
        
            If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            
            If op <> 0 Then
            
             percent_change = ((cp - op) / op)
             ws.Range("K" & Summary_Table_Row).Value = percent_change
             ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
             
            Else
            
            ws.Range("K" & Summary_Table_Row).Value = "Null"
            ws.Range("K" & Summary_Table_Row).HorizontalAlignment = xlCenter
            End If
            
        Summary_Table_Row = Summary_Table_Row + 1
        total_sv = 0
        row_counter = 0
        
    Else
    
        row_counter = row_counter + 1
        total_sv = total_sv + ws.Cells(i, 7).Value
    
    End If
    
 Next i
     ws.Columns("I:L").AutoFit
     'BONUS
     ws.Cells(2, 15).Value = "Greatest % Increase"
     ws.Cells(3, 15).Value = "Greatest % Decrease"
     ws.Cells(4, 15).Value = "Greatest Total Volume"
     ws.Cells(1, 16).Value = "Ticker"
     ws.Cells(1, 17).Value = "Value"
     Dim min As Double
     Dim max As Double
     Dim GTV As LongLong
     
        max = ws.Cells(3, 11).Value
        min = ws.Cells(3, 11).Value
        GTV = ws.Cells(3, 12).Value
     For i = 3 To (Summary_Table_Row - 1)
     
        If ws.Cells(i - 1, 11).Value <> "Null" Then
        
            If max < ws.Cells(i - 1, 11).Value Then
                
                max = ws.Cells(i - 1, 11).Value
                
            End If
             
            If min > ws.Cells(i - 1, 11).Value Then
            
                min = ws.Cells(i - 1, 11).Value
            
            End If
            
            If GTV < ws.Cells(i - 1, 12).Value Then
                
                GTV = ws.Cells(i - 1, 12).Value
                
            End If
        End If
        
        Next i
        
        For i = 2 To Summary_Table_Row
            
                If ws.Cells(i, 11).Value = max Then
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                      ws.Cells(2, 17).Value = max
                End If
                
                If ws.Cells(i, 11).Value = min Then
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(3, 17).Value = min
                End If
                
                If ws.Cells(i, 12).Value = GTV Then
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(4, 17).Value = GTV
                End If
                
        Next i
            ws.Columns("O:Q").AutoFit
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
Next ws


End Sub

