VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Tickercheck():
    ws_count = 1
    For Each ws In Worksheets
        Worksheets(ws_count).Activate
        
        
        Dim greatinc As Double
        Dim bigdec As Double
        Dim lrgvol As Double
        lrgvol = 0
        greatinc = 0
        bigdec = 0
        ws_count = ws_count + 1
        r = 2
        Dim LL As Long
        LL = Cells(Rows.Count, 1).End(xlUp).Row
        
        
        For i = 2 To LL
            
        
            Cells(r, 9).Value = Cells(i, 1).Value
            vol_count = vol_count + Cells(i, 7).Value
            op_Date = Cells(i, 2).Value
            open_date = Cells(2, 2).Value
            close_date = Cells(2, 2).Value
            If close_date < op_Date Then
            close_date = op_Date
            End If
            If open_date = op_Date Then
                yr_open = Cells(i, 3).Value
            End If
            If close_date = op_Date Then
                yr_close = Cells(i, 3).Value
            End If
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                yr_change = yr_close - yr_open
                pc_change = yr_close / yr_open
                pc_change = pc_change - 1
                If pc_change > greatinc Then
                    greatinc = pc_change
                    Cells(2, 14).Value = Cells(i, 1).Value
                    Cells(2, 15).Value = greatinc
                End If
                If pc_change < bigdec Then
                    bigdec = pc_change
                    Cells(3, 14).Value = Cells(i, 1).Value
                    Cells(3, 15).Value = bigdec
                End If
                    If vol_count > lrgvol Then
                    lrgvol = vol_count
                End If
                
                Cells(4, 14).Value = Cells(i, 1).Value
                Cells(4, 15).Value = lrgvol
                Cells(r, 12).Value = vol_count
                Cells(r, 10).Value = yr_change
                Cells(r, 11).Value = pc_change
                
                If yr_change < 0 Then
                    Cells(r, 10).Interior.ColorIndex = 3
                Else
                    Cells(r, 10).Interior.ColorIndex = 4
                End If
                Cells(r, 11).NumberFormat = "0.00%"
            
                yr_close = 0
                yr_open = 0
                vol_count = 0
            
            r = r + 1
            End If
        Next i
        Cells(3, 15).NumberFormat = "0.00%"
        Cells(2, 15).NumberFormat = "0.00%"
        Cells(4, 15).NumberFormat = "0"
    MsgBox ("Worksheet complete")
    
Next ws
 
    
End Sub

