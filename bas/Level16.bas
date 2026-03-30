Attribute VB_Name = "Level16"
Sub Level16()

    Dim lastRow As Long
    Dim outRow As Long
    Dim i As Long
    Dim j As Long
    Dim isExist As Boolean
    Dim arr As Variant
    Dim resultArr() As Variant
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    arr = Range("A1:A" & lastRow).Value '入力配列
    ReDim resultArr(1 To lastRow, 1 To 2)   '結果配列（最大サイズで確保）
    
    outRow = 1
    
    For i = 1 To lastRow
    
        isExist = False
    
        For j = 1 To outRow - 1  'C列にすでにあるかチェック
        
            If arr(i, 1) = resultArr(j, 1) Then
            
                isExist = True
                resultArr(j, 2) = resultArr(j, 2) + 1
                Exit For
            
            End If
        
        Next j
            
        If Not isExist Then   'なければ追加

            resultArr(outRow, 1) = arr(i, 1)
            resultArr(outRow, 2) = 1
            outRow = outRow + 1
        
        End If
    
    Next i
    
    Range("C1").Resize(outRow - 1, 2).Value = resultArr '一括出力

End Sub
