Attribute VB_Name = "Level19"
Sub Level19()

    'á@îzóÒÇ…ÉfÅ[É^Çì«Ç›çûÇﬁ
    Dim dataArr As Variant
    Dim lastRow As Long
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
    dataArr = Range("A2:B" & lastRow).Value
    
    'áADictionaryÇèÄîı
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    'áBÉãÅ[Évèàóù
    Dim i As Long
    Dim key As Variant
    
    For i = 1 To UBound(dataArr, 1)
    
        key = dataArr(i, 1)
        
        'áCèåè
        If dataArr(i, 2) >= 100 Then
        
            If dict.Exists(key) Then
                dict(key) = dict(key) + dataArr(i, 2)
            Else
                dict.Add key, dataArr(i, 2)
            End If
        
        End If
        
    Next i
    
    'áDåãâ èoóÕ
    Dim r As Long
    r = 2
    
    For Each key In dict.Keys
        Cells(r, 4).Value = key
        Cells(r, 5).Value = dict(key)
        r = r + 1
    Next key
        
End Sub
