Attribute VB_Name = "Level18_1DictionarySum"
Sub Level18_DictionarySum()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, lastRow As Long, key As Variant
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        key = Cells(i, 1).Value
        If dict.Exists(key) Then
            dict(key) = dict(key) + Cells(i, 2).Value
        Else
            dict.Add key, Cells(i, 2).Value
        End If
    Next i
    
    Dim r As Long
    r = 2
    For Each key In dict.keys
        Cells(r, 4).Value = key
        Cells(r, 5).Value = dict(key)
        r = r + 1
    Next key

End Sub
