Attribute VB_Name = "Level18_2ArraySum"
Sub Level18_ArraySum()
    Dim dataArr As Variant
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
    dataArr = Range("A2:B5").Value
    'dataArr = Range("A2:B" & lastRow).Value
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, key As Variant
    For i = 1 To UBound(dataArr, 1)
        key = dataArr(i, 1)
        If dict.Exists(key) Then
            dict(key) = dict(key) + dataArr(i, 2)
        Else
            dict.Add key, dataArr(i, 2)
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
