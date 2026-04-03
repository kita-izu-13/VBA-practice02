Attribute VB_Name = "Level17"
Sub Level17_Dictionary()

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long
    Dim i As Long
    Dim key As String
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
    
    For i = 1 To lastRow
    
        key = Cells(i, 1).Value
        
        If dict.Exists(key) Then
            dict(key) = dict(key) + 1
        Else
            dict.Add key, 1
        End If
        
    Next i
    
    'èoóÕ
    Dim row As Long
    row = 1
    
    Dim k As Variant
    
    For Each k In dict.keys
        Cells(row, 3).Value = k
        Cells(row, 4).Value = dict(k)
        row = row + 1
    Next k

End Sub
