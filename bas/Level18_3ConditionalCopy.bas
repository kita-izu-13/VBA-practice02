Attribute VB_Name = "Level18_3ConditionalCopy"
Sub Level18_ConditionalCopy()
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Set wsSrc = Worksheets("18Data")
    Set wsDst = Worksheets("18Filtered")
    
    Dim dataArr As Variant
    Dim rowDst As Long
    rowDst = 1
    
    dataArr = wsSrc.Range("A2:C" & wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row).Value
    
    Dim i As Long
    For i = 1 To UBound(dataArr, 1)
        If dataArr(i, 3) > 100 Then
            wsDst.Cells(rowDst, 1).Resize(1, 3).Value = Application.Index(dataArr, i, 0)
            rowDst = rowDst + 1
        End If
    Next i

End Sub
