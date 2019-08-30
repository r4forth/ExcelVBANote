Sub FontSet()
    With Worksheets("字型設定").Range("A1:B369").Font
        .Name = "微軟正黑體"
        .Size = 14
        .Bold = True
        .ColorIndex = 3
    End With
End Sub
