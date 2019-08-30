Sub 計算指定範圍的值是否大於一千的個數()
    Dim mycount As Integer, rng As Range
    For Each rng In Range("A1:B14")
        If rng.Value > 1000 Then
            With rng.Font
                .Color = vbRed
                .Name = "微軟正黑體"
                .Bold = True
            End With
            mycount = mycount + 1
        End If
    Next
    Debug.Print "A:B14 中大於 1000 的資料格個數: " & mycount
End Sub
