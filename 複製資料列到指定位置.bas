Sub 複製資料列到指定位置()
    Range("A1:A8").Select   ' 選擇資料
    Selection.Copy          ' 複製資料內容
    Range("E1").Select      ' 選擇要貼上資料的位置
    ActiveSheet.Paste       ' 貼上資料
End Sub
