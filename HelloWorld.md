# 第一支 VBA 程式

如何在 Excel 環境中編寫第一支 VBA 程式。

* 在 Excel 視窗中按下 alt + F11 組合鍵，啟用 VBE 視窗。
* 視窗記得要開啟 **專案總管**、**即時運算**視窗(按下 ctrl + G 組合鍵)。
* 在活頁簿中加入一個模組，插入一個程序，叫做**HelloWorld**。
* 在 Excel 中加入一個按鈕，指定對應的程式為**HelloWorld**。 

## 程式碼

```
Public Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub
```
