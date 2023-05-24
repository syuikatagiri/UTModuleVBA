Attribute VB_Name = "CreateNumberHeaders"
Sub CreateNumberHeaders()

    Dim lastCol As Long: lastCol = 15

    ' 列見出しの下に「列見出し+数値」を設定
    For col = 1 To lastCol
        Cells(2, col).Value = Cells(1, col).Value + "1"
    Next

    ' 「列見出し 数値」を下方向にオートフィル
    Range("A2:O2").AutoFill Range("A2:O30"), xlFillSeries

End Sub
