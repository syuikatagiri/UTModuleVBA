Attribute VB_Name = "PrintColumnWidth"
Sub PrintColumnWidth()
    Dim r As Range
    Dim iColumnWidth
    Dim i As Long
    Dim iColumnCount
    
    '　幅を知りたい列を指定
    Range("A:T").Select
    iColumnCount = Selection.Columns.Count
    
    ' 選択範囲の先頭列から最終列までループ
    For i = 1 To iColumnCount
        ' 列範囲を取得
        Set r = Selection.Columns(i)
        
        ' 指定列の幅を取得
        iColumnWidth = r.ColumnWidth
        ' イミディエイトウィンドウに表示
        Debug.Print iColumnWidth
    Next i
End Sub
