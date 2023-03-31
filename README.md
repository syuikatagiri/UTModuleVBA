
Sub CheckDuplicate_calc()

Dim lastCol As Long: lastCol = 100

    For col = 1 To lastCol

        Columns(col).Select
        Selection.FormatConditions.AddUniqueValues
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).DupeUnique = xlDuplicate
        With Selection.FormatConditions(1).Font
            .Color = -16383844
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
    Next
    
End Sub
' 列見出しの一つ下のセルに、
' "列見出し1"という文字列を埋め込んで、
' 最後にオートフィルかけるマクロ。
'
' 最終的に各列のセルに "列見出し+連番の数値" がうめこまれる
Sub ReferenceCheck_Takt()
    
    Dim lastCol As Long: lastCol = 6
    Dim lastRow As Long: lastRow = 1
    
    ' 列見出し+1を設定
    For col = 1 To lastCol
        Cells(3, col).Value = Cells(2, col).Value + "1"
    Next
    
    ' 連番になるようオートフィル
    Range("B3:F3").AutoFill _
    Destination:=Range(Cells(3, 2), Cells(lastRow, lastCol))
    
End Sub



Sub ReferenceCheck_VerSubTakt()

    Dim colName() As Variant
    Dim colArray(9) As Long
    Dim col As Long

    colName = Array("SA_1*開始*", "SA_1*終了*", "SA_2*開始*", "SA_2*終了*", "SA_3*開始*", "SA_3*終了*", "SA_4*開始*", "SA_4*終了*", "U5")
    For i = 0 To 8
        col = Range("A1:CV1").Find(colName(i)).Column
        Cells(2, col).Value = Cells(1, col).Value + "1"
    Next
    
    Dim lastCol As Long: lastCol = 100
    Dim lastRow As Long: lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' 連番になるようオートフィル
    Range("O2:CV2").AutoFill _
    Destination:=Range(Cells(2, 15), Cells(lastRow, lastCol))
      
End Sub
'
' 列幅の確認用関数
'
Sub PrintColWidth()
    Dim r As Range      '// Rangeオブジェクト
    Dim iColumnWidth    '// 列の幅
    Dim i As Long            '// ループカウンタ
    Dim iColumnCount    '// 選択範囲の列数
    
    '// 複数列の取得
    Range("A:T").Select
    iColumnCount = Selection.Columns.Count
    
    '// 選択範囲の先頭列から最終列までループ
    For i = 1 To iColumnCount
        '// 列範囲を取得
        Set r = Selection.Columns(i)
        
        '// 指定列の幅を取得
        iColumnWidth = r.ColumnWidth
        Debug.Print iColumnWidth
    Next i

End Sub

Sub 全シートをA1セルに移動し選択した状態にする()
    Dim objSheets As Sheets
    Dim objSheet As Object 'Sheet以外も入るので汎用的なObject型にしています。

    'アクティブブックの全シートをオブジェクトにセットします。
    Set objSheets = ActiveWorkbook.Worksheets

    For Each objSheet In objSheets
        'Selectメソッドを効かせるためシートをアクティブにします。
        objSheet.Activate
        'A1セルに移動し選択します。
        objSheet.Range("A1").Select
    Next
    
    'オブジェクトを解放します。
    Set objSheets = Nothing
    Set objSheet = Nothing
    
    Worksheets(1).Select
    
End Sub
