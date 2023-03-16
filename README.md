# UTModuleVBA
テスト用モジュール。18日に整理。

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


' 用途 : VLOOKUPなどの参照設定が適切にされているか確認するためのマクロ。
' 前提 : 二次元表の各列が参照先として設定されていて、列見出しが存在することが前提。
' 機能 : 指定された列の、列見出しの1行下に、「列見出し+1」を埋め込み、
'           その後オートフィルをかける。
Sub ReferenceCheck_Ver1()
    
    Dim lastCol As Long: lastCol = 100
    Dim lastRow As Long: lastRow = 20
    
    ' 列見出し+1を15列目から最終行まで埋め込む。
    For col = 15 To lastCol
        Cells(2, col).Value = Cells(1, col).Value + "1"
    Next
    
    ' 15列目から最終行までオートフィルをかける。
    Range("O2:CV2").AutoFill _
    Destination:=Range(Cells(2, 15), Cells(20, lastCol))
    
End Sub
' 用途 : VLOOKUPなどの参照設定が適切にされているか確認するためのマクロ。
' 前提 : 二次元表の各列が参照先として設定されていて、列見出しが存在することが前提。
' 機能 : 指定された列の、列見出しの1行下に、「列見出し+1」を埋め込み、
'           その後オートフィルをかける。
Sub ReferenceCheck_Ver2()
    Dim colName() As Variant
    Dim colArray(9) As Long
    Dim col As Long
'仮データ
    colName = Array("名前", "生年月日", "住所")
    For i = 0 To 3
        col = Range("A1:C1").Find(colName(i)).Column
        Cells(2, col).Value = Cells(1, col).Value + "1"
    Next
    
    Dim lastCol As Long: lastCol = 100
    Dim lastRow As Long: lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' 連番になるようオートフィル
    Range("A1:C1").AutoFill _
    Destination:=Range(Cells(2, 15), Cells(lastRow, lastCol))
      
End Sub
