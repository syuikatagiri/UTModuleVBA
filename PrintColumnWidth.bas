Attribute VB_Name = "PrintColumnWidth"
Sub PrintColumnWidth()
    Dim r As Range
    Dim iColumnWidth
    Dim i As Long
    Dim iColumnCount
    
    '�@����m�肽������w��
    Range("A:T").Select
    iColumnCount = Selection.Columns.Count
    
    ' �I��͈͂̐擪�񂩂�ŏI��܂Ń��[�v
    For i = 1 To iColumnCount
        ' ��͈͂��擾
        Set r = Selection.Columns(i)
        
        ' �w���̕����擾
        iColumnWidth = r.ColumnWidth
        ' �C�~�f�B�G�C�g�E�B���h�E�ɕ\��
        Debug.Print iColumnWidth
    Next i
End Sub
