Attribute VB_Name = "CreateNumberHeaders"
Sub CreateNumberHeaders()

Dim lastCol As Long: lastCol = 15

' �񌩏o���̉��Ɂu�񌩏o��+���l�v��ݒ�
For col = 1 To lastCol
    Cells(2, col).Value = Cells(1, col).Value + "1"
Next

' �u�񌩏o�� ���l�v���������ɃI�[�g�t�B��
Range("A2:O2").AutoFill Range("A2:O30"), xlFillSeries

End Sub
