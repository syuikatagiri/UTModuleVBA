Attribute VB_Name = "selectA1AllSheets"
Sub selectA1AllSheets()
    Dim objSheets As Sheets
    Dim objSheet As Object
    
    Set objSheets = ActiveWorkbook.Worksheets
    
    For Each objSheet In objSheets
        objSheet.Activate
        objSheet.Range("A1").Select
    Next
 
    Worksheets(1).Select
End Sub
