Attribute VB_Name = "Module1"
Sub Hesapla()
    Dim i As Long
    Dim sonSatir As Long
    
    sonSatir = Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 1 To sonSatir
        If IsNumeric(Cells(i, 2).Value) Then
            Cells(i, 3).Value = Application.WorksheetFunction.Round((Cells(i, 2).Value * 100) / 60, 0)
        End If
    Next i
End Sub

Sub Temizle()
    Columns("B:C").ClearContents
End Sub
