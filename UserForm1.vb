Private Sub CommandButton1_Click()
Dim BarisSel As Long

Sheets("Sheet2").Activate
BarisSel = Application.WorksheetFunction.CountA(Range("A:A")) + 2
Cells(BarisSel, 1) = TextBox1.Text
Cells(BarisSel, 2) = TextBox2.Text

End Sub

Private Sub CommandButton2_Click()
Unload UserForm1
End Sub


