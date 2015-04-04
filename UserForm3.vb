Private Sub CommandButton1_Click()
Dim BarisSel As Long

Sheets("Sheet3").Activate
BarisSel = Application.WorksheetFunction.CountA(Range("A:A")) + 2
Cells(BarisSel, 1) = TextBox1.Text
Cells(BarisSel, 2) = TextBox2.Text
Cells(BarisSel, 3) = TextBox3.Text

End Sub

Private Sub CommandButton2_Click()
Unload UserForm3
End Sub


Private Sub CommandButton3_Click()
Me.Hide
Call TampilTanggalWaktu
Unload Me

UserForm3.Show
End Sub
