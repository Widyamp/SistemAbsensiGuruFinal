Private Sub Absensi_Click()
UserForm2.Show
End Sub

Private Sub CommandButton1_Click()
UserName = InputBox("Password:")
If UserName <> "Widya" Then GoTo Salah
MsgBox "Login Berhasil. Selamat Datang Widya"
Exit Sub
Salah:
MsgBox "Maaf. Password yang Anda Masukan Salah!"
End Sub


