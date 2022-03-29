
Private Sub UserForm_Initialize()
ComboBox1.AddItem ""
ComboBox1.AddItem "homme"
ComboBox1.AddItem "femme"
ComboBox1.AddItem "Autre"
error.Caption = ""
End Sub

Private Sub caseclear_Click()
Dim A As Long
A = Cells(Rows.Count, 1).End(xlUp).Row
Cells(A, 1).Value = Clear
Cells(A, 2).Value = Clear
Cells(A, 3).Value = Clear
Cells(A, 4).Value = Clear
Cells(A, 5).Value = Clear
Cells(A, 6).Value = Clear
End Sub

Private Sub efface_Click()
TextBox1 = Clear
TextBox2 = Clear
TextBox3 = Clear
TextBox4 = Clear
TextBox5 = Clear
ComboBox1 = Clear
End Sub


Private Sub Suivant_Click()
UserForm1.Hide
UserForm2.Show
End Sub

Private Sub Valider_Click()
Dim A As Long
A = Cells(Rows.Count, 1).End(xlUp).Row + 1
If UserForm1.ComboBox1 = "" Then
    error.Caption = "Error : Le genre de l'utilisateur est vide."
    End If
If UserForm1.TextBox5 = "" Then
    error.Caption = "Error : Le cours de l'utilisateur est vide."
    End If
If UserForm1.TextBox4 = "" Then
    error.Caption = "Error : Le prof de l'utilisateur est vide."
    End If
If UserForm1.TextBox3 = "" Then
    error.Caption = "Error : Le classe de l'utilisateur est vide."
    End If
If UserForm1.TextBox2 = "" Then
    error.Caption = "Error : Le prenom de l'utilisateur est vide."
    End If
If UserForm1.TextBox1 = "" Then
    error.Caption = "Error : Le nom de l'utilisateur est vide."
    End If
Cells(A, 1).Value = TextBox1.Value
Cells(A, 2).Value = TextBox2.Value
Cells(A, 3).Value = TextBox3.Value
Cells(A, 4).Value = TextBox4.Value
Cells(A, 5).Value = TextBox5.Value
Cells(A, 6).Value = ComboBox1.Value
TextBox1 = Clear
TextBox2 = Clear
TextBox3 = Clear
TextBox4 = Clear
TextBox5 = Clear
ComboBox1 = Clear
End Sub
