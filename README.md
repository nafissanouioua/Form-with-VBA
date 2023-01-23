# Form-with-VBA
My frist form with VBA
Option Explicit

Private Sub CommandButton5_Click()
Sheets("form").Activate

End Sub


Sub afficher_formulaire()

UserForm1.Show

End Sub


Private Sub CommandButton6_Click()
 TextBox1 = ""
TextBox3 = ""
TextBox4 = ""
TextBox5 = ""
ComboBox1 = ""
 
End Sub

Private Sub CommandButton7_Click()
Unload Me
End Sub


Private Sub CommandButton4_Click()

Sheets("form").Activate


    Range("A1").Select
    Selection.End(xlDown).Select
     Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    
    ActiveCell = TextBox1
    Selection.Offset(0, 1).Value = TextBox3
      Selection.Offset(0, 2).Value = TextBox4
    Selection.Offset(0, 3).Value = TextBox5
      Selection.Offset(0, 4).Value = ComboBox1
      
      MsgBox "Votre élement a été enregistré"
      
    
End Sub


