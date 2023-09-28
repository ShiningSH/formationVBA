Option Explicit
' Annuler
Private Sub CommandButton1_Click()
    Unload Me
End Sub
' Go
Private Sub CommandButton2_Click()
    Dim duree As String
    duree = TextBox3.Value
    If Not IsNumeric(duree) Then
        MsgBox "La durée doit être un nombre de jours"
        Unload Me
        Exit Sub
    End If
    Call callRemplirDonnees(duree)
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Initialize()
    Dim LargeurEcran As Long
    Dim HauteurEcran As Long
    
    ' Obtenez la largeur et la hauteur de l'écran principal
    LargeurEcran = Application.Width
    HauteurEcran = Application.Height
    
    ' Positionnez le UserForm au centre de l'écran principal
    Me.StartUpPosition = 0 ' 0 signifie que vous définissez la position manuellement
    Me.Left = (LargeurEcran - Me.Width) / 2
    Me.Top = (HauteurEcran - Me.Height) / 2
End Sub

