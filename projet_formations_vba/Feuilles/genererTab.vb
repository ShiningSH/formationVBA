Option Explicit

Private Sub CommandButton1_Click()
    Call Dassault
    Unload Me
    menu.Hide
End Sub

Private Sub CommandButton2_Click()
    Call Air_France
    Unload Me
    menu.Hide
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

