Option Explicit
Dim Destinataire As String
Dim emplacement As String
Dim emplacementDestination As String
Dim OptionButton As Boolean
' Bouton Annuler
Public Sub CommandButton1_Click()
    Unload Me
End Sub
' Bouton Go
Public Sub CommandButton2_Click()
    
    Call creationMail(emplacement, Destinataire, emplacementDestination)
    
End Sub

Private Sub Label1_Click()

End Sub
' Air france
Public Sub OptionButton1_Click()

    OptionButton = True
    
    If OptionButton = True Then
        OptionButton = False
        Destinataire = "Air France"
        emplacement = empl_stockage_fichier_AFI   ' Ne pas oublier de fermer les slash inversés \\
        emplacementDestination = empl_stockage_envoyAFI
    Else
        emplacement = ""
    End If
    
End Sub
' Dassault
Public Sub OptionButton2_Click()
    
    OptionButton = True
    
    If OptionButton = True Then
        Destinataire = "Dassault"
        OptionButton = False
        emplacement = emplacement_stockage_fichier_dassault  ' Ne pas oublier de fermer les slash inversés \
        emplacementDestination = empl_stockage_envoyDassault
    Else
        emplacement = ""
    End If
    
End Sub
Public Sub UserForm_Click()
    
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
