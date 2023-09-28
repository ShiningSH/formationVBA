Option Explicit

' Lance la verification
Private Sub bouttonLancerVerification_Click()
    Dim motDePasseAttendu As String
    Dim motDePasseEntree As String
    ' Définir le mot de passe attendu
    motDePasseAttendu = "thisIsNot"
    ' Récupérer le mot de passe entré dans la TextBox
    motDePasseEntree = Trim(password.Value)
    
    
    Dim validation As Worksheet     ' Pages cachée 1
    Dim matrice As Worksheet        ' Pages cachée 2
    Set validation = Workbooks(nom_fichier_source).Worksheets(nom_feuille_datas_validation)
    Set matrice = Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices)
    
    ' Vérifier si le mot de passe entré correspond au mot de passe attendu
    If motDePasseEntree = motDePasseAttendu Then
        MsgBox "Mot de passe correct !" & vbCrLf & "Debug : N'oubliez pas d'actualiser les données dans le module 'formation1_root' "
        matrice.Visible = xlSheetVisible
        validation.Visible = xlSheetVisible
        Unload Me
    Else
        Call OuvrirFichierAudio
    End If
End Sub
Private Sub Fermer_Click()
    
    Unload Me
    
    
End Sub

' Titre
Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

' Liste deroulante selection admin ou compte invité
Private Sub listeDeroulante_Change()
    If listeDeroulante.Value = "Invité" Then
        password.Visible = False
        Fermer.Visible = True
        Label1.Caption = "Sélectionnez un compte"
        Label3.Visible = False
        bouttonLancerVerification.Visible = False
    Else
        password.Visible = True
        bouttonLancerVerification.Visible = True
        Label1.Caption = "Entrez le mot de passe"
        Fermer.Visible = False
        Label3.Visible = True
    End If
End Sub
' Entrer le mot de passe
Private Sub password_Change()
    
End Sub
Private Sub UserForm_Initialize()
    listeDeroulante.Value = "Invité"
    password.Visible = False
    bouttonLancerVerification.Visible = False
    Fermer.Visible = True
    
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

Private Sub UserForm_Click()

End Sub

