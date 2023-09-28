Option Explicit

Private Sub CommandButton1_Click()
    acces_mission
    Unload Me
    MsgBox "Opération terminée", vbInformation, "Message"
    
End Sub

Private Sub CommandButton2_Click()
    Call expiration_date
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    Call resetCommentaires
    Unload Me
    
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub CommandButton5_Click()
    Application.ScreenUpdating = False
    Dim MonClasseur As Workbook
    Dim MaFeuille As Worksheet
    Dim MaPlage As Range
    
    Set MonClasseur = Workbooks(nom_fichier_source)
    Set MaFeuille = MonClasseur.Worksheets(nom_feuille_ENG_007)
    
    ' Plage (range) a traiter en utilisant Set
    Set MaPlage = MaFeuille.Range("H9:BV118")
    
    
     ' Mettre en gris les cellules vides
    Call MettreEnGrisLesCellulesVides(MaPlage)

    
    ' Mettre en blanc si modif
    Call ResetCouleurBlanc(MaPlage)
    
    
    Unload Me
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton6_Click()
    Call Generer_Tableaux.Show
End Sub

Private Sub CommandButton7_Click()
    Call Afficher
    Unload Me
End Sub

Private Sub CommandButton8_Click()
    Call Login.Show
    Unload Me
End Sub

Private Sub CommandButton9_Click()
    Application.ScreenUpdating = False
    
    Call reset
    Unload Me
    Application.ScreenUpdating = True
    Dim MonClasseur As Workbook
    Dim MaFeuille As Worksheet
    
End Sub
Sub reset()
    Dim MaPlage As Range
    Dim MonClasseur As Workbook
    Dim MaFeuille As Worksheet
    
    Set MonClasseur = Workbooks(nom_fichier_source)
    Set MaFeuille = MonClasseur.Worksheets(nom_feuille_ENG_007)
    
    ' Plage (range) a traiter en utilisant Set
    Set MaPlage = MaFeuille.Range("H9:BV118")
    
    Call resetPlage(MaPlage)
    
     ' Mettre en gris les cellules vides
    Call MettreEnGrisLesCellulesVides(MaPlage)

    
    ' Mettre en blanc si modif
    Call ResetCouleurBlanc(MaPlage)
    
End Sub

Private Sub Label1_Click()

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

