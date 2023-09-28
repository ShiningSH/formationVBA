Option Explicit
Dim Cellule As Range
Sub MettreEnGrisLesCellulesVides(MaPlage As Range)

    ' Parcours chaque cellule dans la plage
    For Each Cellule In MaPlage
        ' Vérifiez si la cellule est vide (c'est-à-dire si elle contient "")
        If Cellule.Value = "" Then
            ' Si la cellule est vide, mettez sa couleur de fond en gris
            Cellule.Interior.Color = RGB(89, 89, 89) ' Vous pouvez ajuster la couleur en fonction de vos préférences
            ' Cellule.Value = "N/A"
        End If
    Next Cellule
    
End Sub
Sub ResetCouleurBlanc(MaPlage As Range)

    For Each Cellule In MaPlage
        ' Vérifiez si la couleur de fond de la cellule est gris (RGB(192, 192, 192))
        If Cellule.Interior.Color = RGB(89, 89, 89) And Cellule.Value <> "" Then
            ' Si la couleur e fond est gris, réinitialisez-la à blanc
            Cellule.Interior.Color = RGB(255, 255, 255)
        End If
    Next Cellule
    
End Sub

Sub resetPlage(MaPlage)
   MaPlage.Interior.Color = RGB(255, 255, 255) ' Réinitialise la couleur de fond à blanc
End Sub

Sub appelMenu()

    menu.Show
End Sub

