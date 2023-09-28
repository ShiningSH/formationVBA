Option Explicit
Dim i As Integer
Dim ii As Integer
Sub modificationDeLaLigne(cheminDossier As String, indice As Integer, nomFichierActuelDassault As String, nomFeuilleDassault As String)

    Dim fso As Object                   ' FileSystemObject
    Dim dossier As Object               ' Folder
    Dim fichier As Object               ' File
    Dim nomFichierBase As String
    Dim cheminFichier As String
    Dim fichierOuvert As Workbook
    
    indice = indice - 1                 ' Calcul l'ancien indice
        
    
    ' Nom de base du fichier
    nomFichierBase = "TDCR5_Sunaero_rev"
    
    ' Créer une instance de FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Parcourir les fichiers dans le dossier
    For Each fichier In fso.GetFolder(cheminDossier).Files
        If InStr(fichier.Name, nomFichierBase & indice) > 0 Then
        
            ' Construire le chemin complet du fichier
            cheminFichier = fichier.Path
            Exit For ' Sortir de la boucle dès qu'on a trouvé le fichier
        End If
    Next fichier
    
    ' Vérifier si le fichier a été trouvé
    If Len(cheminFichier) > 0 Then
        ' Ouvrir le fichier Excel et attribuer à un objet Workbook
        Set fichierOuvert = Workbooks.Open(cheminFichier)
        
        ' Utiliser fichierOuvert pour effectuer des opérations sur le fichier
        ' Par exemple : fichierOuvert.Sheets(1).Range("A1").Value = "Nouvelle valeur"
    Else
        MsgBox "Fichier introuvable pour l'indice " & indice
    End If
    
    ' Lire et ecrire l'ancien tableau
    Dim monTableau() As Variant
    monTableau = calculerMatrice(fichierOuvert, "TDCR5")
    
    ' Fermer le fichier
    fichierOuvert.Close SaveChanges:=False
    
    ' Comparer les deux fichiers
    Call comparerLesDates(nomFichierActuelDassault, nomFeuilleDassault, monTableau, fichierOuvert)
    
    
    ' Libérer les objets FileSystemObject
    Set fichier = Nothing
    Set dossier = Nothing
    Set fso = Nothing
    
    
End Sub
Sub comparerLesDates(fichierActuel As String, Feuille As String, matriceAncienFichier As Variant, nomAncienFichier As Workbook)
    
    ' Compter le nombre de colonne / lignes dans le nouveau fichier
    Dim nbColonnesMaxNouveau As Integer
    Dim nbLignesMaxNouveau As Integer
    
    ' Critere communs des deux fichiers
    Dim critereCommunNouveau As String
    Dim critereCommunAncien As String
    
    Dim k As Integer ' Compteur de correspondance pour trouver a qu'elle ligne dans l'ancien fichier se trouve l'employé actuel
    k = 0
    
    ' Compter jusqu'à la derniere colonne du fichier actuel
    ' Initialisation a 1
    nbColonnesMaxNouveau = 1
    While Workbooks(fichierActuel).Worksheets(Feuille).Cells(1, nbColonnesMaxNouveau).Value <> "TDCR5"
        nbColonnesMaxNouveau = nbColonnesMaxNouveau + 1
    Wend
    
    
    ' Compter le nombre de lignes (On peut avoir un nombre de salarié supplementaire mais pas un nombre de colonne supplementaire)
    ' Initialisation a 1
    nbLignesMaxNouveau = 1
    While Workbooks(fichierActuel).Worksheets(Feuille).Cells(nbLignesMaxNouveau + 4, 1).Value <> ""
        nbLignesMaxNouveau = nbLignesMaxNouveau + 1
    Wend
    
    ' Tester si le fichier Dassault a évolué depuis le dernier envoie
    Dim nbColonnesMaxAncien As Integer ' Variable anciens fichier envoyé
    nbColonnesMaxAncien = UBound(matriceAncienFichier, 2)
    
    If nbColonnesMaxNouveau = nbColonnesMaxAncien Then
        
        For i = 1 To nbLignesMaxNouveau
            ' Prendre le premier critere commun (NOM + PRENOM)
            critereCommunNouveau = Workbooks(fichierActuel).Worksheets(Feuille).Cells(i + 4, 1).Value & Workbooks(fichierActuel).Worksheets(Feuille).Cells(i + 4, 2).Value
            
            ' ??
            
            If i < UBound(matriceAncienFichier, 1) - 2 Then
            
                critereCommunAncien = matriceAncienFichier(i + 2, 1) & matriceAncienFichier(i + 2, 2) ' & matriceAncienFichier(i + 2, 2)
            Else
                critereCommunAncien = ""
            End If
ConditionDessus:
            If critereCommunNouveau = critereCommunAncien And critereCommunNouveau <> "" Then
                
                For ii = 1 To nbColonnesMaxNouveau - 8
                    ' Si jamais on a la chance d'avoir le meme ordre des lignes
                    If k = 0 Then
                        
                        ' Vérifie que la date n'ai pas changé
                        If matriceAncienFichier(i + 2, ii + 4) <> Workbooks(fichierActuel).Worksheets(Feuille).Cells(i + 4, ii + 4).Value Then
                            Workbooks(fichierActuel).Worksheets(Feuille).Cells(i + 4, 14).Value = "Oui"
                            Exit For
                        Else
                            Workbooks(fichierActuel).Worksheets(Feuille).Cells(i + 4, 14).Value = "Non"
                        End If
                            
                    ' Si on a pas le meme ordre dans nos lignes
                    ElseIf k <> 0 Then
                    
                        If matriceAncienFichier(k + 2, ii + 4) <> Workbooks(fichierActuel).Worksheets(Feuille).Cells(i + 4, ii + 4).Value Then
                            Workbooks(fichierActuel).Worksheets(Feuille).Cells(i + 4, 14).Value = "Oui"
                            Exit For
                        Else
                            Workbooks(fichierActuel).Worksheets(Feuille).Cells(i + 4, 14).Value = "Non"
                        End If
                        
                    End If
                    
                Next ii
                
            ' trouver la position exacte dans l'ancien fichier
            ElseIf critereCommunNouveau <> critereCommunAncien And critereCommunNouveau <> "" Then
                k = 0 ' Reset du compteur à 0
                While critereCommunNouveau <> critereCommunAncien
                    If k < UBound(matriceAncienFichier, 1) - 2 Then
                        k = k + 1
                        critereCommunAncien = matriceAncienFichier(k + 2, 1) & matriceAncienFichier(k + 2, 2)
                    Else
                        
                        GoTo prochainEmployé
                        
                    End If
                    
                Wend
                GoTo ConditionDessus
            End If
            
prochainEmployé:
        Next i
        
    Else
        MsgBox "Le fichier Dassault a changé. Veuillez remplir manuellement la cellule Modification de la ligne "
    
        Exit Sub
    End If
    
    
    
    ' Savoir si les colonnes ont changé (Comparer avec la limite de colonne fonction existante)
        ' msgbox "Veuillez remplir manuellement la celule" => booleen
        
    ' Sinon savoir
            ' si les dates sont différentes de l'ancien au nouveau fichier
                ' oui dans le fichier excel
            ' Else
                ' non
                
    

End Sub
Public Function calculerMatrice(nomFichier As Workbook, nomFeuille As String) As Variant
    
    
    Dim nbColonnesMax As Integer
    Dim nbLignesmax As Integer
    Dim premiere_ligne As Integer
    
    ' Calcul des lignes max
    nbLignesmax = 1
    
    ' Trouver le nombre de colonnes maximum
    While nomFichier.Worksheets(nomFeuille).Cells(nbLignesmax + 2, 1).Value <> ""
        nbLignesmax = nbLignesmax + 1
    Wend
    
    
    ' Calcul des colonnes max
    nbColonnesMax = 1
    
    ' Trouver le nombre de colonnes maximum
    While nomFichier.Worksheets(nomFeuille).Cells(1, nbColonnesMax).Value <> "TDCR5"
        nbColonnesMax = nbColonnesMax + 1
    Wend
    
    
    ReDim mon_tableau(1 To nbLignesmax, 1 To nbColonnesMax)
    
    ' Lire tableau
    For i = 1 To nbLignesmax
        For ii = 1 To nbColonnesMax
            mon_tableau(i, ii) = nomFichier.Worksheets(nomFeuille).Cells(2 + i, ii).Value
        Next ii
    Next i
    

    calculerMatrice = mon_tableau
End Function



