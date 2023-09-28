Option Explicit
' Procedure si l'employé peut acceder a la mission en fonction de l
Sub acces_mission()

    Application.ScreenUpdating = False
    
    ' Déclaration des variables :
    
    Dim i As Integer                                    ' Compteur 1
    Dim ii As Integer                                   ' Compteur 2
    Dim iii As Integer                                  ' Compteur 3
    Dim categorieENG As String                          ' Categorie en cours ENG 007
    Dim categorieM1065 As String                        ' Categorie en cours M1065
    Dim nombreLignesENG007 As Integer                   ' Nombre de lignes max ENG 007 (categorie)
    Dim nombreLignesM1065 As Integer                    ' Nombre de lignes max M1065 (categorie)
    Dim nbMaxM1065 As Integer                           ' Limite de colonnes M1065
    Dim colonneMatrice As Variant                       ' Numéro colonne actuelle ENG007
    Dim ok As Integer                                   ' Compter le nb de formations à jour
    Dim ko As Integer                                   ' Compter le nb de formations manquantes
    Dim dejaFaux As Boolean                             ' Retourner True or False si une formation est pas a jour
    Dim prochaineButee() As Variant                     ' Mon tableau de prochaine butée
    Dim dateButee As Variant                            ' Ma date de prochaine butée
    
    ' Nombre de lignes dans l'ENG 007 :
    nombreLignesENG007 = 1
    While Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(nombreLignesENG007 + 8, 1) <> ""
        nombreLignesENG007 = nombreLignesENG007 + 1
    Wend
    nombreLignesENG007 = nombreLignesENG007 - 1 ' (Moins nb vehicules ??)
    
    ' Reset de la range
    Range("E9:E" & nombreLignesENG007 + 8).ClearContents
    Range("E9:E" & nombreLignesENG007 + 8).ClearFormats
    
    ' Nombre de lignes dans M1065.0_Appendix B_8
    nombreLignesM1065 = 1
    While Workbooks(nom_fichier_source).Worksheets(nom_feuille_M1065).Cells(nombreLignesM1065 + 9, 1) <> ""
        nombreLignesM1065 = nombreLignesM1065 + 1
    Wend
    nombreLignesM1065 = nombreLignesM1065 - 1
    
    ' Nombre max lignes dans la range tri_des_matrices
    nbMaxM1065 = 1
    While Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Cells(nbMaxM1065, 15) <> ""
        nbMaxM1065 = nbMaxM1065 + 1
    Wend
    nbMaxM1065 = nbMaxM1065 - 2
    
    ' Boucle catégorie ENG 007
    For i = 1 To nombreLignesENG007
        
        ' Initialisation des variables au prochain user
        ReDim prochaineButee(1 To 30) ' Reset du tableau
        ko = 0
        ok = 0
        dejaFaux = False
        
        'Initialisation de la categorie
        categorieENG = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(i + 8, 4)
        
        ' Boucle pour aller chercher la categorie dans la matrice
        For ii = 1 To nombreLignesM1065

            categorieM1065 = Workbooks(nom_fichier_source).Worksheets(nom_feuille_M1065).Cells(ii + 9, 1)
            
            ' Si les deux categories sont égales
            If categorieM1065 = categorieENG Then
                For iii = 1 To 30
                    
                    ' Debug : Affiche le nom de la formation obligatoire
                     
                    ' MsgBox ("Le nom de la formation dans la matrice trouvée est : " & Workbooks(nom_fichier_source).Worksheets(nom_feuille_M1065).Cells(4, iii + 2).Value)
                    
                    ' Vérifier si la formation est demandée
                    
                    If Workbooks(nom_fichier_source).Worksheets(nom_feuille_M1065).Cells(9 + ii, iii + 2).Value = "X" Then
                        
                        
                        ' Trouver le correspondant
                        colonneMatrice = Application.VLookup(Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices) _
                        .Range("O" & iii + 1), Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Range("O1:P31"), 2, False)
                        
                        ' Remplir le tableau des dates obligatoires
                        prochaineButee(iii) = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice).Value
                        
                        
                        ' Debug : Afficher la valeur dans la cellule
                        ' MsgBox Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice).Value
                        
                        ' Vérifier si la formation est dépassée
                        If Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice).Value > Date _
                           And Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice).Value <> "" Then
                            
                            ' Mon compteur ok s'incrémente de 1
                            ok = ok + 1
                            
                            ' Si vrai
                            With Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, 5)
                                .Value = "Autorisé"
                                .Font.Color = RGB(255, 255, 255) ' Blanc
                                .Interior.Color = RGB(50, 205, 50) ' Vert
                                .Font.Size = 8
                                .Font.Name = "Times New Roman"
                            End With
                        
                        ' Cas particuliers formations devant juste être passées sans révisions (Ex : English, sensibilisation qualité)
                        ElseIf colonneMatrice = 26 Or _
                               colonneMatrice = 22 Or _
                               colonneMatrice = 25 Or _
                               colonneMatrice = 8 Or _
                               colonneMatrice = 10 Or _
                               colonneMatrice = 11 Or _
                               colonneMatrice = 12 Or _
                               colonneMatrice = 18 Or _
                               colonneMatrice = 19 Then
                            
                            ' La formation a été donnée
                            If Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice).Value <> "" Then
                                
                                ok = ok + 1
                                With Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, 5)
                                    .Value = "Autorisé"
                                    .Font.Color = RGB(255, 255, 255) ' Blanc
                                    .Interior.Color = RGB(50, 205, 50) ' Vert
                                    .Font.Size = 8
                                    .Font.Name = "Times New Roman"
                                End With
                                
                            ' La formation n'a pas été donnée
                            Else
                                dejaFaux = True ' Donc mon booleen passe en faux
                                                                ' Dire que dans quelle cellule la formation est manquante
                                With Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice)
                                    .Interior.Color = RGB(255, 0, 0)
                                End With
    
                                ' Mettre a jour "Non autorisé (*)"
                                With Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, 5)
                                    .Value = "Non autorisé (*)"
                                    .Font.Color = RGB(255, 255, 255) ' Blanc
                                    .Interior.Color = RGB(220, 20, 60) ' Rouge
                                    .Font.Size = 8
                                    .Font.Bold = True
                                    .Font.Name = "Times New Roman"
                                    .Characters(13, 16).Font.Bold = False
                                End With
                            
                            End If
                                
                        ' Rec = ""  a jour mais qq chose en en init a jour
                        ElseIf (colonneMatrice = 14 Or _
                                colonneMatrice = 16) And _
                               (Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice).Value = "" And _
                                Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice - 1).Value > Date) Then

                            ' Mon compteur ok s'incrémente de 1
                            ok = ok + 1
                                
                            With Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, 5)
                                .Value = "Autorisé"
                                .Font.Color = RGB(255, 255, 255) ' Blanc
                                .Interior.Color = RGB(50, 205, 50) ' Vert
                                .Font.Size = 8
                                .Font.Name = "Times New Roman"
                            End With
                              
                        ' Cas particuliers : Mon Rec est remplie et Initial = Rien ou Date Périmée
         
                        ElseIf (colonneMatrice = 13 Or _
                               colonneMatrice = 15) And _
                               (Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice + 1).Value <> "") Then
                             
                            ' Mon rec est a jour
                            
                            If Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice + 1).Value > Date And _
                               Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice).Value <> "" Then
                               
                                ' Mon compteur ok s'incrémente de 1
                                ok = ok + 1
                                
                                With Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, 5)
                                    .Value = "Autorisé"
                                    .Font.Color = RGB(255, 255, 255) ' Blanc
                                    .Interior.Color = RGB(50, 205, 50) ' Vert
                                    .Font.Size = 8
                                    .Font.Name = "Times New Roman"
                                End With

                            End If
             
                        ' Debug
                        ' MsgBox Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice).Value
                        
                        ' Si date est dépassée par rapport a aujourd'hui ou si cellule vide et que j'ai rien en rec
suite:
                        ElseIf Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice).Value < Date Or _
                               Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice).Value = "" Then
    

                            dejaFaux = True
                            
                            ko = ko + 1
                            
                            ' Dire que dans quelle cellule la formation est manquante
                            With Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, colonneMatrice)
                                .Interior.Color = RGB(255, 0, 0)
                            End With

                            ' Mettre a jour "Non autorisé (*)"
                            With Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, 5)
                                .Value = "Non autorisé (*)"
                                .Font.Color = RGB(255, 255, 255) ' Blanc
                                .Interior.Color = RGB(220, 20, 60) ' Rouge
                                .Font.Size = 8
                                .Font.Bold = True
                                .Font.Name = "Times New Roman"
                                .Characters(13, 16).Font.Bold = False
                            
                            End With
                            
                            ' Je change de personne
                            ' GoTo debut
                        End If
                        
                    ' Si j'ai rien dans ma case
                    ElseIf Workbooks(nom_fichier_source).Worksheets(nom_feuille_M1065).Cells(9 + ii, iii + 2).Value = "" Then
                    
                        ' Date pas a jour donc le compteur des ko est incrémenté
                        ' ko = ko + 1
                        
                    End If
                    
boucleSuivante:
                    If iii = 30 Then
                      GoTo debut2
                    End If
                Next iii
                
            ' Si c'est un véhicule
            ElseIf categorieENG = "Administration Staff" Or _
                   categorieENG = "Accounting Manager" Or _
                   categorieENG = "Administrative Assistant" Or _
                   categorieENG = "International Sales Director" Or _
                   categorieENG = "Production and Purchasing Manager" Or _
                   categorieENG = "Manager Customer Support" Then
                
                With Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, 5)
                    .Value = "Non applicable"
                    .Font.Color = RGB(255, 255, 255) ' Blanc
                    .Interior.Color = RGB(255, 165, 0) ' Orange
                    .Font.Size = 8
                    .Font.Bold = True
                    .Font.Name = "Times New Roman"
                    .Characters(13, 16).Font.Bold = False
                End With
                
                Exit For
            
            End If

            

        Next ii
        
' Sortir si la case = N/A
debut:
' Retour au départ
debut2:

    ' Si j'ai déjà eu faux
    If ko > 0 And ok < 1 Or dejaFaux = True Then

  
        
        With Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, 5)
            ' .Value = "Non autorisé (Sauf si formation planifiée dans les 6 mois et accompagnée)"
            .Value = "Non autorisé (*)"
            .Font.Color = RGB(255, 255, 255) ' Blanc
            .Interior.Color = RGB(220, 20, 60) ' Rouge
            .Font.Size = 8
            .Font.Bold = True
            .Font.Name = "Times New Roman"
            .Characters(13, 16).Font.Bold = False
        End With
        
    End If
    
        dateButee = TrouverDateMin(prochaineButee)
        
        If dateButee < "20/09/2555" Then
            Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, 77).Value = dateButee
        End If

    Next i
    
    
    
    Application.ScreenUpdating = True
End Sub
Function ContientLettres(ByVal texte As String) As Boolean
    Dim i As Integer
    Dim lettreTrouvee As Boolean
    lettreTrouvee = False
    
    For i = 1 To Len(texte)
        If IsAlpha(Mid(texte, i, 1)) Then
            lettreTrouvee = True
            Exit For
        End If
    Next i
    
    ContientLettres = lettreTrouvee
End Function
Function IsAlpha(ByVal char As String) As Boolean
    IsAlpha = UCase(char) >= "A" And UCase(char) <= "Z"
End Function
Function TrouverDateMin(tableau() As Variant) As Date
    Dim i As Integer
    Dim valeurMin As Date
    
    ' Initialisez la valeur minimale avec une date maximale
    valeurMin = DateValue("31/12/9999")
    
    ' Parcourez le tableau pour trouver la date minimale différente de ""
    For i = LBound(tableau) To UBound(tableau)
        If IsDate(tableau(i)) And tableau(i) <> "" Then
            If tableau(i) < valeurMin Then
                valeurMin = tableau(i)
            End If
        End If
    Next i
    
    ' Renvoie la valeur minimale trouvée
    TrouverDateMin = valeurMin
End Function







