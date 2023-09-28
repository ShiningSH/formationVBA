Option Explicit
Sub Air_France()
    
    ' Désactiver les maj de l'ecran
    Application.ScreenUpdating = False
    
    ' Appel la fonction pour creer le fichier et enregistrer son nom
    Dim fileAirFrance As String
    fileAirFrance = createFile
    
    ' Modification du nom retourné pour retrouver dans fichier
    fileAirFrance = fileAirFrance & ".xlsx"
    
    ' Ouverture du fichier template d'origine AirFrance
    Dim emplacementTemplate As String
    emplacementTemplate = emplacement_templateAFI
    Dim fichierTemplate As Workbook
    Set fichierTemplate = openFile(emplacementTemplate)
    
    ' Vérifier si le fichier "TEMPLATE" a été trouvé et ouvert avec succès
    If Not fichierTemplate Is Nothing Then
        ' Copier toutes les informations de ce fichier
        Cells.Select
        Selection.Copy
        
        ' Fermer le fichier "TEMPLATE" sans enregistrer les modifications (vous pouvez commenter cette ligne si vous souhaitez enregistrer les modifications)
    Else
        MsgBox "Le fichier ""TEMPLATE"" n'a pas été trouvé dans le dossier."
        Exit Sub ' Quitter la macro si le fichier    n'a pas été trouvé
    End If

    ' Ouverture du fichier Air France
    Dim fichierAirFrance As Workbook
    Dim fichierDestination As String ' Chemin de destination du stockage de ce fichier
    fichierDestination = empl_stockage_fichier_AFI
    Set fichierAirFrance = openFile(fichierDestination)
    
    ' Vérifier si le fichier nom_feuille_ecriture_AFI a été trouvé et ouvert avec succes
    If Not fichierAirFrance Is Nothing Then
    
        ' Coller les informations copiées depuis "TEMPLATE" dans le fichier nom_feuille_ecriture_AFI
        Range("A1").Select
        ActiveSheet.Paste
        
        ' Desactive la selection
        Application.CutCopyMode = False
        
        ' Fermer le fichier "TEMPLATE" sans enregistrer les modifications
        fichierTemplate.Close False
        
        ' Compter le nombre de salaries sur le fichier Main
        Dim plage As Range
        Set plage = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Range("A9:A110")  ' Mettre a jour la plage ci-besoin
        Dim nombreSalaries As Integer
        nombreSalaries = Application.WorksheetFunction.CountA(plage)
        
        ' Procedure ajoutant que les salaries ayant "VRAI" a la colonne "BD" du fichier ENG 007
        Dim mainArray() As String
        Dim nombreEmployeAF As Integer
        nombreEmployeAF = importer_tableau(fileAirFrance, mainArray, nombreSalaries, 5, 0)
        
        ' Sauvegarde du fichier
        fichierAirFrance.Save
    Else
        ' Pas de fichier trouvé
        MsgBox "Le template AFI n'a pas été trouvé dans le dossier. Veuillez vous rendre ici : " & emplacement_templateAFI
        Exit Sub
    End If
     
    ' Remplir les dates
    Call remplirDatesAirFrance(fileAirFrance, nombreEmployeAF)
    Call miseEnFormeDocument(nombreEmployeAF, fileAirFrance)
    
    fichierAirFrance.Save ' Sauvegarde du fichier

    Call instructionFinale("Air France")
    
    ' Re activer le rafraîchissement de la page
    Application.ScreenUpdating = True
    
End Sub
Sub instructionFinale(nom)
    MsgBox "Le fichier " & nom & " a été généré ! Veuillez suivre les instructions ci-dessous pour garantir l'intégrité des données :" & vbCrLf & _
    vbCrLf & "1. Prenez le temps de vérifier visuellement toutes les informations contenues dans le fichier Air France." & vbCrLf & _
    vbCrLf & "2. Si vous constatez des informations manquantes, suivez attentivement les étapes suivantes :" & vbCrLf & _
    vbCrLf & "   a. Référez-vous aux instructions fournies pour obtenir les informations manquantes." & vbCrLf & _
    vbCrLf & "   b. Suivez rigoureusement les procédures indiquées dans les instructions pour récupérer les données nécessaires." & vbCrLf & _
    vbCrLf & "   c. En cas de difficulté, n'hésitez pas à contacter notre service d'assistance dédié pour obtenir de l'aide supplémentaire." & vbCrLf & _
    vbCrLf & "3. Une fois que toutes les informations requises ont été vérifiées et complétées, assurez-vous de sauvegarder le fichier Air France." & vbCrLf & "Bonne journée !", vbInformation, "Vérification des informations"
End Sub
Sub miseEnFormeDocument(nombreEmployeAF As Integer, fileAirFrance As String)
    
    ' Redimensionner les colonnes
    Columns("A:A").Select
    Selection.ColumnWidth = 25
    Columns("B:B").Select
    Selection.ColumnWidth = 25
    Columns("C:C").Select
    Selection.ColumnWidth = 25
    Columns("D:W").Select
    Selection.ColumnWidth = 10
   
    Rows(nombreEmployeAF + 6 & ":" & nombreEmployeAF + 10).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.ClearFormats

    ' Le tri de la plage pour avoir tout ensemble
    Range("A6:A48").Select
    ActiveWorkbook.Worksheets(nom_feuille_ecriture_AFI).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(nom_feuille_ecriture_AFI).Sort.SortFields.Add2 Key:=Range("A6:A48" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(nom_feuille_ecriture_AFI).Sort
        .SetRange Range("A6:Y48")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Message en bas de page

    Range("B" & (nombreEmployeAF + 8) & ":D" & (nombreEmployeAF + 9)).Select
    Range("D" & (nombreEmployeAF + 9)).Activate

    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    ' Ecrire message
    Workbooks(fileAirFrance).Worksheets(nom_feuille_ecriture_AFI).Cells(nombreEmployeAF + 8, 2) = "(6) = Uniquement si impliqué dans la Maintenance d'aéronefs non EASA et hors Habilitation APRS"
    
    ' Mettre une épaisseur de ligne
    Range("B" & (nombreEmployeAF + 8) & ":D" & (nombreEmployeAF + 9)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
   
   ' S'occuper de la mise en forme conditionnel + Traits de tableau
    
    ' Mettre en forme les données
    Range("E6:X68").Select
    
    With Selection.Font
        .Name = "Calibri"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    
    ' Tracere les tableaux
    Dim lignesEnTout As Integer
    lignesEnTout = nombreEmployeAF + 5
    
    Range("A6:Y" & lignesEnTout).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    
    ' Supprimer
    
    
    Rows(lignesEnTout + 1 & ":50").Select
    Selection.FormatConditions.Delete
    
End Sub
Sub remplirDatesAirFrance(nomFichier As Variant, nombreSalaries As Integer)
    
    
    ' Mes variables :
    
    Dim i As Integer                            ' Compteur de recherche de ligne (Fichier Air France)
    Dim c As Integer                            ' Compteur pour Boucle de recopiage
    Dim k As Integer                            ' Compteur de recherche de ligne (Fichier Main)

    Dim critereCommunAirFrance As String        ' Criteres communs (NOM + Prénom) Air France File
    Dim critereCommunMain As String             ' Criteres communs (NOM + Prénom) Fichier main
    
    Dim formuleIndex As Integer
    Dim colENG As String
    
    Dim maDate As Date                          ' Date de base dans le fichier Main
    Dim nouvelleDate As Date                    ' Date demandee pour le fichier Airbus
    Dim dateAEnlever As Integer                 ' Ceci est le nombre année a enlever par rapport a la date
    Dim maValeur As String
    
    For i = 1 To nombreSalaries

        ' Check de la fin des lignes
        If Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_AFI).Cells(i + 5, 2).Value <> " " Then
            ' Critere unique fichier Air France
            critereCommunAirFrance = Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_AFI).Cells(i + 5, 2).Value & Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_AFI).Cells(i + 5, 3).Value
            ' Critere unique fichier main
            critereCommunMain = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(k + 8, 1).Value & Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(k + 8, 2).Value

            ' Check de la fin des lignes
            If Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(i + 5, 1).Value <> " " Then
a:
                If critereCommunAirFrance = critereCommunMain Then
                    
                    ' la ligne est trouvée ==> les donnees des colonnes
                    For c = 1 To 20 ' nb de col dans fichier Air France
                        colENG = "B" & c + 1
                                
                        ' Vérifie que dans ma matrice on ne copie pas les colonnes non utilisées du fichier Air France
                        ' Attention ce code peut etre modifié si on modifie la taille des matrices
                        If Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Cells(c + 1, 3).Value <> "RIEN" Or _
                            Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Cells(Rows.Count, 3).End(xlUp).Row < 23 Then
                                
                            ' Trouver la colonne dans le fichier Main des dates (Index Equiv)
                            formuleIndex = Application.WorksheetFunction.Index(Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Range("A1:D22"), _
                            Application.WorksheetFunction.Match(Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Range(colENG), Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Range("B1:B22"), 0), 4)
                            
                            ' Trouver la date        
                            maValeur = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(k + 8, formuleIndex)

                            If maValeur = "Oui" Then
                                Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_AFI).Cells(i + 5, c + 4).Value = "Yes"
                            ElseIf maValeur = "Non" Then
                                Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_AFI).Cells(i + 5, c + 4).Value = "No"
                            ' Date différente de rien ?
                            ElseIf InStr(Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(k + 8, formuleIndex).Value, "/") > 0 Then
                                maDate = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(k + 8, formuleIndex)
                                dateAEnlever = 0
                                If Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(7, formuleIndex) = 2 Or Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(7, formuleIndex) = 3 Or Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(7, formuleIndex) <> "-" Then
                                    dateAEnlever = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(7, formuleIndex)
                                End If
                                ' Lire la nouvelle date
                                nouvelleDate = DateAdd("yyyy", -dateAEnlever, maDate)
                                ' Ecrire la date dans le fichier Air France
                                Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_AFI).Cells(i + 5, c + 4).Value = nouvelleDate
                            End If
                        End If
                    Next c
                                        
                ' Recherche la ligne dans le fichier main
                ElseIf critereCommunAirFrance <> critereCommunMain Then
                    
                    ' [KO] Prevoir un conditionnel pour signaler qu'on a pas trouvé l'utilisateur dans le fichier ce qui arrete la recherche du fichier
                    k = 0 ' Reset du compteur a 0
                    
                    While critereCommunAirFrance <> critereCommunMain
                        k = k + 1 ' Incremente de 1 a chaque début
                        critereCommunMain = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(k + 8, 1).Value & Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(k + 8, 2).Value
                    Wend
                    GoTo a
                End If
            End If
        End If
    Next i
End Sub
Private Function createFile()

    Dim nomFichier As String
    Dim rev As String
    Dim dateDuJour As String
    
    Dim nbFichiersEnvoyesString As String
    
    nbFichiersEnvoyesString = empl_stockage_envoyAFI
    
    ' Obtenir la révision
    rev = CompterFichiersDansDossier(nbFichiersEnvoyesString) + 16
    
    ' Obtenir la date du jour et la formater au format "ddmmyyyy"
    dateDuJour = Format(Date, "ddmmyyyy")
    
    ' Définir le nom du fichier avec la révision et la date du jour
    nomFichier = "SUNAERO Appendix 2 SUBCONTRACTORSS STAFFS TRAININGS AUTORISATIONS_rev" & rev & " du " & dateDuJour
    
    ' Ajouter un nouveau classeur et donner un nom personnalisé
    Dim nouveauClasseur As Workbook
    Set nouveauClasseur = Workbooks.Add
    nouveauClasseur.SaveAs empl_stockage_fichier_AFI & nomFichier & ".xlsx"
    
    ' Renomer la feuille SUNAERO
    Sheets("Feuil1").Name = nom_feuille_ecriture_AFI
    
    ' Retourner le nom du fichier
    createFile = nomFichier
End Function
Public Function openFile(cheminDossier As String) As Workbook
    
    Dim fichierTrouve As String
    
    ' Recherche du premier fichier dans le dossier
    fichierTrouve = Dir(cheminDossier & "\*")
  
    ' Ouvrir le seul fichier trouvé et retourner l'objet Workbook
    Set openFile = Workbooks.Open(cheminDossier & "\" & fichierTrouve)
End Function
Function ObtenirNomFichier()
    Dim cheminDossier As String
    Dim nomFichier As String

    ' Obtenez le nom du premier fichier dans le dossier spécifié
    nomFichier = Dir(empl_stockage_fichier_AFI)

    ' Vérifiez si un fichier a été trouvé
    If nomFichier <> "" Then
        ' Affichez le nom du fichier
        ObtenirNomFichier = nomFichier
    Else
        MsgBox "Aucun fichier trouvé dans le dossier spécifié."
    End If
End Function
Sub SupprimerEspacesFinRange(ByRef rng As Range)
    Dim cell As Range
    ' Vérifier si la plage n'est pas vide
    If Not rng Is Nothing Then
        ' Parcourir chaque cellule dans la plage
        For Each cell In rng
            ' Utiliser la fonction Trim pour supprimer les espaces à la fin de la valeur de la cellule
            cell.Value = Trim(cell.Value)
        Next cell
    End If
End Sub
Public Function importer_tableau(nomFichier As Variant, mon_tableau() As String, nombre_lignes As Integer, premiere_ligne As Integer, premiere_colonne As Integer) As Integer

    Dim i As Integer
    Dim J As Integer
    Dim nombreVrai As Integer
    Dim nbColonnesMax As Integer
    
    nbColonnesMax = 1
    
    ' Trouver le nombre de colonnes maximum
    While Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8, nbColonnesMax).Value <> "Dassault"
        nbColonnesMax = nbColonnesMax + 1
    Wend
    
    ReDim mon_tableau(1 To nombre_lignes, 1 To nbColonnesMax)
    
    ' Lire tableau
    For i = 1 To nombre_lignes
        For J = 1 To nbColonnesMax
            mon_tableau(i, J) = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, J).Value
        Next J
    Next i
    
    nombreVrai = 0
    
    ' ecrire
    For i = 1 To nombre_lignes
        If mon_tableau(i, nbColonnesMax - 1) = "Oui" Then
            nombreVrai = 1 + nombreVrai ' Incremente le nombre de fois ou j'ai trouve un employé airFrance
            Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_AFI).Cells(nombreVrai + premiere_ligne, 2).Value = mon_tableau(i, 1)
            Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_AFI).Cells(nombreVrai + premiere_ligne, 3).Value = mon_tableau(i, 2)
            Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_AFI).Cells(nombreVrai + premiere_ligne, 4).Value = mon_tableau(i, 4)
            Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_AFI).Cells(nombreVrai + premiere_ligne, 1).Value = mon_tableau(i, nbColonnesMax - 3)
        End If
    Next i
    importer_tableau = nombreVrai
End Function

