Option Explicit
' Procedure principale
Sub Dassault()
    Application.ScreenUpdating = False
    Dim nomFeuilleDassault As String
    
    ' Indice de rev du fichier
    
    Dim revision As Integer
    revision = CompterFichiersDansDossier(empl_stockage_envoyDassault) + 5 ' Calcul l'indice du prochain fichier
    nomFeuilleDassault = nom_feuille_ecriture_dassault
    
    ' Procedure qui cree le fichier Dassault
    Call creationFichier(revision, nomFeuilleDassault)
    Call instructionFinale("Dassault")
    Application.ScreenUpdating = True
End Sub
Sub creationFichier(revision As Integer, nomFeuilleDassault As String)
    Dim cheminSource As String
    Dim cheminDestination As String
    Dim nomFichier As String
    Dim nouveauNomFichier As String
    Dim wb As Workbook
    Dim cheminAncienFichier As String
    
    
    ' Specifier le chemin source et le nom du fichier a copier
    cheminSource = emplacement_templateDassault
    nomFichier = "template_dassault.xlsx" ' nom du fichier template
    
    ' Chemin de destination
    cheminDestination = emplacement_stockage_fichier_dassault
    
    ' Définir le nouveau nom de fichier pour le fichier collé
    nouveauNomFichier = "TDCR5_Sunaero_rev" & revision & ".xlsx"
    
    ' Copier le fichier
    FileCopy cheminSource & nomFichier, cheminDestination & nouveauNomFichier
    
    ' Ouvrir le fichier collé et l'associer a un objet Workbook
    Set wb = Workbooks.Open(cheminDestination & nouveauNomFichier)
    
    
    
        
    ' Remplir le fichier des dates depuis l'ENG 007
    Call remplirDatas(nouveauNomFichier)
    
    ' Mise en page
    Workbooks(nouveauNomFichier).Worksheets(nom_feuille_ecriture_dassault).Cells(2, 6).Value = Date
    Workbooks(nouveauNomFichier).Worksheets(nom_feuille_ecriture_dassault).Cells(2, 11).Value = revision
    
    ' Remplir la cellule Modif ligne Dassault
    cheminAncienFichier = empl_stockage_envoyDassault
    
    
    
    ' Procedure qui lance si la ligne a été modifié dans le fichier généré
    Call modificationDeLaLigne(cheminAncienFichier, revision, nouveauNomFichier, nomFeuilleDassault)
    
    ' Sauvegarder le fichier associé a l'objet Workbook
    wb.Save
    
    ' Fermer le fichier
    ' wb.Close
    
    ' Libérer la mémoire
    Set wb = Nothing
    
End Sub
Sub remplirDatas(nomComplet As String)
    
    ' Compter le nombre de salaries sur le fichier Main
    Dim plage As Range
    Set plage = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Range("A9:A110")  ' Mettre a jour la plage ci-besoin
    Dim nombreSalaries As Integer
    nombreSalaries = Application.WorksheetFunction.CountA(plage)
        
    ' Remplir les informations diverses (4 premieres colonnes du fichier)
    Dim ligneDassault As Integer
    Dim mainArray() As String
    ligneDassault = importer_tableau(nomComplet, mainArray, nombreSalaries, 4, 0)
    
    
    ' Remplir les dates
    Call datesDeFin(nomComplet, ligneDassault, nombreSalaries)
    
    
End Sub
Public Sub datesDeFin(fichierDassault As String, nbLignesDassault As Integer, nbLignesMain As Integer)
     
    Dim critereCommunDassault As String
    Dim critereCommunMain As String
    Dim i As Integer                                    ' Compteur de position pour les lignes dans le fichier Dassault
    Dim c As Integer                                    ' Compteur de position des colonnes
    Dim k As Integer                                    ' Compteur de position pour les lignes dans le fichier main
    Dim colMatrice As String                            ' Cellule qui s'incremente
    Dim formuleIndex As Integer                         ' Col. retour matrice
    Dim maDate As Date                                  ' Date ecrite
    Dim monBoolean As String                            ' Case a cochee fin de colonne (Variable a renommer)
    Dim nbval As Long                                   ' Nombre de valeurs dans la matrice
    Dim matriceIndex As String                          ' Matrice générale ou je vais chercher ma colonne
    Dim zoneRecherche As String                         ' Colonne ou je souhaite chercher mon correspondant
    
    nbval = Application.WorksheetFunction.CountA(Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Range("H1:H10")) ' Changer la range si pas bon
    matriceIndex = "I1:K" & nbval
    zoneRecherche = "I1:I" & nbval
    k = 1
    
    ' Definir le critere commun Dassault
    critereCommunMain = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(9, 1).Value & Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(9, 2).Value

    For i = 1 To nbLignesDassault
         critereCommunDassault = Workbooks(fichierDassault).Worksheets(nom_feuille_ecriture_dassault).Cells(4 + i, 1).Value & Workbooks(fichierDassault).Worksheets(nom_feuille_ecriture_dassault).Cells(4 + i, 2).Value
a:
        If critereCommunDassault = critereCommunMain Then
            
            For c = 1 To nbval - 1
                            
                colMatrice = "I" & c + 1
                
                ' Vérifier qu'on cherche bien que des entiers
                If Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Cells(c + 1, 10).Value <> "RIEN" Then
                                    

                    ' Trouver la colonne dans le fichier Main des dates (Index Equiv)
                    formuleIndex = Application.WorksheetFunction.Index(Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Range(matriceIndex), _
                    Application.WorksheetFunction.Match(Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Range(colMatrice), Workbooks(nom_fichier_source).Worksheets(nom_feuille_matrices).Range(zoneRecherche), 0), 3)
                    
                    
                    ' Lire les cellules du fichier
                    monBoolean = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(k + 8, formuleIndex)
                    
                    ' Badge DFS ?
                    If monBoolean = "Oui" And formuleIndex = 42 Then
                        Workbooks(fichierDassault).Worksheets(nom_feuille_ecriture_dassault).Cells(i + 4, 15) = "Oui"
                    ' Ecrire les dates dans le fichier
                    ElseIf monBoolean = "Non" And formuleIndex = 42 Then
                        Workbooks(fichierDassault).Worksheets(nom_feuille_ecriture_dassault).Cells(i + 4, 15) = "Non"
                    ElseIf InStr(monBoolean, "/") > 0 Then
                        Dim formattedDate As String
                        formattedDate = ConvertToNumericDate(monBoolean)
                        Workbooks(fichierDassault).Worksheets(nom_feuille_ecriture_dassault).Cells(i + 4, c + 4) = formattedDate
                    End If
                 End If
            Next c
            
        ElseIf critereCommunDassault <> critereCommunMain Then
            
            k = 0 ' Reset du compteur a 0
            While critereCommunDassault <> critereCommunMain
                k = k + 1 ' Incremente de 1 a chaque début
                critereCommunMain = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(k + 8, 1).Value & Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(k + 8, 2).Value
            Wend
            GoTo a
            
        End If
    Next i
End Sub
Function ConvertToNumericDate(dateStr As String) As Long
    Dim dateParts() As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim numericDate As Long
    
    dateParts = Split(dateStr, "/")
    
    If UBound(dateParts) <> 2 Then
        ConvertToNumericDate = -1 ' Indiquer une erreur si le format n'est pas correct
        Exit Function
    End If
    
    On Error Resume Next
    dayValue = CLng(dateParts(0))
    monthValue = CLng(dateParts(1))
    yearValue = CLng(dateParts(2))
    On Error GoTo 0
    
    If Err.Number <> 0 Then
        ConvertToNumericDate = -1 ' Indiquer une erreur si la conversion échoue
    Else
        numericDate = DateSerial(yearValue, monthValue, dayValue)
        ConvertToNumericDate = CLng(numericDate)
    End If
End Function

' Fonction d'ecriture si l'employe a sa case de coché
Public Function importer_tableau(nomFichier As Variant, mon_tableau() As String, nombre_lignes As Integer, premiere_ligne As Integer, premiere_colonne As Integer) As Integer
    ' Premiere ligne / colonne dans le fichier Dassaul
    Dim i As Integer
    Dim J As Integer
    Dim nombreVrai As Integer
    Dim nombre_colonnes As Integer
    nombre_colonnes = 1
    
    While Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8, nombre_colonnes).Value <> "Dassault"
        nombre_colonnes = nombre_colonnes + 1
    Wend
    
    ' Redim le tableau en fonction des parametres
    ReDim mon_tableau(1 To nombre_lignes, 1 To nombre_colonnes)
    ' Lire tableau
    For i = 1 To nombre_lignes
        For J = 1 To nombre_colonnes
            mon_tableau(i, J) = Workbooks(nom_fichier_source).Worksheets(nom_feuille_ENG_007).Cells(8 + i, J + 0).Value
        Next J
    Next i
    
    ' Ecrire Dassault
    For i = 1 To nombre_lignes
        If mon_tableau(i, nombre_colonnes) = "Oui" Then
            nombreVrai = 1 + nombreVrai ' Incremente le nombre de fois ou j'ai trouve un employé airFrance
            Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_dassault).Cells(nombreVrai + premiere_ligne, 1).Value = mon_tableau(i, 1)
            Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_dassault).Cells(nombreVrai + premiere_ligne, 2).Value = mon_tableau(i, 2)
            Workbooks(nomFichier).Worksheets(nom_feuille_ecriture_dassault).Cells(nombreVrai + premiere_ligne, 4).Value = mon_tableau(i, 3)
        End If
    Next i
    importer_tableau = nombreVrai
End Function
Function CompterFichiersDansDossier(cheminDossier As String) As Long
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFiles As Object
    
    ' Créer un objet FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Vérifier si le dossier existe
    If objFSO.FolderExists(cheminDossier) Then
        ' Obtenir l'objet Folder associé au chemin du dossier
        Set objFolder = objFSO.GetFolder(cheminDossier)
        
        ' Obtenir la collection des fichiers dans le dossier
        Set objFiles = objFolder.Files
        
        ' Renvoyer le nombre de fichiers dans le dossier
        CompterFichiersDansDossier = objFiles.Count
    Else
        ' Si le dossier n'existe pas, renvoyer 0
        CompterFichiersDansDossier = 0
    End If
    
    ' Libérer la mémoire
    Set objFiles = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing
End Function


