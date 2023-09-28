Option Explicit
' Afficher le user form
Sub Afficher()
    Dim fermerFile As Variant
    If MsgBox("Avez-vous bien fermé et sauvegardé votre fichier Excel ?", vbYesNoCancel, "Instruction") = vbYes Then
        Envoyer.Show
    ElseIf MsgBox("Je le ferme tout de suite ?", vbYesNoCancel, "Instruction") = vbYes Then
        ' close le fichier air france
        ' close le fichier dassault
        fermerFile = FermerFichierExcelParEmplacement(empl_stockage_fichier_AFI, "Air France")  ' Ferme Air France
        
        fermerFile = FermerFichierExcelParEmplacement(emplacement_stockage_fichier_dassault, "Dassault")  ' Ferme Dassault
        
        Envoyer.Show ' Démarrer le user form
    End If
End Sub
Function FermerFichierExcelParEmplacement(ByVal emplacement As String, Destinataire As String, Optional ByVal sauvegarderAvantFermer As Boolean = True) As Boolean
    Dim excelApp As Excel.Application
    Dim wb As Excel.Workbook
    Dim fileCount As Integer
    Dim fileName As String
    Dim nbFichierDansEmplacement As Integer
    
    
    
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    On Error GoTo 0

    If excelApp Is Nothing Then
        Set excelApp = CreateObject("Excel.Application")
    End If

    ' Vérifier le nombre de fichiers Excel ouverts dans cet emplacement
    fileCount = 0
    For Each wb In excelApp.Workbooks
        If InStr(1, wb.FullName, emplacement, vbTextCompare) > 0 Then
            fileCount = fileCount + 1
            If fileCount > 1 Then
                Exit For
            End If
        End If
    Next wb
    
    If fileCount = 1 Then
        ' Fermer le seul fichier trouvé
        For Each wb In excelApp.Workbooks
            If InStr(1, wb.FullName, emplacement, vbTextCompare) > 0 Then
                If sauvegarderAvantFermer Then
                    wb.Save
                End If
                wb.Close SaveChanges:=False
                FermerFichierExcelParEmplacement = True
                MsgBox "Le fichier Excel " & Destinataire & " a bien été fermé."
            End If
        Next wb
    Else
        ' Afficher un message d'erreur s'il y a plusieurs fichiers ou aucun fichier
        If fileCount = 0 Then
            MsgBox "Aucun fichier Excel de " & Destinataire & " trouvé à l'emplacement spécifié ou fichier déjà fermé.", vbExclamation
        Else
            MsgBox "Plusieurs fichiers Excel trouvés dans le reperoire " & Destinataire & " à l'emplacement spécifié. Veuillez fermer manuellement les fichiers Excel avant d'utiliser cette fonction.", vbExclamation
        End If
        FermerFichierExcelParEmplacement = False
    End If

    ' Fermer l'application Excel si elle a été créée par cette fonction
    If Err.Number = 429 Then
        excelApp.Quit
    End If

    Set wb = Nothing
    Set excelApp = Nothing
End Function
Sub creationMail(emplacement As String, destinataireSociete As String, emplacementDestination As String)
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim cheminFichier As String
    Dim objetMail As String
    Dim Destinataire As String
    Dim corpsMail As String
    Dim copieAdresse As String
    
    
    ' Définir les parametres de l'e-mail
    ' Air France
    If destinataireSociete = "Air France" Then
        objetMail = "Fwd: Tableau des formations SUNAERO MAJ " & Date
        Destinataire = "thbaillou@airfrance.fr"
        copieAdresse = "pprevost@sunaero.com ; mlaune@sunaero.com ; lbussiere@sunaero.com ; ebononi@sunaero.com ; frcastel@airfrance.fr ; jbgiannoli@sunaero.com ; chgonord@airfrance.fr ; vagobertiere@airfrance.fr ; mazbiti@sunaero.com ; sebastien.gelez@sunaero.com"
        cheminFichier = empl_stockage_fichier_AFI
        corpsMail = "Bonjour," & vbCrLf & vbCrLf & "Ci-joint vous trouverez le fichier de ce mois."
        
    Else ' Dassault
        objetMail = "Fwd: Tableaux de formations et liste des intervenants SUNAERO "
        Destinataire = "Alexis.DACOSTA@dassault-falcon.com"
        copieAdresse = "Coralie.CARICHON@dassault-falcon.com ; mlaune@sunaero.com ; jbgiannoli@sunaero.com ; philippe.lequeu@sunaero.com ; pprevost@sunaero.com ; lbussiere@sunaero.com ; ebononi@sunaero.com ; sebastien.gelez@sunaero.com"
                        
        cheminFichier = emplacement_stockage_fichier_dassault
        corpsMail = "Bonjour," & vbCrLf & vbCrLf & "Ci-joint vous trouverez le fichier de ce mois."
        
    End If
    
    ' Vérifier si le dossier existe et contient un seul fichier xlsx
    
    Dim nomFichier As String
    
    
    nomFichier = Dir(emplacement & "*.xlsx")
    If nomFichier = "" Then
        MsgBox "Aucun fichier Excel trouvé dans l'emplacement spécifié.", vbExclamation
        Envoyer.Hide ' Fermer le userform
        Exit Sub
    End If

    Dim nbFichiers As Integer
    Do While nomFichier <> ""
        nbFichiers = nbFichiers + 1
        nomFichier = Dir
    Loop
    
    If nbFichiers > 1 Then
        MsgBox "Plusieurs fichiers xlsx trouvés dans l'emplacement spécifié. Assurez-vous qu'il n'y a qu'un seul fichier xlsx.", vbExclamation
        Exit Sub
    End If
    
    ' Fermer et save le fichier Excel si ouvert
    
    
    
    ' Fermer le UserForm
    Envoyer.Hide
     
     
    ' Créer et envoyer l'e-mail
    On Error Resume Next
    Set outlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    If outlookApp Is Nothing Then
        MsgBox "Application Outlook introuvable. Assurez-vous qu'Outlook est installé sur votre ordinateur.", vbExclamation
        Envoyer.Hide ' Fermer le userform
        Exit Sub
    End If
    
    Set outlookMail = outlookApp.CreateItem(0)
    With outlookMail
        .CC = copieAdresse
        .To = Destinataire
        .subject = objetMail
        .Body = corpsMail
        .Attachments.Add cheminFichier & Dir(cheminFichier & "*.xlsx")
        .Display ' Utilisez .Send si vous souhaitez envoyer l'e-mail directement sans afficher la fenêtre de composition
    End With
    
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    
    Call Archiver(emplacement, emplacementDestination)
    
End Sub
Sub Archiver(cheminSource As String, cheminDestination As String)
    ' Vérifiez si le répertoire source existe
    If Dir(cheminSource, vbDirectory) = "" Then
        MsgBox "Le répertoire source n'existe pas.", vbExclamation, "Répertoire source introuvable"
        Exit Sub
    End If

    ' Récupère le nom du premier fichier Excel dans le répertoire source
    Dim nomFichier As String
    nomFichier = Dir(cheminSource & "*.xlsx")

    ' Vérifiez si le fichier source existe avant de le déplacer
    If nomFichier <> "" Then
        ' Déplacez le fichier
        FileCopy cheminSource & nomFichier, cheminDestination & nomFichier
        If Dir(cheminDestination & nomFichier) <> "" Then
            Kill cheminSource & nomFichier ' Supprimez le fichier source
            MsgBox "Le fichier a été déplacé avec succès vers : " & cheminDestination, vbInformation, "Déplacement réussi"
        Else
            MsgBox "Le déplacement du fichier a échoué.", vbExclamation, "Échec du déplacement"
        End If
    Else
        ' Le fichier source n'existe pas
        MsgBox "Aucun fichier Excel trouvé dans le répertoire source.", vbExclamation, "Fichier introuvable"
    End If
End Sub
Sub test()

    Call Archiver("S:\Service EMEA\8_Formations & Certificats\VBA Formations\release\Dassault\", "S:\Service EMEA\8_Formations & Certificats\VBA Formations\release\Dassault\Envoyé\")
    

End Sub


