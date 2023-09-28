Option Explicit

'*----------------------------------------------------------------------------------------------------------*

'       ___ ___ ___  __   ___                  __        __      __   __             ___  __
' |\/| |__   |   |  |__) |__      /\        | /  \ |  | |__)    |  \ /  \ |\ | |\ | |__  /__`    .
' |  | |___  |   |  |  \ |___    /~~\    \__/ \__/ \__/ |  \    |__/ \__/ | \| | \| |___ .__/    .
                                                                                                 
                                                                                                 
'*----------------------------------------------------------------------------------------------------------*
Sub ouvrirLogin()
    
    Call Login.Show

End Sub
' Fichier Formation Source :


Function nom_fichier_source() As String
    nom_fichier_source = ThisWorkbook.Name
End Function
' Nom de la feuille commune de toutes les formations
Function nom_feuille_ENG_007() As String
    ' Feuille située a la 2 eme position
    Dim sheetName As String
    On Error Resume Next
    sheetName = ThisWorkbook.Sheets(2).Name
    On Error GoTo 0
    
    ' Attribuer sa valeur
    nom_feuille_ENG_007 = sheetName
End Function
' Nom de la feuille stockant la matrice des formations obligatoires
Function nom_feuille_M1065() As String
    ' Feuille située a la 3 eme position
    Dim sheetName As String
    On Error Resume Next
    sheetName = ThisWorkbook.Sheets(3).Name
    On Error GoTo 0
    
    ' Attribuer sa valeur
    nom_feuille_M1065 = sheetName
End Function
Function nom_feuille_matrices() As String
    ' Feuille située a la 4 eme position
    Dim sheetName As String
    On Error Resume Next
    sheetName = ThisWorkbook.Sheets(4).Name
    On Error GoTo 0
    
    ' Attribuer sa valeur
    nom_feuille_matrices = sheetName
End Function
Function nom_feuille_datas_validation() As String
    ' Feuille située a la 5 eme position
    Dim sheetName As String
    On Error Resume Next
    sheetName = ThisWorkbook.Sheets(5).Name
    On Error GoTo 0
    
    ' Attribuer sa valeur
    nom_feuille_datas_validation = sheetName
End Function
Function passwordSheet()
    passwordSheet = 12
End Function
Function plage_donnees() As String
    plage_donnees = "H9:BJ119" ' Spécifier la plage de données sur laquelle on marque les dates
End Function


' XXXXX :

' Emplacement du template AFI
Function emplacement_templateAFI() As String
    ' Retourner l'emplacement du fichier
    emplacement_templateAFI = "S:\Service EMEA\8_Formations & Certificats\VBA Formations\release\Air France\Template\"
End Function
' Emplacement temporaire ou est stocké le fichier en cours de création
Function empl_stockage_fichier_AFI() As String
    empl_stockage_fichier_AFI = "S:\Service EMEA\8_Formations & Certificats\VBA Formations\release\Air France\"
End Function
' Emplacement des fichiers AFI envoyés
Function empl_stockage_envoyAFI() As String
    empl_stockage_envoyAFI = "S:\Service EMEA\8_Formations & Certificats\VBA Formations\release\Air France\Envoyé\"
End Function
' Nom de la feuille du fichier AFI
Function nom_feuille_ecriture_AFI() As String
    nom_feuille_ecriture_AFI = "SUNAERO"
End Function



' XXXXX :

' Emplacement des envoyés
Function empl_stockage_envoyDassault() As String
    empl_stockage_envoyDassault = "S:\Service EMEA\8_Formations & Certificats\VBA Formations\release\Dassault\Envoyé\"
End Function

Function nom_feuille_ecriture_dassault() As String
    nom_feuille_ecriture_dassault = "TDCR5"
End Function
Function emplacement_templateDassault() As String
    emplacement_templateDassault = "S:\Service EMEA\8_Formations & Certificats\VBA Formations\release\Dassault\Template\"
End Function
' Emplacement du fichier généré en cours de modification
Function emplacement_stockage_fichier_dassault() As String
    emplacement_stockage_fichier_dassault = "S:\Service EMEA\8_Formations & Certificats\VBA Formations\release\Dassault\"
End Function


'-----------------------------------------------------------------------------------------------------------*
'//                                                                                                         |
'//                                                                                                         |
'                                                                                                           |
'                                                                                                           |
'                                                                                                           |
'                                                                         !J?7~~?PBBJ                       |
'                                                  ...:::...          :!^.^JBBB##J^.                        |
'                                           .^!?J5PPGGGBBGGGP5YJ7~:.   JBBPPJ!P#J                           |
'                                       .~?5GBBBBBBBBBBBBBBBBB####BG57: ^7.   GY                            |
'                                    .!YGBBBBBBBBBBBBBBBBBBBB###BG5?7!!7?~.  .^                             |
'                                  ^JPBBGGGGGGGGGGBBBBBBBBBBPY?!^~!?5GB##BP?:                               |
'                                ^YGGGGGGGGGGGGGGGGBBBGPY7~^^~7YPB##BBBBBBB#GJ:                             |
'                              .?GGGGGGGGGGGGGGGGBGPJ!: .~JPGB#BBBBBBBBBBBBBB#G7                            |
'                             ^5GGGGGGGGGGGGGGGGPJ~.  ~YGBBBBBBBBBBBBBBBBBBBBBB#Y.                          |
'                            ~PGPPPPPPPPPPPPPGG?:   ~5BBBGGGGGGBBBBBBBBBBBBBBBBB#P:                         |
'                           ^PPPPPPPPPPPPPPPPP^    !GBGGGGGGGGGGGGGGGGBBBBBBBBBBB#P.                        |
'                          .5PPPPPPPPPPPPPPPP~     YBGGGGGGGGGGGGGGGGGGGGGBBBBBBBB#Y                        |
'                          7P5555555PPPPPPPPP^     ^5GGGGGGGGGGGGGGGGGGGGGGGGBBBBBBB:                       |
'                         .Y555555555555555PPY:     .~?Y5PPPGGGGGGGGGGGGGGGGGGGGBBBB::                      |
'                         :555555555555555555P5?^.       .:^^~!7?YPGGGGGGGGGGGGGGGBBP:                      |
'                         :55555555555555555555PP5YJ?77!!~^:      .~?PGPPGGGGGGGGGGBP.                      |
'                         :YYYYYYYYYYY555555555555PPPPPPPPPP5?~      :JPPPPGGGGGGGGBP:                      |
'                          5 PPPPPPPPGGY:7777777777777777777777        JYYYYYYYYYYYYY!                      |
'                          !YYYYYYYYYYYYYYYY5555555555PPPPPPPPPG5.     ^PPPPPPGGGGGG|.                      |
'                          .Print.Print YJJJJYYYYYYYYYYYY555555555     PPPPG^: PPPPP:                       |
'                           :JJJJJJJJJYYYYYYYYY555555555PPPPPPPGP.     7PPPPPPPPGGY.                        |
'                            :?JJJJJJJJJYYYYYYYY555555555PPPPPPP!     ~555PPPPPPGY.                         |
'                             .7JJJJJJJJJYYYYYYYY55555555PPPPG5~     !5555PPPPPG?.                          |
'                               ~?JJJJJJJJYYYYYYY5555555PPPP57.    ^J555555PPP5~                            |
'                                .!?JJJJJJYYYYYYY555555PP5?~.   .^?5555555PP57.                             |
'                                  .~?JYYYYYYYYYY55555J!^.   .^!JYYY5555P5Y!.                               |
'                                    .^!JYYYYYY555J7^.    .~7JYYYYY5555J7^                                  |
'                                       .:~7JYJ7~:    .^!?JYYYYYYYYJ7~:                                     |
'                                           ...    .~7JJYJJJJ?7!~^.                                         |
'                                                                                                           |
'                                                                                                           |
'                             _____    __  __   _   __   ___       ______   ____     ____                   |
'                            / ___/   / / / /  / | / /  /   |     / ____/  / __ \   / __ \                  |
'                            \__ \   / / / /  /  |/ /  / /| |    / __/    / /_/ /  / / / /                  |
'                           ___/ /  / /_/ /  / /|  /  / ___ |   / /___   / _, _/  / /_/ /                   |
'                          /____/   \____/  /_/ |_/  /_/  |_|  /_____/  /_/ |_|   \____/                    |
'                                                                                                           |
'                                                                                                           |
'                                                                                                           |
'                                                                                                           |
'*----------------------------------------------------------------------------------------------------------*
'//                             Auteur       : XXXXX                                                        |
'//                             Description  : TABLEAU DES FORMATIONS                                       |
'//                             Date         : 01/08/2023                                                   |
'//                             Version:       Version 1                                                    |
'*----------------------------------------------------------------------------------------------------------*
                                   
        




















