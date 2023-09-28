Private Sub Workbook_Open()
    ' Reglage du zoom de la page
    Application.ScreenUpdating = False
    ActiveWindow.Zoom = 115
    ' Cacher les pages admin
    Dim data_validation As Worksheet
    Dim tri_des_matrices As Worksheet
    Set data_validation = Workbooks(nom_fichier_source).Worksheets(nom_feuille_datas_validation)
    data_validation.Visible = xlSheetVeryHidden
    Set tri_des_matrices = Worksheets(nom_feuille_matrices)
    tri_des_matrices.Visible = xlSheetVeryHidden
    ' Verouiller les feuilles
    Sheets(nom_feuille_ENG_007).Select
    ' Mettre en gris les cellules vides
    Call reset
    ' Mettre a jour acces mission
    ' Call acces_mission
    Application.ScreenUpdating = True
End Sub

    
    
    
    
    
    
    


