Option Explicit

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    
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

    Dim imagePath As String
    imagePath = "S:\Service EMEA\8_Formations & Certificats\VBA Formations\release\old\Image\ref.gif"
    
    ' Charger l'image dans le contrôle Image1
    Image1.Picture = LoadPicture(imagePath)
    
    
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        AppActivate "Lecteur multimédia"
        Application.SendKeys "%{F4}"
    End If
End Sub

