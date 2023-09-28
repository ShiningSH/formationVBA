Sub OuvrirFichierAudio()
    Dim cheminMusique As String
    cheminMusique = "S:\Service EMEA\8_Formations & Certificats\VBA Formations\release\old\Image\easteregg.mp3"
    Shell "explorer.exe """ & cheminMusique & """", vbNormalFocus
    Application.SendKeys "^%{DOWN}"
    Application.Wait Now + TimeValue("00:00:6")
    Call Central_Park_Contol_Console.Show
End Sub

