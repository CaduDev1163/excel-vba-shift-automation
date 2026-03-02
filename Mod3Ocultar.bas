Attribute VB_Name = "Mod3Ocultar"
Sub OcultarEstrutura()

    Sheets("LOG").Visible = xlSheetVeryHidden
    Sheets("MODELO_TURNO").Visible = xlSheetVeryHidden
    
    MsgBox "Planilhas administrativas ocultadas.", vbInformation

End Sub
