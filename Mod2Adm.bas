Attribute VB_Name = "Mod2Adm"
Sub AcessoAdministrador()

    Dim Senha As String
    
    Senha = InputBox("Digite a senha de administrador:")
    
    If Senha = "1234" Then
        
        Sheets("LOG").Visible = xlSheetVisible
        Sheets("MODELO_TURNO").Visible = xlSheetVisible
        
        MsgBox "Acesso liberado.", vbInformation
        
    Else
        
        MsgBox "Senha incorreta!", vbCritical
        
    End If

End Sub
