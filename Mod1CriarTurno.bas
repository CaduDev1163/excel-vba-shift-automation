Attribute VB_Name = "Mod1CriarTurno"
Option Explicit

Public Sub CriarTurno()

    Dim wsModelo As Worksheet
    Dim wsNova As Worksheet
    Dim wsLog As Worksheet
    Dim NomePlanilha As String
    Dim Usuario As String
    Dim UltimaLinha As Long
    
    ' Definir nome da nova planilha (data atual)
   NomePlanilha = Format(Now, "dd.mm._hh.mm")
    
    ' Referências
    Set wsModelo = ThisWorkbook.Sheets("MODELO_TURNO")
    Set wsLog = ThisWorkbook.Sheets("LOG")
    
    ' Verificar se já existe planilha com essa data
    If PlanilhaExiste(NomePlanilha) Then
        MsgBox "O turno de hoje já foi criado!", vbExclamation
        Exit Sub
    End If
    
    ' Copiar modelo
    wsModelo.Copy After:=Sheets(Sheets.Count)
    Set wsNova = ActiveSheet
    wsNova.Name = NomePlanilha
    
    ' Registrar Log
    Usuario = Environ("Username")
    
    UltimaLinha = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    
    wsLog.Cells(UltimaLinha, 1).Value = Date
    wsLog.Cells(UltimaLinha, 2).Value = Usuario
    wsLog.Cells(UltimaLinha, 3).Value = Date
    wsLog.Cells(UltimaLinha, 4).Value = Time
    wsLog.Cells(UltimaLinha, 5).Value = Environ("ComputerName")
    wsLog.Cells(UltimaLinha, 6).Value = ThisWorkbook.Name
    
    MsgBox "Turno criado com sucesso!", vbInformation
    
    Sheets("LOG").Visible = xlSheetVeryHidden
    Sheets("MODELO_TURNO").Visible = xlSheetVeryHidden
    
    'MsgBox "Planilhas administrativas ocultadas.", vbInformation

End Sub

Private Function PlanilhaExiste(Nome As String) As Boolean

    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Nome)
    On Error GoTo 0
    
    PlanilhaExiste = Not ws Is Nothing

End Function


