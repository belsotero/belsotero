Attribute VB_Name = "modPowerShell"
Option Explicit

Public Function CaminhoScript(ByVal nomeScript As String) As String
    CaminhoScript = ThisWorkbook.Path & "\..\03_PowerShell\" & nomeScript
End Function

Public Function MontarPastaSaidaExecucao() As String
    Dim baseSaida As String
    baseSaida = ObterValorConfiguracao("PastaSaida", ThisWorkbook.Path & "\..\04_Saidas")
    MontarPastaSaidaExecucao = baseSaida
End Function

Public Function ExecutarScriptPowerShell(ByVal script As String, ByVal argumentos As String, ByVal acao As String) As Long
    On Error GoTo Trata
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & script & """ " & argumentos
    RegistrarLog "modPowerShell", acao, "INICIO", cmd
    ExecutarScriptPowerShell = sh.Run(cmd, 1, True)
    If ExecutarScriptPowerShell = 0 Then
        RegistrarLog "modPowerShell", acao, "OK", "Script concluído com sucesso.", script
    Else
        RegistrarLog "modPowerShell", acao, "ERRO", "Script retornou código " & ExecutarScriptPowerShell, script
    End If
    Exit Function
Trata:
    TratarErro "modPowerShell", acao, Err.Number, Err.Description, script
    ExecutarScriptPowerShell = -1
End Function

Public Sub ExecutarInventarioPowerShell()
    Dim pasta As String: pasta = ObterValorConfiguracao("PastaAnalisada")
    If Len(pasta) = 0 Then MsgBox "Escolha uma pasta antes de executar o inventário.", vbExclamation: Exit Sub
    Dim args As String
    args = ArgsComuns(pasta)
    ExecutarScriptPowerShell CaminhoScript("InventarioArquivos.ps1"), args, "Inventario"
    ExecutarScriptPowerShell CaminhoScript("DiagnosticoArquivos.ps1"), args, "Diagnostico"
    ImportarCSVsRecentes
End Sub

Public Sub ExecutarExtracaoPowerShell()
    Dim pasta As String: pasta = ObterValorConfiguracao("PastaAnalisada")
    If Len(pasta) = 0 Then MsgBox "Escolha uma pasta antes de extrair conteúdo.", vbExclamation: Exit Sub
    ExecutarScriptPowerShell CaminhoScript("ExtrairConteudo.ps1"), ArgsComuns(pasta), "Extracao"
    ImportarCSVsRecentes
End Sub

Public Sub ExecutarDesbloqueioSeguro()
    Dim pasta As String: pasta = ObterValorConfiguracao("PastaAnalisada")
    If Len(pasta) = 0 Then MsgBox "Escolha uma pasta antes de verificar bloqueios.", vbExclamation: Exit Sub
    If MsgBox("A rotina apenas lista e desbloqueia arquivos confiáveis dentro da pasta escolhida. Deseja continuar para a etapa de confirmação do PowerShell?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    ExecutarScriptPowerShell CaminhoScript("DesbloquearArquivosConfiaveis.ps1"), ArgsComuns(pasta) & " -Confirmar", "DesbloqueioSeguro"
End Sub

Public Function ArgsComuns(ByVal pasta As String) As String
    ArgsComuns = "-PastaRaiz """ & pasta & """" & _
                 " -PastaSaida """ & MontarPastaSaidaExecucao() & """" & _
                 " -AnalisarSubpastas " & IIf(UCase$(ObterValorConfiguracao("AnalisarSubpastas", "Sim")) = "SIM", "$true", "$false") & _
                 " -IgnorarOcultos " & IIf(UCase$(ObterValorConfiguracao("IgnorarOcultos", "Sim")) = "SIM", "$true", "$false") & _
                 " -IgnorarTemporarios " & IIf(UCase$(ObterValorConfiguracao("IgnorarTemporarios", "Sim")) = "SIM", "$true", "$false") & _
                 " -TamanhoMaximoExtracaoMB " & ObterValorConfiguracao("TamanhoMaximoExtracaoMB", "25") & _
                 " -Tipos """ & ObterValorConfiguracao("TiposArquivos", ".txt,.csv,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.pdf,.msg") & """"
End Function
