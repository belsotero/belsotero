Attribute VB_Name = "modTratamentoErros"
Option Explicit

Public Const ABA_LOG As String = "LOG_EXECUCAO"

Public Sub RegistrarLog(ByVal modulo As String, ByVal acao As String, ByVal status As String, ByVal mensagem As String, Optional ByVal caminho As String = "", Optional ByVal detalheTecnico As String = "")
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(ABA_LOG)
    Dim linha As Long
    linha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(linha, 1).Value = Now
    ws.Cells(linha, 2).Value = modulo
    ws.Cells(linha, 3).Value = acao
    ws.Cells(linha, 4).Value = status
    ws.Cells(linha, 5).Value = mensagem
    ws.Cells(linha, 6).Value = caminho
    ws.Cells(linha, 7).Value = detalheTecnico
End Sub

Public Sub TratarErro(ByVal modulo As String, ByVal acao As String, ByVal numeroErro As Long, ByVal descricaoErro As String, Optional ByVal caminho As String = "")
    RegistrarLog modulo, acao, "ERRO", MensagemAmigavelErro(numeroErro, descricaoErro), caminho, CStr(numeroErro) & " - " & descricaoErro
    MsgBox MensagemAmigavelErro(numeroErro, descricaoErro), vbExclamation, "Central Inteligente de Arquivos"
End Sub

Public Function MensagemAmigavelErro(ByVal numeroErro As Long, ByVal descricaoErro As String) As String
    Select Case numeroErro
        Case 53
            MensagemAmigavelErro = "Arquivo ou script não encontrado. Verifique se a estrutura do projeto está completa."
        Case 70
            MensagemAmigavelErro = "Permissão negada. Verifique se o arquivo está aberto, protegido ou bloqueado por política corporativa."
        Case 76
            MensagemAmigavelErro = "Caminho não encontrado. Escolha novamente a pasta de origem ou saída."
        Case Else
            MensagemAmigavelErro = "A operação não pôde ser concluída. Detalhe técnico: " & descricaoErro
    End Select
End Function
