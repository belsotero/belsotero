Attribute VB_Name = "modInterface"
Option Explicit

Public Sub PrepararPainel()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("PAINEL")
    ws.Cells.Clear
    ws.Range("A1").Value = "Central Inteligente de Arquivos"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "Pasta analisada"
    ws.Range("A4").Value = "Pasta de saída"
    ws.Range("A5").Value = "Última atualização"
    CriarBotoesPainel
    AtualizarResumoPainel
End Sub

Public Sub CriarBotoesPainel()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("PAINEL")
    Dim nomes As Variant, macros As Variant
    nomes = Array("Escolher pasta", "Executar inventário", "Atualizar dados", "Extrair conteúdo", "Pesquisar palavra-chave", "Visualizar arquivo selecionado", "Abrir arquivo original", "Abrir pasta de origem", "Exportar trechos para Word", "Exportar trechos para Excel", "Gerar relatório HTML", "Limpar resultados", "Abrir pasta de saída", "Ver logs", "Desbloquear arquivos confiáveis")
    macros = Array("EscolherPasta", "ExecutarInventario", "AtualizarDados", "ExtrairConteudo", "PesquisarPalavraChave", "VisualizarArquivoSelecionado", "AbrirArquivoSelecionado", "AbrirPastaDeOrigemSelecionada", "ExportarTrechosParaWord", "ExportarTrechosParaExcel", "GerarRelatorioHTML", "LimparResultados", "AbrirPastaSaida", "VerLogs", "ExecutarDesbloqueioSeguro")
    Dim i As Long, btn As Button
    ws.Buttons.Delete
    For i = LBound(nomes) To UBound(nomes)
        Set btn = ws.Buttons.Add(20 + (i Mod 3) * 180, 110 + Int(i / 3) * 32, 165, 24)
        btn.Caption = nomes(i)
        btn.OnAction = macros(i)
    Next i
End Sub

Public Sub AtualizarResumoPainel()
    On Error Resume Next
    With ThisWorkbook.Worksheets("PAINEL")
        .Range("B3").Value = ObterValorConfiguracao("PastaAnalisada")
        .Range("B4").Value = ObterValorConfiguracao("PastaSaida")
        .Range("B5").Value = Now
        .Columns("A:B").AutoFit
    End With
End Sub

Public Sub VisualizarArquivoSelecionado()
    On Error GoTo Trata
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("PAINEL")
    Dim r As Range: Set r = ActiveCell.EntireRow
    ws.Range("A18:B26").ClearContents
    ws.Range("A18").Value = "Visualização rápida"
    ws.Range("A19").Value = "Nome": ws.Range("B19").Value = r.Cells(1, 2).Value
    ws.Range("A20").Value = "Tipo": ws.Range("B20").Value = r.Cells(1, 4).Value
    ws.Range("A21").Value = "Tamanho MB": ws.Range("B21").Value = r.Cells(1, 8).Value
    ws.Range("A22").Value = "Data modificação": ws.Range("B22").Value = r.Cells(1, 10).Value
    ws.Range("A23").Value = "Caminho": ws.Range("B23").Value = r.Cells(1, 5).Value
    ws.Range("A24").Value = "Status": ws.Range("B24").Value = r.Cells(1, 11).Value & " / " & r.Cells(1, 12).Value
    ws.Range("A25").Value = "Observações": ws.Range("B25").Value = r.Cells(1, 18).Value
    Exit Sub
Trata:
    TratarErro "modInterface", "VisualizarArquivoSelecionado", Err.Number, Err.Description
End Sub

Public Sub AbrirArquivoSelecionado()
    On Error GoTo Trata
    Dim caminho As String: caminho = ActiveCell.EntireRow.Cells(1, 5).Value
    If Len(caminho) > 0 Then ThisWorkbook.FollowHyperlink caminho
    Exit Sub
Trata:
    TratarErro "modInterface", "AbrirArquivoSelecionado", Err.Number, Err.Description
End Sub

Public Sub VerLogs()
    ThisWorkbook.Worksheets("LOG_EXECUCAO").Activate
End Sub
