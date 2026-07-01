Attribute VB_Name = "modPrincipal"
Option Explicit

Public Sub InicializarCentral()
    On Error GoTo Trata
    Application.ScreenUpdating = False
    CriarAbasPadrao
    PrepararCabecalhos
    PrepararConfiguracoesPadrao
    PrepararPainel
    RegistrarLog "modPrincipal", "InicializarCentral", "OK", "Central inicializada."
    Application.ScreenUpdating = True
    MsgBox "Central Inteligente de Arquivos inicializada com sucesso.", vbInformation
    Exit Sub
Trata:
    Application.ScreenUpdating = True
    TratarErro "modPrincipal", "InicializarCentral", Err.Number, Err.Description
End Sub

Private Sub CriarAbasPadrao()
    Dim abas As Variant: abas = Array("PAINEL", "CONFIGURACOES", "INVENTARIO_ARQUIVOS", "CONTEUDO_EXTRAIDO", "RESULTADOS_BUSCA", "TRECHOS_SELECIONADOS", "DUPLICADOS", "ALERTAS", "LOG_EXECUCAO", "ESTATISTICAS")
    Dim i As Long
    For i = LBound(abas) To UBound(abas)
        If Not ExisteAba(CStr(abas(i))) Then ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = CStr(abas(i))
    Next i
End Sub

Private Function ExisteAba(ByVal nome As String) As Boolean
    On Error Resume Next
    ExisteAba = Not ThisWorkbook.Worksheets(nome) Is Nothing
End Function

Private Sub PrepararCabecalhos()
    DefinirCabecalho "CONFIGURACOES", Array("Chave", "Valor", "Descrição")
    DefinirCabecalho "INVENTARIO_ARQUIVOS", Array("ID", "Nome", "Extensão", "Tipo", "Caminho completo", "Pasta de origem", "Tamanho KB", "Tamanho MB", "Data criação", "Data modificação", "Status leitura", "Status extração", "Hash SHA256", "Duplicado nome", "Duplicado tamanho", "Duplicado hash", "Alerta", "Observações", "Abrir arquivo", "Abrir pasta")
    DefinirCabecalho "CONTEUDO_EXTRAIDO", Array("ID arquivo", "Nome arquivo", "Caminho", "Tipo", "Página/Aba/Slide/Seção", "Trecho extraído", "Palavra-chave", "Data extração", "Status", "Observações erro")
    DefinirCabecalho "RESULTADOS_BUSCA", Array("Palavra-chave", "Arquivo", "Tipo", "Caminho", "Pasta origem", "Trecho encontrado", "Data modificação", "Abrir arquivo", "Abrir pasta")
    DefinirCabecalho "TRECHOS_SELECIONADOS", Array("Selecionado", "Palavra-chave", "Arquivo", "Tipo", "Caminho", "Trecho", "Categoria", "Observação", "Data seleção", "Exportado")
    DefinirCabecalho "DUPLICADOS", Array("Critério", "Hash SHA256", "Nome", "Tamanho", "Caminho", "Pasta origem", "Data modificação", "Observação")
    DefinirCabecalho "ALERTAS", Array("ID", "Tipo alerta", "Severidade", "Arquivo", "Caminho", "Detalhe", "Ação recomendada", "Data/hora")
    DefinirCabecalho "LOG_EXECUCAO", Array("Data/hora", "Módulo", "Ação", "Status", "Mensagem", "Caminho relacionado", "Detalhe técnico")
    DefinirCabecalho "ESTATISTICAS", Array("Indicador", "Valor", "Observação")
End Sub

Private Sub DefinirCabecalho(ByVal aba As String, ByVal campos As Variant)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(aba)
    Dim i As Long
    For i = LBound(campos) To UBound(campos)
        ws.Cells(1, i + 1).Value = campos(i)
        ws.Cells(1, i + 1).Font.Bold = True
    Next i
    ws.Columns.AutoFit
End Sub

Private Sub PrepararConfiguracoesPadrao()
    If Len(ObterValorConfiguracao("AnalisarSubpastas")) = 0 Then DefinirValorConfiguracao "AnalisarSubpastas", "Sim"
    If Len(ObterValorConfiguracao("IgnorarOcultos")) = 0 Then DefinirValorConfiguracao "IgnorarOcultos", "Sim"
    If Len(ObterValorConfiguracao("IgnorarTemporarios")) = 0 Then DefinirValorConfiguracao "IgnorarTemporarios", "Sim"
    If Len(ObterValorConfiguracao("TamanhoMaximoExtracaoMB")) = 0 Then DefinirValorConfiguracao "TamanhoMaximoExtracaoMB", "25"
    If Len(ObterValorConfiguracao("TiposArquivos")) = 0 Then DefinirValorConfiguracao "TiposArquivos", ".txt,.csv,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.pdf,.msg"
    If Len(ObterValorConfiguracao("PastaSaida")) = 0 Then DefinirValorConfiguracao "PastaSaida", ThisWorkbook.Path & "\..\04_Saidas"
End Sub

Public Sub ExecutarInventario()
    ExecutarInventarioPowerShell
    AtualizarResumoPainel
End Sub

Public Sub AtualizarDados()
    ImportarCSVsRecentes
    AtualizarResumoPainel
End Sub

Public Sub ExtrairConteudo()
    ExecutarExtracaoPowerShell
    AtualizarResumoPainel
End Sub

Public Sub PesquisarPalavraChave()
    PesquisarConteudoExtraido
End Sub

Public Sub LimparResultados()
    If MsgBox("Limpar resultados importados no Excel? Os arquivos originais não serão alterados.", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    LimparAba "RESULTADOS_BUSCA"
    LimparAba "TRECHOS_SELECIONADOS"
    RegistrarLog "modPrincipal", "LimparResultados", "OK", "Resultados limpos no Excel."
End Sub

Private Sub LimparAba(ByVal nomeAba As String)
    With ThisWorkbook.Worksheets(nomeAba)
        If .Rows.Count > 1 Then .Rows("2:" & .Rows.Count).ClearContents
    End With
End Sub
