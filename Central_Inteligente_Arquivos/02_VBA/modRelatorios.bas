Attribute VB_Name = "modRelatorios"
Option Explicit

Public Sub ImportarCSVsRecentes()
    On Error GoTo Trata
    ImportarCSVSeExistir ArquivoMaisRecente("inventario_*.csv"), "INVENTARIO_ARQUIVOS"
    ImportarCSVSeExistir ArquivoMaisRecente("conteudo_extraido_*.csv"), "CONTEUDO_EXTRAIDO"
    ImportarCSVSeExistir ArquivoMaisRecente("duplicados_*.csv"), "DUPLICADOS"
    ImportarCSVSeExistir ArquivoMaisRecente("alertas_*.csv"), "ALERTAS"
    RegistrarLog "modRelatorios", "ImportarCSVsRecentes", "OK", "CSVs recentes importados quando encontrados."
    Exit Sub
Trata:
    TratarErro "modRelatorios", "ImportarCSVsRecentes", Err.Number, Err.Description
End Sub

Private Function PastaCSV() As String
    PastaCSV = ObterValorConfiguracao("PastaSaida", ThisWorkbook.Path & "\..\04_Saidas") & "\CSV\"
End Function

Private Function ArquivoMaisRecente(ByVal padrao As String) As String
    Dim nome As String, melhor As String, dt As Date
    nome = Dir(PastaCSV() & padrao)
    Do While Len(nome) > 0
        If FileDateTime(PastaCSV() & nome) > dt Then
            dt = FileDateTime(PastaCSV() & nome)
            melhor = PastaCSV() & nome
        End If
        nome = Dir
    Loop
    ArquivoMaisRecente = melhor
End Function

Private Sub ImportarCSVSeExistir(ByVal arquivo As String, ByVal aba As String)
    If Len(arquivo) = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(aba)
    ws.Rows("2:" & ws.Rows.Count).ClearContents
    With ws.QueryTables.Add(Connection:="TEXT;" & arquivo, Destination:=ws.Range("A2"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFilePlatform = 65001
        .TextFileColumnDataTypes = Array(1)
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    ws.Rows(2).Delete
    ws.Columns.AutoFit
End Sub

Public Sub PesquisarConteudoExtraido()
    On Error GoTo Trata
    Dim termo As String: termo = InputBox("Digite a palavra-chave para pesquisar:", "Pesquisar conteúdo")
    If Len(Trim$(termo)) = 0 Then Exit Sub
    Dim origem As Worksheet: Set origem = ThisWorkbook.Worksheets("CONTEUDO_EXTRAIDO")
    Dim destino As Worksheet: Set destino = ThisWorkbook.Worksheets("RESULTADOS_BUSCA")
    destino.Rows("2:" & destino.Rows.Count).ClearContents
    Dim linha As Long, out As Long: out = 2
    For linha = 2 To origem.Cells(origem.Rows.Count, 1).End(xlUp).Row
        If InStr(1, CStr(origem.Cells(linha, 6).Value), termo, vbTextCompare) > 0 Then
            destino.Cells(out, 1).Value = termo
            destino.Cells(out, 2).Value = origem.Cells(linha, 2).Value
            destino.Cells(out, 3).Value = origem.Cells(linha, 4).Value
            destino.Cells(out, 4).Value = origem.Cells(linha, 3).Value
            destino.Cells(out, 5).Value = Left$(CStr(origem.Cells(linha, 3).Value), InStrRev(CStr(origem.Cells(linha, 3).Value), "\") - 1)
            destino.Cells(out, 6).Value = origem.Cells(linha, 6).Value
            destino.Cells(out, 7).Value = ""
            destino.Hyperlinks.Add destino.Cells(out, 8), CStr(origem.Cells(linha, 3).Value), , , "Abrir arquivo"
            destino.Hyperlinks.Add destino.Cells(out, 9), destino.Cells(out, 5).Value, , , "Abrir pasta"
            out = out + 1
        End If
    Next linha
    RegistrarLog "modRelatorios", "PesquisarConteudoExtraido", "OK", "Pesquisa concluída: " & termo
    destino.Activate
    Exit Sub
Trata:
    TratarErro "modRelatorios", "PesquisarConteudoExtraido", Err.Number, Err.Description
End Sub

Public Sub ExportarTrechosParaExcel()
    On Error GoTo Trata
    ThisWorkbook.Worksheets("TRECHOS_SELECIONADOS").Copy
    ActiveWorkbook.SaveAs ObterValorConfiguracao("PastaSaida", ThisWorkbook.Path & "\..\04_Saidas") & "\Excel\trechos_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx", xlOpenXMLWorkbook
    ActiveWorkbook.Close False
    RegistrarLog "modRelatorios", "ExportarTrechosParaExcel", "OK", "Trechos exportados para Excel."
    Exit Sub
Trata:
    TratarErro "modRelatorios", "ExportarTrechosParaExcel", Err.Number, Err.Description
End Sub

Public Sub GerarRelatorioHTML()
    Dim pasta As String: pasta = ObterValorConfiguracao("PastaAnalisada")
    ExecutarScriptPowerShell CaminhoScript("GerarRelatorios.ps1"), ArgsComuns(pasta), "GerarRelatorioHTML"
End Sub
