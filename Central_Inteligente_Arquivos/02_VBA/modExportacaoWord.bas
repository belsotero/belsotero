Attribute VB_Name = "modExportacaoWord"
Option Explicit

Public Sub ExportarTrechosParaWord()
    On Error GoTo Trata
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("TRECHOS_SELECIONADOS")
    Dim app As Object: Set app = CreateObject("Word.Application")
    Dim doc As Object: Set doc = app.Documents.Add
    app.Visible = True
    doc.Content.InsertAfter "Central Inteligente de Arquivos" & vbCrLf
    doc.Content.InsertAfter "Exportação: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    doc.Content.InsertAfter "Pasta analisada: " & ObterValorConfiguracao("PastaAnalisada") & vbCrLf & vbCrLf
    doc.Content.InsertAfter "Observação de rastreabilidade: os trechos abaixo referenciam os caminhos originais. Nenhum arquivo original foi alterado." & vbCrLf & vbCrLf
    Dim linha As Long
    For linha = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If UCase$(CStr(ws.Cells(linha, 1).Value)) = "SIM" Or CStr(ws.Cells(linha, 1).Value) = "1" Or CBoolSafe(ws.Cells(linha, 1).Value) Then
            doc.Content.InsertAfter "Arquivo: " & ws.Cells(linha, 3).Value & vbCrLf
            doc.Content.InsertAfter "Tipo: " & ws.Cells(linha, 4).Value & vbCrLf
            doc.Content.InsertAfter "Palavra-chave: " & ws.Cells(linha, 2).Value & vbCrLf
            doc.Content.InsertAfter "Caminho: " & ws.Cells(linha, 5).Value & vbCrLf
            doc.Content.InsertAfter "Trecho: " & ws.Cells(linha, 6).Value & vbCrLf & vbCrLf
        End If
    Next linha
    Dim saida As String
    saida = ObterValorConfiguracao("PastaSaida", ThisWorkbook.Path & "\..\04_Saidas") & "\Word\trechos_" & Format(Now, "yyyymmdd_hhnnss") & ".docx"
    doc.SaveAs2 saida
    RegistrarLog "modExportacaoWord", "ExportarTrechosParaWord", "OK", "Documento Word gerado.", saida
    Exit Sub
Trata:
    TratarErro "modExportacaoWord", "ExportarTrechosParaWord", Err.Number, Err.Description
End Sub

Private Function CBoolSafe(ByVal valor As Variant) As Boolean
    On Error Resume Next
    CBoolSafe = CBool(valor)
End Function
