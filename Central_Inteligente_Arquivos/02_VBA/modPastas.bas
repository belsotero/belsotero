Attribute VB_Name = "modPastas"
Option Explicit

Public Function ObterValorConfiguracao(ByVal chave As String, Optional ByVal padrao As String = "") As String
    On Error GoTo Fim
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("CONFIGURACOES")
    Dim achado As Range
    Set achado = ws.Columns(1).Find(What:=chave, LookAt:=xlWhole, MatchCase:=False)
    If Not achado Is Nothing Then
        ObterValorConfiguracao = CStr(achado.Offset(0, 1).Value)
        Exit Function
    End If
Fim:
    ObterValorConfiguracao = padrao
End Function

Public Sub DefinirValorConfiguracao(ByVal chave As String, ByVal valor As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("CONFIGURACOES")
    Dim achado As Range
    Set achado = ws.Columns(1).Find(What:=chave, LookAt:=xlWhole, MatchCase:=False)
    If achado Is Nothing Then
        Dim linha As Long: linha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(linha, 1).Value = chave
        ws.Cells(linha, 2).Value = valor
    Else
        achado.Offset(0, 1).Value = valor
    End If
End Sub

Public Sub EscolherPasta()
    On Error GoTo Trata
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "Escolha a pasta para analisar"
    dlg.AllowMultiSelect = False
    If dlg.Show = -1 Then
        DefinirValorConfiguracao "PastaAnalisada", dlg.SelectedItems(1)
        RegistrarLog "modPastas", "EscolherPasta", "OK", "Pasta escolhida pela usuária.", dlg.SelectedItems(1)
        AtualizarResumoPainel
    End If
    Exit Sub
Trata:
    TratarErro "modPastas", "EscolherPasta", Err.Number, Err.Description
End Sub

Public Sub EscolherPastaSaida()
    On Error GoTo Trata
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "Escolha a pasta de saída dos relatórios"
    dlg.AllowMultiSelect = False
    If dlg.Show = -1 Then
        DefinirValorConfiguracao "PastaSaida", dlg.SelectedItems(1)
        RegistrarLog "modPastas", "EscolherPastaSaida", "OK", "Pasta de saída escolhida.", dlg.SelectedItems(1)
        AtualizarResumoPainel
    End If
    Exit Sub
Trata:
    TratarErro "modPastas", "EscolherPastaSaida", Err.Number, Err.Description
End Sub

Public Sub AbrirPastaSaida()
    AbrirCaminho ObterValorConfiguracao("PastaSaida", ThisWorkbook.Path & "\..\04_Saidas")
End Sub

Public Sub AbrirPastaDeOrigemSelecionada()
    On Error GoTo Trata
    Dim caminho As String: caminho = ActiveCell.EntireRow.Cells(1, 6).Value
    If Len(caminho) = 0 Then caminho = ActiveCell.EntireRow.Cells(1, 5).Value
    If Len(caminho) > 0 Then AbrirCaminho caminho
    Exit Sub
Trata:
    TratarErro "modPastas", "AbrirPastaDeOrigemSelecionada", Err.Number, Err.Description
End Sub

Public Sub AbrirCaminho(ByVal caminho As String)
    If Len(Dir(caminho, vbDirectory)) = 0 Then
        MsgBox "Caminho não encontrado: " & caminho, vbExclamation
        Exit Sub
    End If
    ThisWorkbook.FollowHyperlink caminho
End Sub
