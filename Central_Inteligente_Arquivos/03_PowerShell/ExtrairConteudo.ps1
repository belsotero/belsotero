[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PastaRaiz,[Parameter(Mandatory=$true)][string]$PastaSaida,[bool]$AnalisarSubpastas=$true,[bool]$IgnorarOcultos=$true,[bool]$IgnorarTemporarios=$true,[int]$TamanhoMaximoExtracaoMB=25,[string]$Tipos='')
function New-SafeDirectory { param([string]$Path) if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null } }
function Cut-Text { param([string]$Text,[int]$Max=4000) if ($null -eq $Text) { '' } elseif ($Text.Length -gt $Max) { $Text.Substring(0,$Max) } else { $Text } }
$csvDir=Join-Path $PastaSaida 'CSV'; $logDir=Join-Path $PastaSaida 'Logs'; New-SafeDirectory $csvDir; New-SafeDirectory $logDir
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'; $log=Join-Path $logDir "extracao_$stamp.log"
function Write-Log { param($m,$l='INFO') Add-Content -LiteralPath $log -Encoding UTF8 -Value "$(Get-Date -Format s),ExtrairConteudo,$l,$m" }
$latestInv=Get-ChildItem -LiteralPath $csvDir -Filter 'inventario_*.csv' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $latestInv) { Write-Log 'Inventário não encontrado.' 'ERRO'; exit 2 }
$items=Import-Csv -LiteralPath $latestInv.FullName
$rows=@()
$word=$null; $excel=$null; $ppt=$null
try {
foreach ($i in $items) {
    $status='OK'; $obs=''; $texto=''; $secao='Inicial'
    try {
        if ([double]$i.TamanhoMB -gt $TamanhoMaximoExtracaoMB) { throw "Arquivo acima do limite de $TamanhoMaximoExtracaoMB MB" }
        switch ($i.Extensao.ToLowerInvariant()) {
            '.txt' { $texto = Get-Content -LiteralPath $i.CaminhoCompleto -Raw -ErrorAction Stop }
            '.csv' { $texto = Get-Content -LiteralPath $i.CaminhoCompleto -Raw -ErrorAction Stop }
            '.doc' { if (-not $word) { $word=New-Object -ComObject Word.Application; $word.Visible=$false }; $d=$word.Documents.Open($i.CaminhoCompleto,$false,$true); $texto=$d.Content.Text; $d.Close($false) }
            '.docx' { if (-not $word) { $word=New-Object -ComObject Word.Application; $word.Visible=$false }; $d=$word.Documents.Open($i.CaminhoCompleto,$false,$true); $texto=$d.Content.Text; $d.Close($false) }
            '.xls' { if (-not $excel) { $excel=New-Object -ComObject Excel.Application; $excel.Visible=$false }; $wb=$excel.Workbooks.Open($i.CaminhoCompleto,0,$true); $parts=@(); foreach($ws in $wb.Worksheets){$secao=$ws.Name; $parts += ($ws.UsedRange.Text)}; $texto=$parts -join "`n"; $wb.Close($false) }
            '.xlsx' { if (-not $excel) { $excel=New-Object -ComObject Excel.Application; $excel.Visible=$false }; $wb=$excel.Workbooks.Open($i.CaminhoCompleto,0,$true); $parts=@(); foreach($ws in $wb.Worksheets){$secao=$ws.Name; $parts += ($ws.UsedRange.Text)}; $texto=$parts -join "`n"; $wb.Close($false) }
            '.ppt' { if (-not $ppt) { $ppt=New-Object -ComObject PowerPoint.Application }; $pr=$ppt.Presentations.Open($i.CaminhoCompleto,$true,$false,$false); $parts=@(); foreach($s in $pr.Slides){ foreach($shape in $s.Shapes){ if($shape.HasTextFrame){ if($shape.TextFrame.HasText){ $parts += $shape.TextFrame.TextRange.Text } } } }; $texto=$parts -join "`n"; $pr.Close() }
            '.pptx' { if (-not $ppt) { $ppt=New-Object -ComObject PowerPoint.Application }; $pr=$ppt.Presentations.Open($i.CaminhoCompleto,$true,$false,$false); $parts=@(); foreach($s in $pr.Slides){ foreach($shape in $s.Shapes){ if($shape.HasTextFrame){ if($shape.TextFrame.HasText){ $parts += $shape.TextFrame.TextRange.Text } } } }; $texto=$parts -join "`n"; $pr.Close() }
            '.pdf' { $status='PENDENTE_OCR'; $obs='PDF pesquisável/OCR preparado para melhoria futura nesta versão.' }
            '.msg' { $status='NAO_SUPORTADO'; $obs='Extração de e-mail preparada para evolução futura.' }
            default { $status='INCOMPATIVEL'; $obs='Tipo não compatível para extração segura nesta versão.' }
        }
    } catch { $status='ERRO'; $obs=$_.Exception.Message; Write-Log "Erro em $($i.CaminhoCompleto): $obs" 'ERRO' }
    $rows += [pscustomobject]@{IDArquivo=$i.ID;NomeArquivo=$i.Nome;Caminho=$i.CaminhoCompleto;Tipo=$i.Tipo;PaginaAbaSlideSecao=$secao;TrechoExtraido=Cut-Text $texto;PalavraChave='';DataExtracao=Get-Date;Status=$status;ObservacoesErro=$obs}
}
} finally { if($word){$word.Quit()}; if($excel){$excel.Quit()}; if($ppt){$ppt.Quit()} }
$rows | Export-Csv -LiteralPath (Join-Path $csvDir "conteudo_extraido_$stamp.csv") -NoTypeInformation -Encoding UTF8
Write-Log 'Extração concluída.'
