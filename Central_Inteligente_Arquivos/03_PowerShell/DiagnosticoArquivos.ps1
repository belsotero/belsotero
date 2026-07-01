[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PastaRaiz,[Parameter(Mandatory=$true)][string]$PastaSaida,[bool]$AnalisarSubpastas=$true,[bool]$IgnorarOcultos=$true,[bool]$IgnorarTemporarios=$true,[int]$TamanhoMaximoExtracaoMB=25,[string]$Tipos='')
function New-SafeDirectory { param([string]$Path) if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null } }
$csvDir = Join-Path $PastaSaida 'CSV'; $logDir = Join-Path $PastaSaida 'Logs'; New-SafeDirectory $csvDir; New-SafeDirectory $logDir
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'; $log = Join-Path $logDir "diagnostico_$stamp.log"
function Write-Log { param($m,$l='INFO') Add-Content -LiteralPath $log -Encoding UTF8 -Value "$(Get-Date -Format s),DiagnosticoArquivos,$l,$m" }
$latestInv = Get-ChildItem -LiteralPath $csvDir -Filter 'inventario_*.csv' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $latestInv) { Write-Log 'Inventário não encontrado.' 'ERRO'; exit 2 }
$items = Import-Csv -LiteralPath $latestInv.FullName
$duplicados = @()
foreach ($g in $items | Group-Object Nome | Where-Object Count -gt 1) { foreach ($i in $g.Group) { $duplicados += [pscustomobject]@{Criterio='Nome';HashSHA256=$i.HashSHA256;Nome=$i.Nome;Tamanho=$i.TamanhoMB;Caminho=$i.CaminhoCompleto;PastaOrigem=$i.PastaOrigem;DataModificacao=$i.DataModificacao;Observacao='Mesmo nome encontrado'} } }
foreach ($g in $items | Where-Object {[double]$_.TamanhoKB -gt 0} | Group-Object TamanhoKB | Where-Object Count -gt 1) { foreach ($i in $g.Group) { $duplicados += [pscustomobject]@{Criterio='Tamanho';HashSHA256=$i.HashSHA256;Nome=$i.Nome;Tamanho=$i.TamanhoMB;Caminho=$i.CaminhoCompleto;PastaOrigem=$i.PastaOrigem;DataModificacao=$i.DataModificacao;Observacao='Mesmo tamanho encontrado'} } }
foreach ($g in $items | Where-Object {$_.HashSHA256} | Group-Object HashSHA256 | Where-Object Count -gt 1) { foreach ($i in $g.Group) { $duplicados += [pscustomobject]@{Criterio='HashSHA256';HashSHA256=$i.HashSHA256;Nome=$i.Nome;Tamanho=$i.TamanhoMB;Caminho=$i.CaminhoCompleto;PastaOrigem=$i.PastaOrigem;DataModificacao=$i.DataModificacao;Observacao='Conteúdo duplicado por hash'} } }
$alertas = @(); $id=0
foreach ($i in $items) {
    $checks = @()
    if ([double]$i.TamanhoKB -eq 0) { $checks += @('Arquivo 0 KB|Alta|Verificar se o arquivo está vazio ou corrompido') }
    if (-not $i.Extensao) { $checks += @('Sem extensão|Media|Confirmar tipo antes de abrir') }
    if ($i.Nome -like '~$*' -or $i.Extensao -in '.tmp','.temp') { $checks += @('Temporário|Baixa|Normalmente gerado por aplicativo aberto') }
    if ([double]$i.TamanhoMB -gt $TamanhoMaximoExtracaoMB) { $checks += @('Muito grande|Media|Não extrair automaticamente') }
    if ($i.CaminhoCompleto.Length -gt 240) { $checks += @('Caminho longo|Media|Pode falhar em aplicações antigas') }
    if ($i.StatusLeitura -ne 'OK') { $checks += @('Erro de leitura|Alta|Consultar log técnico') }
    foreach ($c in $checks) { $id++; $p=$c.Split('|'); $alertas += [pscustomobject]@{ID=$id;TipoAlerta=$p[0];Severidade=$p[1];Arquivo=$i.Nome;Caminho=$i.CaminhoCompleto;Detalhe=$p[0];AcaoRecomendada=$p[2];DataHora=Get-Date} }
}
$duplicados | Export-Csv -LiteralPath (Join-Path $csvDir "duplicados_$stamp.csv") -NoTypeInformation -Encoding UTF8
$alertas | Export-Csv -LiteralPath (Join-Path $csvDir "alertas_$stamp.csv") -NoTypeInformation -Encoding UTF8
Write-Log 'Diagnóstico concluído.'
