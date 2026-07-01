[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PastaRaiz,[Parameter(Mandatory=$true)][string]$PastaSaida,[bool]$AnalisarSubpastas=$true,[bool]$IgnorarOcultos=$true,[bool]$IgnorarTemporarios=$true,[int]$TamanhoMaximoExtracaoMB=25,[string]$Tipos='')
function New-SafeDirectory { param([string]$Path) if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null } }
$csvDir=Join-Path $PastaSaida 'CSV'; $htmlDir=Join-Path $PastaSaida 'HTML'; $logDir=Join-Path $PastaSaida 'Logs'; New-SafeDirectory $htmlDir; New-SafeDirectory $logDir
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'; $html=Join-Path $htmlDir "relatorio_$stamp.html"; $log=Join-Path $logDir "relatorio_$stamp.log"
function Latest($filter){ Get-ChildItem -LiteralPath $csvDir -Filter $filter -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
$inv=Latest 'inventario_*.csv'; $dup=Latest 'duplicados_*.csv'; $alt=Latest 'alertas_*.csv'
$body = @("<h1>Central Inteligente de Arquivos</h1>","<p>Gerado em $(Get-Date)</p>","<p>Pasta analisada: $PastaRaiz</p>")
foreach($f in @($inv,$dup,$alt)){ if($f){ $data=Import-Csv $f.FullName; $body += "<h2>$($f.Name)</h2>"; $body += ($data | Select-Object -First 200 | ConvertTo-Html -Fragment) } }
ConvertTo-Html -Title 'Central Inteligente de Arquivos' -Body ($body -join "`n") | Set-Content -LiteralPath $html -Encoding UTF8
Add-Content -LiteralPath $log -Encoding UTF8 -Value "$(Get-Date -Format s),GerarRelatorios,OK,$html"
