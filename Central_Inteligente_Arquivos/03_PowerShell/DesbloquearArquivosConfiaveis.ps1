[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PastaRaiz,[Parameter(Mandatory=$true)][string]$PastaSaida,[bool]$AnalisarSubpastas=$true,[bool]$IgnorarOcultos=$true,[bool]$IgnorarTemporarios=$true,[int]$TamanhoMaximoExtracaoMB=25,[string]$Tipos='', [switch]$Confirmar)
function New-SafeDirectory { param([string]$Path) if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null } }
$csvDir=Join-Path $PastaSaida 'CSV'; $logDir=Join-Path $PastaSaida 'Logs'; New-SafeDirectory $csvDir; New-SafeDirectory $logDir
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'; $csv=Join-Path $csvDir "arquivos_bloqueados_$stamp.csv"; $log=Join-Path $logDir "desbloqueio_$stamp.log"
function Write-Log { param($m,$l='INFO') Add-Content -LiteralPath $log -Encoding UTF8 -Value "$(Get-Date -Format s),DesbloquearArquivosConfiaveis,$l,$m" }
$enum=@{LiteralPath=$PastaRaiz;File=$true;Force=$true;ErrorAction='SilentlyContinue'}; if($AnalisarSubpastas){$enum.Recurse=$true}
$blocked = foreach($f in Get-ChildItem @enum){ $ads = Get-Item -LiteralPath $f.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue; if($ads){ [pscustomobject]@{Nome=$f.Name;Caminho=$f.FullName;Pasta=$f.DirectoryName;TamanhoKB=[math]::Round($f.Length/1KB,2);DataModificacao=$f.LastWriteTime;Stream='Zone.Identifier'} } }
$blocked | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
Write-Log "Lista de arquivos bloqueados gerada: $csv"
if(-not $Confirmar){ Write-Log 'Execução apenas listou arquivos bloqueados. Use -Confirmar para confirmar.'; exit 0 }
if(-not $blocked -or $blocked.Count -eq 0){ Write-Log 'Nenhum arquivo bloqueado encontrado.'; exit 0 }
Write-Host "Arquivos bloqueados listados em: $csv"
$answer = Read-Host "Digite DESBLOQUEAR para aplicar Unblock-File somente nesses arquivos confiáveis"
if($answer -ne 'DESBLOQUEAR'){ Write-Log 'Usuária não confirmou desbloqueio.' 'AVISO'; exit 0 }
foreach($b in $blocked){ try { Unblock-File -LiteralPath $b.Caminho -ErrorAction Stop; Write-Log "Desbloqueado: $($b.Caminho)" } catch { Write-Log "Falha ao desbloquear $($b.Caminho): $($_.Exception.Message)" 'ERRO' } }
