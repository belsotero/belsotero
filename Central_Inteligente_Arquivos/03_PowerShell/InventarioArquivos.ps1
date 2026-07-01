[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$PastaRaiz,
    [Parameter(Mandatory=$true)][string]$PastaSaida,
    [bool]$AnalisarSubpastas = $true,
    [bool]$IgnorarOcultos = $true,
    [bool]$IgnorarTemporarios = $true,
    [int]$TamanhoMaximoExtracaoMB = 25,
    [string]$Tipos = '.txt,.csv,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.pdf,.msg'
)
$ErrorActionPreference = 'Continue'
function New-SafeDirectory { param([string]$Path) if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null } }
function Write-Log { param([string]$Message,[string]$Level='INFO') $line = "$(Get-Date -Format s),InventarioArquivos,$Level,$Message"; Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 }
$csvDir = Join-Path $PastaSaida 'CSV'; $logDir = Join-Path $PastaSaida 'Logs'
New-SafeDirectory $csvDir; New-SafeDirectory $logDir
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:LogPath = Join-Path $logDir "inventario_$stamp.log"
$outCsv = Join-Path $csvDir "inventario_$stamp.csv"
Write-Log "Início do inventário em $PastaRaiz"
$allowed = $Tipos.Split(',') | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ }
$enumParams = @{ LiteralPath = $PastaRaiz; File = $true; Force = $true; ErrorAction = 'SilentlyContinue' }
if ($AnalisarSubpastas) { $enumParams.Recurse = $true }
$id = 0
$result = foreach ($file in Get-ChildItem @enumParams) {
    $id++
    $statusLeitura = 'OK'; $obs = ''
    try {
        if ($IgnorarOcultos -and (($file.Attributes -band [IO.FileAttributes]::Hidden) -ne 0)) { continue }
        if ($IgnorarTemporarios -and ($file.Name -like '~$*' -or $file.Extension -in '.tmp','.temp')) { continue }
        if ($allowed.Count -gt 0 -and $file.Extension.ToLowerInvariant() -notin $allowed) { continue }
        $hash = ''
        try { $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256 -ErrorAction Stop).Hash } catch { $statusLeitura = 'ERRO_HASH'; $obs = $_.Exception.Message }
        [pscustomobject]@{
            ID = $id; Nome = $file.Name; Extensao = $file.Extension; Tipo = if ($file.Extension) { $file.Extension.TrimStart('.').ToUpperInvariant() } else { 'SEM_EXTENSAO' }
            CaminhoCompleto = $file.FullName; PastaOrigem = $file.DirectoryName
            TamanhoKB = [math]::Round($file.Length/1KB,2); TamanhoMB = [math]::Round($file.Length/1MB,2)
            DataCriacao = $file.CreationTime; DataModificacao = $file.LastWriteTime
            StatusLeitura = $statusLeitura; StatusExtracao = 'NAO_EXTRAIDO'; HashSHA256 = $hash
            DuplicadoNome = 'Nao'; DuplicadoTamanho = 'Nao'; DuplicadoHash = 'Nao'; Alerta = 'Nao'; Observacoes = $obs
            AbrirArquivo = $file.FullName; AbrirPasta = $file.DirectoryName
        }
    } catch {
        Write-Log "Erro ao processar $($file.FullName): $($_.Exception.Message)" 'ERRO'
    }
}
$result | Export-Csv -LiteralPath $outCsv -NoTypeInformation -Encoding UTF8
Write-Log "Inventário concluído: $outCsv"
