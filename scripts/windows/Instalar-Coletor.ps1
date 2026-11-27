param(
    [Parameter(Mandatory = $true)]
    [uri]$Site
)

$ErrorActionPreference = 'Stop'
if (!$Site.IsAbsoluteUri -or $Site.Scheme -ne 'https' -or $Site.AbsolutePath -ne '/' -or
    $Site.Query -or $Site.Fragment -or $Site.UserInfo) {
    throw 'Informe somente a origem HTTPS do ambiente de teste.'
}

$template = Join-Path $PSScriptRoot '..\..\backend\app\templates\coletor\coletor.ps1'
$destino = Join-Path $env:LOCALAPPDATA 'StockSystem\Coletor'
$arquivo = Join-Path $destino 'Coletor.ps1'
$origem = $Site.GetLeftPart([UriPartial]::Authority).Replace("'", "''")
$cabecalho = @'
param([string]$Link)
if ($args.Count -or $Link -cnotmatch '\Astocksystem://coletar/([A-Za-z0-9_-]{22})/?\z') { exit 2 }
$codigo = $Matches[1]
$urlEnvio = '__ORIGEM__/api/coletor/' + $codigo + '/dados'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

'@
$script = $cabecalho.Replace('__ORIGEM__', $origem) +
    (Get-Content -LiteralPath $template -Raw -Encoding UTF8).Replace("'__URL_ENVIO__'", '$urlEnvio').Replace('__VERSAO__', '1')
New-Item -ItemType Directory -Path $destino -Force | Out-Null
Set-Content -LiteralPath $arquivo -Value $script -Encoding UTF8

$registro = 'HKCU:\Software\Classes\stocksystem'
$executavel = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
New-Item -Path $registro -Force | Out-Null
Set-Item -LiteralPath $registro -Value 'URL:Stock System'
New-ItemProperty -LiteralPath $registro -Name 'URL Protocol' -Value '' -Force | Out-Null
$comando = Join-Path $registro 'shell\open\command'
New-Item -Path $comando -Force | Out-Null
Set-Item -LiteralPath $comando -Value ('"{0}" -NoProfile -WindowStyle Hidden -File "{1}" "%1"' -f $executavel, $arquivo)
Write-Host 'Coletor instalado para este usuario do Windows.'
