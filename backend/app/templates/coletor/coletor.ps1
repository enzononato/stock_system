# Coletor de dados da máquina - Patrimônio. Somente leitura; nada é gravado nesta máquina.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

function Ler { param($bloco) try { & $bloco } catch { $null } }

function Texto {
    param($valor)
    if ($null -eq $valor) { return $null }
    $t = ([string]$valor).Trim()
    if ($t.Length -eq 0) { return $null }
    return $t
}

function IdAnyDesk {
    $conf = Ler { Get-Content -LiteralPath 'C:\ProgramData\AnyDesk\system.conf' -ErrorAction Stop }
    foreach ($linha in @($conf)) {
        if ($linha -match '^\s*ad\.anynet\.id\s*=\s*(\d+)') { return $Matches[1] }
    }
    $executaveis = @()
    if (${env:ProgramFiles(x86)}) { $executaveis += (Join-Path ${env:ProgramFiles(x86)} 'AnyDesk\AnyDesk.exe') }
    if ($env:ProgramFiles) { $executaveis += (Join-Path $env:ProgramFiles 'AnyDesk\AnyDesk.exe') }
    foreach ($exe in $executaveis) {
        if (Test-Path -LiteralPath $exe) {
            $saida = Ler { & $exe --get-id 2>$null }
            $achado = Texto ((@($saida) | Select-Object -First 1))
            if ($achado -and $achado -match '^\d+$') { return $achado }
        }
    }
    return $null
}

function TipoAdaptador {
    param($adaptador)
    $midia = [string]$adaptador.PhysicalMediaType
    $ndis = $adaptador.NdisPhysicalMedium
    if ($midia -match '802\.11' -or $ndis -eq 9) { return 'wifi' }
    if ($midia -match '802\.3') { return 'ethernet' }
    return 'outro'
}

Write-Host 'Coletando os dados desta máquina para o Patrimônio…'

try {
    $cs   = Ler { Get-CimInstance -ClassName Win32_ComputerSystem }
    $csp  = Ler { Get-CimInstance -ClassName Win32_ComputerSystemProduct }
    $bios = Ler { Get-CimInstance -ClassName Win32_BIOS }
    $caixa = Ler { Get-CimInstance -ClassName Win32_SystemEnclosure }
    $cpus = Ler { @(Get-CimInstance -ClassName Win32_Processor) }
    $pentes = Ler { @(Get-CimInstance -ClassName Win32_PhysicalMemory) }
    $so   = Ler { Get-CimInstance -ClassName Win32_OperatingSystem }

    $serial = Texto $bios.SerialNumber
    if (-not $serial) { $serial = Texto $csp.IdentifyingNumber }

    $chassi = @()
    foreach ($c in @($caixa.ChassisTypes)) {
        if ($null -ne $c) { $chassi += [int]$c }
    }

    $ramBytes = $null
    if ($pentes) {
        $soma = ($pentes | Measure-Object -Property Capacity -Sum).Sum
        if ($soma) { $ramBytes = [int64]$soma }
    }
    if (-not $ramBytes -and $cs.TotalPhysicalMemory) { $ramBytes = [int64]$cs.TotalPhysicalMemory }

    $discos = @()
    $lidos = Ler { @(Get-PhysicalDisk) }
    foreach ($d in @($lidos)) {
        if ($null -eq $d) { continue }
        $tam = $null
        if ($null -ne $d.Size) { $tam = [int64]$d.Size }
        $discos += [ordered]@{
            tipo       = Texto $d.MediaType
            bytes      = $tam
            barramento = Texto $d.BusType
        }
    }

    $comGateway = @{}
    $configs = Ler { @(Get-NetIPConfiguration) }
    foreach ($cfg in @($configs)) {
        if ($null -ne $cfg -and $cfg.IPv4DefaultGateway) { $comGateway[[int]$cfg.InterfaceIndex] = $true }
    }

    $adaptadores = @()
    $fisicos = Ler { @(Get-NetAdapter -Physical) }
    foreach ($a in @($fisicos)) {
        if ($null -eq $a) { continue }
        $adaptadores += [ordered]@{
            nome      = Texto $a.Name
            descricao = Texto $a.InterfaceDescription
            mac       = Texto $a.MacAddress
            tipo      = TipoAdaptador $a
            ativo     = ([string]$a.Status -eq 'Up')
            gateway   = [bool]$comGateway[[int]$a.ifIndex]
        }
    }

    $primeiraCpu = @($cpus) | Select-Object -First 1
    $dados = [ordered]@{
        versao_coletor = '__VERSAO__'
        fabricante     = Texto $cs.Manufacturer
        modelo         = Texto $cs.Model
        versao_produto = Texto $csp.Version
        serial         = $serial
        host           = Texto $(if ($cs.Name) { $cs.Name } else { $env:COMPUTERNAME })
        dominio        = Texto $cs.Domain
        no_dominio     = $(if ($null -ne $cs.PartOfDomain) { [bool]$cs.PartOfDomain } else { $null })
        cpu            = Texto $primeiraCpu.Name
        sistema        = Texto $so.Caption
        product_id     = Texto $so.SerialNumber
        anydesk        = Texto (Ler { IdAnyDesk })
        chassi         = @($chassi)
        ram_bytes      = $ramBytes
        discos         = @($discos)
        adaptadores    = @($adaptadores)
    }

    $json = $dados | ConvertTo-Json -Depth 4

    if ($env:COLETOR_SO_MOSTRAR -eq '1') {
        Write-Output $json
        return
    }

    Invoke-RestMethod -Method Post -Uri '__URL_ENVIO__' -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($json)) | Out-Null
    Write-Host 'Pronto! Volte ao navegador.'
    Start-Sleep 5
}
catch {
    $mensagem = $null
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $mensagem = $_.ErrorDetails.Message }
    if (-not $mensagem) { $mensagem = $_.Exception.Message }
    Write-Host ('Não foi possível enviar os dados: ' + $mensagem)
    Read-Host 'Aperte Enter para fechar'
}
