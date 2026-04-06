# =========================================================
# INSTALADOR MULTI-SOFTWARE + TOKENS + CERTIFICADOS + ATALHOS
# =========================================================
# Softwares | Tokens | Certificados Digitais | Atalhos Web | Config VPN
# Versao: 3.4 - Final com Timeout e Config VPN
# =========================================================

#Requires -RunAsAdministrator

$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"

$script:InstallResults = @{}
$script:OriginalExecutionPolicy = $null
$script:RebootRequired = $false

# Timeout padrao para instalacoes (em segundos)
$script:DefaultTimeout = 600  # 10 minutos

# =========================================================
# DETECTAR CAMINHO DO SCRIPT/EXECUTAVEL
# =========================================================

function Get-ScriptDirectory {
    if ($PSScriptRoot) {
        return $PSScriptRoot
    }
    elseif ($MyInvocation.MyCommand.Path) {
        return Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    elseif ($script:MyInvocation.MyCommand.Path) {
        return Split-Path -Parent $script:MyInvocation.MyCommand.Path
    }
    else {
        return Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    }
}

$script:BaseDirectory = Get-ScriptDirectory
Write-Host "[DEBUG] Diretorio base detectado: $script:BaseDirectory" -ForegroundColor Gray

# =========================================================
# BANNER INICIAL
# =========================================================

function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "                                                                " -ForegroundColor Cyan
    Write-Host "     INSTALADOR COMPLETO - SOFTWARES + TOKENS + CERTIFICADOS   " -ForegroundColor Cyan
    Write-Host "                         Versao 3.4                             " -ForegroundColor Cyan
    Write-Host "                                                                " -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  SOFTWARES:" -ForegroundColor Yellow
    Write-Host "    - Adobe Acrobat Reader DC 64-bit"
    Write-Host "    - Oracle Java Runtime Environment"
    Write-Host "    - Google Chrome"
    Write-Host "    - FusionSigner"
    Write-Host "    - Assinador Livre com MobileID"
    Write-Host "    - PJE Office Pro"
    Write-Host "    - FortiClient VPN (com configuracao automatica)"
    Write-Host ""
    Write-Host "  TOKENS E MIDDLEWARE:" -ForegroundColor Yellow
    Write-Host "    - GD Starsign Token"
    Write-Host "    - DXSafe Middleware"
    Write-Host "    - SafeNet Authentication Client"
    Write-Host "    - SafeSign IC Token Admin"
    Write-Host ""
    Write-Host "  CERTIFICADOS DIGITAIS:" -ForegroundColor Yellow
    Write-Host "    - AC SAFEWEB RFB v5"
    Write-Host "    - AC Secretaria da Receita Federal v4"
    Write-Host "    - AC Soluti Multipla v5"
    Write-Host "    - AC Soluti v5"
    Write-Host "    - Autoridade Certificadora Raiz Brasileira v5"
    Write-Host "    - ICP-Brasil v5"
    Write-Host ""
    Write-Host "  ATALHOS NA AREA DE TRABALHO:" -ForegroundColor Yellow
    Write-Host "    - DAM - Diretoria de Administracao de Materiais"
    Write-Host "    - PA Virtual - PGM Rio"
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
}

# =========================================================
# GERENCIAMENTO DE EXECUTIONPOLICY
# =========================================================

function Set-TemporaryExecutionPolicy {
    Write-Host "[POLITICA] Verificando ExecutionPolicy..." -ForegroundColor Cyan
    
    $script:OriginalExecutionPolicy = Get-ExecutionPolicy -Scope CurrentUser
    Write-Host "[POLITICA] Politica atual: $script:OriginalExecutionPolicy" -ForegroundColor Gray
    
    if ($script:OriginalExecutionPolicy -ne "RemoteSigned" -and 
        $script:OriginalExecutionPolicy -ne "Unrestricted") {
        
        Write-Host "[POLITICA] Alterando temporariamente para RemoteSigned..." -ForegroundColor Yellow
        try {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "[POLITICA] OK - Politica alterada com sucesso" -ForegroundColor Green
        }
        catch {
            Write-Warning "[POLITICA] Nao foi possivel alterar ExecutionPolicy"
        }
    } else {
        Write-Host "[POLITICA] OK - ExecutionPolicy ja permite scripts" -ForegroundColor Green
    }
}

function Restore-ExecutionPolicy {
    if ($script:OriginalExecutionPolicy -and 
        $script:OriginalExecutionPolicy -ne "RemoteSigned" -and 
        $script:OriginalExecutionPolicy -ne "Unrestricted") {
        
        Write-Host "`n[POLITICA] Restaurando ExecutionPolicy original..." -ForegroundColor Cyan
        try {
            Set-ExecutionPolicy -ExecutionPolicy $script:OriginalExecutionPolicy -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "[POLITICA] OK - Politica restaurada para: $script:OriginalExecutionPolicy" -ForegroundColor Green
        }
        catch {
            Write-Warning "[POLITICA] Nao foi possivel restaurar ExecutionPolicy"
        }
    }
}

# =========================================================
# VERIFICACAO E INSTALACAO DO WINGET
# =========================================================

function Test-WingetInstalled {
    try {
        $version = winget --version 2>$null
        if ($version -and $LASTEXITCODE -eq 0) {
            return $true
        }
    }
    catch {
        return $false
    }
    return $false
}

function Install-Winget {
    Write-Host "`n[WINGET] Winget nao encontrado. Tentando instalar..." -ForegroundColor Yellow
    
    try {
        $url = "https://aka.ms/getwinget"
        $output = "$env:TEMP\Microsoft.DesktopAppInstaller.msixbundle"
        
        Write-Host "[WINGET] Baixando de: $url" -ForegroundColor Gray
        Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing -ErrorAction Stop
        
        Write-Host "[WINGET] Instalando pacote..." -ForegroundColor Gray
        Add-AppxPackage -Path $output -ErrorAction Stop
        
        Start-Sleep -Seconds 5
        
        if (Test-WingetInstalled) {
            Write-Host "[WINGET] OK - Instalado com sucesso!" -ForegroundColor Green
            return $true
        } else {
            throw "Winget instalado mas nao esta respondendo"
        }
    }
    catch {
        Write-Host "`n[WINGET] FALHA - na instalacao automatica" -ForegroundColor Red
        Write-Host "`nINSTALACAO MANUAL NECESSARIA:" -ForegroundColor Yellow
        Write-Host "1. Abra a Microsoft Store"
        Write-Host "2. Procure por 'App Installer' ou 'Instalador de Aplicativo'"
        Write-Host "3. Clique em 'Atualizar' ou 'Obter'"
        Write-Host "4. Execute este script novamente`n"
        return $false
    }
}

function Ensure-Winget {
    Write-Host "`n[WINGET] Verificando instalacao do Winget..." -ForegroundColor Cyan
    
    if (Test-WingetInstalled) {
        $version = winget --version
        Write-Host "[WINGET] OK - Winget instalado: $version" -ForegroundColor Green
        return $true
    }
    
    return Install-Winget
}

# =========================================================
# INSTALACAO VIA WINGET
# =========================================================

function Install-WingetPackage {
    param (
        [string]$Name,
        [string]$Id
    )
    
    Write-Host "`n----------------------------------------------------------" -ForegroundColor Gray
    Write-Host "  Instalando: $Name" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray
    
    try {
        Write-Host "[INFO] Verificando se ja esta instalado..." -ForegroundColor Gray
        $checkInstalled = winget list --id $Id --exact 2>&1
        
        if ($LASTEXITCODE -eq 0 -and $checkInstalled -notlike "*Nenhum pacote instalado*") {
            Write-Host "[INFO] $Name ja esta instalado" -ForegroundColor Yellow
            $script:InstallResults[$Name] = "JaInstalado"
            return
        }
        
        Write-Host "[INSTALANDO] Executando instalacao via winget..." -ForegroundColor Cyan
        
        $installOutput = winget install `
            --id $Id `
            --exact `
            --scope machine `
            --silent `
            --accept-package-agreements `
            --accept-source-agreements 2>&1
        
        Start-Sleep -Seconds 3
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[SUCESSO] OK - $Name instalado com sucesso" -ForegroundColor Green
            $script:InstallResults[$Name] = "Sucesso"
        } elseif ($LASTEXITCODE -eq -1978335189) {
            Write-Host "[INFO] $Name ja estava instalado" -ForegroundColor Yellow
            $script:InstallResults[$Name] = "JaInstalado"
        } else {
            Write-Warning "[FALHA] Instalacao retornou codigo: $LASTEXITCODE"
            $script:InstallResults[$Name] = "Falhou_$LASTEXITCODE"
        }
    }
    catch {
        Write-Error "[ERRO] Excecao ao instalar: $($_.Exception.Message)"
        $script:InstallResults[$Name] = "Erro_$($_.Exception.Message)"
    }
}

# =========================================================
# INSTALACAO VIA EXECUTAVEL COM TIMEOUT
# =========================================================

function Install-Executable {
    param (
        [string]$Name,
        [string]$FileName,
        [string]$Arguments,
        [int]$TimeoutSeconds = $script:DefaultTimeout
    )
    
    Write-Host "`n----------------------------------------------------------" -ForegroundColor Gray
    Write-Host "  Instalando: $Name" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray
    
    $filePath = Join-Path $script:BaseDirectory $FileName
    
    Write-Host "[DEBUG] Procurando arquivo em: $filePath" -ForegroundColor Gray
    
    if (-not (Test-Path $filePath)) {
        Write-Warning "[AVISO] Arquivo nao encontrado: $FileName"
        Write-Host "[INFO] Caminho esperado: $filePath" -ForegroundColor Yellow
        Write-Host "[INFO] Certifique-se de que o arquivo esta na mesma pasta do instalador" -ForegroundColor Yellow
        $script:InstallResults[$Name] = "ArquivoNaoEncontrado"
        return
    }
    
    try {
        Write-Host "[INFO] Arquivo encontrado: $FileName" -ForegroundColor Gray
        Write-Host "[INFO] Caminho: $filePath" -ForegroundColor Gray
        Write-Host "[INFO] Argumentos: $Arguments" -ForegroundColor Gray
        Write-Host "[INFO] Timeout: $TimeoutSeconds segundos" -ForegroundColor Gray
        Write-Host "[INSTALANDO] Executando instalador..." -ForegroundColor Cyan
        
        $process = Start-Process $filePath `
            -ArgumentList $Arguments `
            -PassThru `
            -NoNewWindow
        
        $finished = $process.WaitForExit($TimeoutSeconds * 1000)
        
        if (-not $finished) {
            Write-Warning "[TIMEOUT] Instalacao excedeu $TimeoutSeconds segundos"
            Write-Host "[INFO] Forcando encerramento do processo..." -ForegroundColor Yellow
            
            try {
                $process.Kill()
                $process.WaitForExit(5000)
            }
            catch {
                Write-Warning "[AVISO] Nao foi possivel encerrar o processo"
            }
            
            $script:InstallResults[$Name] = "Timeout_$TimeoutSeconds`s"
            return
        }
        
        $exitCode = $process.ExitCode
        
        if ($exitCode -eq 0) {
            Write-Host "[SUCESSO] OK - $Name instalado com sucesso" -ForegroundColor Green
            $script:InstallResults[$Name] = "Sucesso"
        } elseif ($exitCode -eq 3010 -or $exitCode -eq 999) {
            Write-Host "[SUCESSO] OK - $Name instalado (reinicializacao necessaria)" -ForegroundColor Green
            $script:InstallResults[$Name] = "Sucesso_RebootRequerido"
            $script:RebootRequired = $true
        } else {
            Write-Warning "[AVISO] Instalador retornou codigo: $exitCode"
            $script:InstallResults[$Name] = "Concluido_Codigo$exitCode"
        }
    }
    catch {
        Write-Error "[ERRO] Falha na instalacao: $($_.Exception.Message)"
        $script:InstallResults[$Name] = "Erro_$($_.Exception.Message)"
    }
}

# =========================================================
# INSTALACAO VIA MSI COM TIMEOUT
# =========================================================

function Install-MSI {
    param (
        [string]$Name,
        [string]$MsiFile,
        [string]$TransformFile = $null,
        [int]$TimeoutSeconds = $script:DefaultTimeout
    )
    
    Write-Host "`n----------------------------------------------------------" -ForegroundColor Gray
    Write-Host "  Instalando: $Name" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray
    
    $msiPath = Join-Path $script:BaseDirectory $MsiFile
    
    Write-Host "[DEBUG] Procurando MSI em: $msiPath" -ForegroundColor Gray
    
    if (-not (Test-Path $msiPath)) {
        Write-Warning "[AVISO] Arquivo MSI nao encontrado: $MsiFile"
        Write-Host "[INFO] Caminho esperado: $msiPath" -ForegroundColor Yellow
        Write-Host "[INFO] Certifique-se de que o arquivo esta na mesma pasta do instalador" -ForegroundColor Yellow
        $script:InstallResults[$Name] = "ArquivoNaoEncontrado"
        return
    }
    
    $arguments = "/i `"$msiPath`" /qn /norestart REBOOT=ReallySuppress ALLUSERS=1"
    
    if ($TransformFile) {
        $transformPath = Join-Path $script:BaseDirectory $TransformFile
        
        if (Test-Path $transformPath) {
            $arguments += " TRANSFORMS=`"$transformPath`""
            Write-Host "[INFO] Transform encontrado: $TransformFile" -ForegroundColor Gray
        } else {
            Write-Warning "[AVISO] Transform nao encontrado: $TransformFile"
            Write-Host "[INFO] Continuando sem Transform..." -ForegroundColor Yellow
        }
    }
    
    try {
        Write-Host "[INFO] MSI: $MsiFile" -ForegroundColor Gray
        Write-Host "[INFO] Timeout: $TimeoutSeconds segundos" -ForegroundColor Gray
        Write-Host "[INSTALANDO] Executando msiexec..." -ForegroundColor Cyan
        
        $process = Start-Process msiexec.exe `
            -ArgumentList $arguments `
            -PassThru `
            -NoNewWindow
        
        $finished = $process.WaitForExit($TimeoutSeconds * 1000)
        
        if (-not $finished) {
            Write-Warning "[TIMEOUT] Instalacao MSI excedeu $TimeoutSeconds segundos"
            Write-Host "[INFO] Forcando encerramento do processo..." -ForegroundColor Yellow
            
            try {
                $process.Kill()
                $process.WaitForExit(5000)
            }
            catch {
                Write-Warning "[AVISO] Nao foi possivel encerrar o processo"
            }
            
            $script:InstallResults[$Name] = "Timeout_$TimeoutSeconds`s"
            return
        }
        
        $exitCode = $process.ExitCode
        
        if ($exitCode -eq 0) {
            Write-Host "[SUCESSO] OK - $Name instalado com sucesso" -ForegroundColor Green
            $script:InstallResults[$Name] = "Sucesso"
        } elseif ($exitCode -eq 3010) {
            Write-Host "[SUCESSO] OK - $Name instalado (reinicializacao necessaria)" -ForegroundColor Green
            $script:InstallResults[$Name] = "Sucesso_RebootRequerido"
            $script:RebootRequired = $true
        } elseif ($exitCode -eq 1638) {
            Write-Host "[INFO] $Name ja esta instalado (versao igual ou superior)" -ForegroundColor Yellow
            $script:InstallResults[$Name] = "JaInstalado"
        } else {
            Write-Warning "[AVISO] MSI retornou codigo: $exitCode"
            $script:InstallResults[$Name] = "Concluido_Codigo$exitCode"
        }
    }
    catch {
        Write-Error "[ERRO] Falha na instalacao: $($_.Exception.Message)"
        $script:InstallResults[$Name] = "Erro_$($_.Exception.Message)"
    }
}

# =========================================================
# CONFIGURACAO DO FORTICLIENT VPN
# =========================================================

function Configure-FortiClientVPN {
    param (
        [string]$Name = "Configuracao FortiClient VPN"
    )
    
    Write-Host "`n----------------------------------------------------------" -ForegroundColor Gray
    Write-Host "  Configurando: $Name" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray
    
    # Caminho do arquivo de configuracao XML
    $cfgFile = Join-Path $script:BaseDirectory "configuracao_VPN.xml"
    
    # Caminho do executavel FCConfig.exe
    $fcConfigExe = "C:\Program Files\Fortinet\FortiClient\FCConfig.exe"
    
    # Verificar se arquivo de configuracao existe
    if (-not (Test-Path $cfgFile)) {
        Write-Warning "[AVISO] Arquivo de configuracao nao encontrado: configuracao_VPN.xml"
        Write-Host "[INFO] Caminho esperado: $cfgFile" -ForegroundColor Yellow
        $script:InstallResults[$Name] = "ArquivoNaoEncontrado"
        return
    }
    
    # Verificar se FortiClient foi instalado
    if (-not (Test-Path $fcConfigExe)) {
        Write-Warning "[AVISO] FortiClient nao encontrado: $fcConfigExe"
        Write-Host "[INFO] Certifique-se de que o FortiClient foi instalado corretamente" -ForegroundColor Yellow
        $script:InstallResults[$Name] = "FortiClientNaoInstalado"
        return
    }
    
    try {
        Write-Host "[INFO] Arquivo de configuracao encontrado: $cfgFile" -ForegroundColor Gray
        Write-Host "[INFO] FortiClient encontrado: $fcConfigExe" -ForegroundColor Gray
        Write-Host "[INFO] Importando configuracao VPN..." -ForegroundColor Cyan
        
        # Comando: FCConfig.exe -m vpn -f "config.xml" -o import -i 1 -p12345678
        $arguments = "-m vpn -f `"$cfgFile`" -o import -i 1 -p12345678"
        
        Write-Host "[DEBUG] Comando:" -ForegroundColor Gray
        Write-Host "[DEBUG] `"$fcConfigExe`" $arguments" -ForegroundColor Gray
        
        # Executar FCConfig
        $process = Start-Process $fcConfigExe `
            -ArgumentList $arguments `
            -Wait `
            -PassThru `
            -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Host "[SUCESSO] OK - Configuracao VPN importada com sucesso" -ForegroundColor Green
            $script:InstallResults[$Name] = "Sucesso"
        } else {
            Write-Warning "[AVISO] FCConfig retornou codigo: $($process.ExitCode)"
            $script:InstallResults[$Name] = "Concluido_Codigo$($process.ExitCode)"
        }
    }
    catch {
        Write-Error "[ERRO] Falha ao configurar VPN: $($_.Exception.Message)"
        $script:InstallResults[$Name] = "Erro_$($_.Exception.Message)"
    }
}

# =========================================================
# INSTALACAO DE CERTIFICADOS DIGITAIS
# =========================================================

function Install-Certificate {
    param (
        [string]$Name,
        [string]$CertFile,
        [string]$StoreLocation = "LocalMachine",
        [string]$StoreName = "Root"
    )
    
    Write-Host "`n----------------------------------------------------------" -ForegroundColor Gray
    Write-Host "  Instalando Certificado: $Name" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray
    
    $certPath = Join-Path $script:BaseDirectory $CertFile
    
    Write-Host "[DEBUG] Procurando certificado em: $certPath" -ForegroundColor Gray
    
    if (-not (Test-Path $certPath)) {
        Write-Warning "[AVISO] Certificado nao encontrado: $CertFile"
        Write-Host "[INFO] Caminho esperado: $certPath" -ForegroundColor Yellow
        $script:InstallResults[$Name] = "ArquivoNaoEncontrado"
        return
    }
    
    try {
        Write-Host "[INFO] Certificado encontrado: $CertFile" -ForegroundColor Gray
        Write-Host "[INFO] Store: $StoreLocation\$StoreName" -ForegroundColor Gray
        Write-Host "[INSTALANDO] Importando certificado..." -ForegroundColor Cyan
        
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($certPath)
        
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
            $StoreName, 
            $StoreLocation
        )
        $store.Open("ReadWrite")
        
        $existingCert = $store.Certificates | Where-Object { 
            $_.Thumbprint -eq $cert.Thumbprint 
        }
        
        if ($existingCert) {
            Write-Host "[INFO] Certificado ja esta instalado" -ForegroundColor Yellow
            $script:InstallResults[$Name] = "JaInstalado"
            $store.Close()
            return
        }
        
        $store.Add($cert)
        $store.Close()
        
        Write-Host "[SUCESSO] OK - Certificado instalado com sucesso" -ForegroundColor Green
        Write-Host "[INFO] Thumbprint: $($cert.Thumbprint)" -ForegroundColor Gray
        $script:InstallResults[$Name] = "Sucesso"
    }
    catch {
        Write-Error "[ERRO] Falha ao instalar certificado: $($_.Exception.Message)"
        $script:InstallResults[$Name] = "Erro_$($_.Exception.Message)"
    }
}

# =========================================================
# CRIACAO DE ATALHOS NA AREA DE TRABALHO
# =========================================================

function Create-WebShortcut {
    param (
        [string]$Name,
        [string]$Url,
        [string]$ShortcutName
    )
    
    Write-Host "`n----------------------------------------------------------" -ForegroundColor Gray
    Write-Host "  Criando Atalho: $Name" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray
    
    try {
        $publicDesktop = [Environment]::GetFolderPath("CommonDesktopDirectory")
        $shortcutPath = Join-Path $publicDesktop "$ShortcutName.url"
        
        Write-Host "[INFO] Criando atalho em: $shortcutPath" -ForegroundColor Gray
        Write-Host "[INFO] URL: $Url" -ForegroundColor Gray
        
        if (Test-Path $shortcutPath) {
            Write-Host "[INFO] Atalho ja existe, substituindo..." -ForegroundColor Yellow
            Remove-Item $shortcutPath -Force
        }
        
        $urlContent = @"
[InternetShortcut]
URL=$Url
IconIndex=0
"@
        
        Set-Content -Path $shortcutPath -Value $urlContent -Encoding ASCII
        
        Write-Host "[SUCESSO] OK - Atalho criado com sucesso" -ForegroundColor Green
        $script:InstallResults[$Name] = "Sucesso"
    }
    catch {
        Write-Error "[ERRO] Falha ao criar atalho: $($_.Exception.Message)"
        $script:InstallResults[$Name] = "Erro_$($_.Exception.Message)"
    }
}

# =========================================================
# RELATORIO FINAL
# =========================================================

function Show-Report {
    Write-Host "`n`n"
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "                                                                " -ForegroundColor Cyan
    Write-Host "                  RELATORIO DE INSTALACAO                       " -ForegroundColor Cyan
    Write-Host "                                                                " -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $totalApps = $script:InstallResults.Count
    $successCount = 0
    $alreadyInstalled = 0
    $failedCount = 0
    $notFoundCount = 0
    $timeoutCount = 0
    
    foreach ($app in $script:InstallResults.Keys | Sort-Object) {
        $result = $script:InstallResults[$app]
        
        $icon = ""
        $color = "White"
        $status = ""
        
        if ($result -eq "Sucesso" -or $result -eq "Sucesso_RebootRequerido") {
            $icon = "[OK]"
            $color = "Green"
            $status = "Instalado com sucesso"
            $successCount++
        } elseif ($result -eq "JaInstalado") {
            $icon = "[>>]"
            $color = "Yellow"
            $status = "Ja estava instalado"
            $alreadyInstalled++
        } elseif ($result -eq "ArquivoNaoEncontrado") {
            $icon = "[XX]"
            $color = "Red"
            $status = "Arquivo nao encontrado"
            $notFoundCount++
        } elseif ($result -like "Timeout_*") {
            $icon = "[!!]"
            $color = "Red"
            $status = "Timeout - Instalacao travou"
            $timeoutCount++
        } elseif ($result -like "Falhou_*" -or $result -like "Erro_*") {
            $icon = "[XX]"
            $color = "Red"
            $status = "Falha na instalacao"
            $failedCount++
        } else {
            $icon = "[??]"
            $color = "Yellow"
            $status = $result
        }
        
        Write-Host "  $icon " -NoNewline -ForegroundColor $color
        Write-Host "$app" -NoNewline
        Write-Host " - " -NoNewline -ForegroundColor Gray
        Write-Host "$status" -ForegroundColor $color
    }
    
    Write-Host "`n----------------------------------------------------------" -ForegroundColor Gray
    Write-Host "  RESUMO:" -ForegroundColor Yellow
    Write-Host "    Total de itens: $totalApps"
    Write-Host "    Instalados com sucesso: " -NoNewline
    Write-Host "$successCount" -ForegroundColor Green
    Write-Host "    Ja instalados: " -NoNewline
    Write-Host "$alreadyInstalled" -ForegroundColor Yellow
    Write-Host "    Falharam: " -NoNewline
    Write-Host "$failedCount" -ForegroundColor Red
    Write-Host "    Timeout (travou): " -NoNewline
    Write-Host "$timeoutCount" -ForegroundColor Red
    Write-Host "    Arquivos nao encontrados: " -NoNewline
    Write-Host "$notFoundCount" -ForegroundColor Red
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray
    
    if ($script:RebootRequired) {
        Write-Host ""
        Write-Host "  [!] ATENCAO: Reinicializacao do sistema necessaria!" -ForegroundColor Yellow
        Write-Host "      Algumas instalacoes requerem reinicializacao para funcionar." -ForegroundColor Yellow
    }
    
    if ($timeoutCount -gt 0) {
        Write-Host ""
        Write-Host "  [!] TIMEOUTS DETECTADOS:" -ForegroundColor Yellow
        Write-Host "      Algumas instalacoes travaram e foram forcadas a encerrar." -ForegroundColor Yellow
        Write-Host "      Verifique manualmente se os softwares foram instalados." -ForegroundColor Yellow
    }
    
    if ($notFoundCount -gt 0) {
        Write-Host ""
        Write-Host "  [i] ARQUIVOS NAO ENCONTRADOS:" -ForegroundColor Yellow
        Write-Host "      Certifique-se de que todos os instaladores estao na" -ForegroundColor Yellow
        Write-Host "      mesma pasta do EXECUTAVEL e tente novamente." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "      Diretorio base: $script:BaseDirectory" -ForegroundColor Gray
    }
    
    Write-Host ""
}

# =========================================================
# FUNCAO PRINCIPAL
# =========================================================

function Main {
    Show-Banner
    Set-TemporaryExecutionPolicy
    
    if (-not (Ensure-Winget)) {
        Write-Host "`n[ERRO] Nao e possivel continuar sem o Winget instalado." -ForegroundColor Red
        Write-Host "Instale manualmente e execute o script novamente.`n"
        Restore-ExecutionPolicy
        Read-Host "Pressione Enter para sair"
        exit 1
    }
    
    # =========================================================
    # SOFTWARES VIA WINGET
    # =========================================================
    
    Write-Host "`n"
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  INSTALACOES VIA WINGET (Internet necessaria)" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    
    Install-WingetPackage -Name "Adobe Acrobat Reader DC 64-bit" -Id "Adobe.Acrobat.Reader.64-bit"
    Install-WingetPackage -Name "Oracle Java Runtime Environment" -Id "Oracle.JavaRuntimeEnvironment"
    Install-WingetPackage -Name "Google Chrome" -Id "Google.Chrome"
    
    # =========================================================
    # SOFTWARES VIA EXECUTAVEL LOCAL
    # =========================================================
    
    Write-Host "`n"
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  INSTALACOES VIA EXECUTAVEL (Arquivos locais)" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    
    Install-Executable `
        -Name "FusionSigner" `
        -FileName "FusionSigner_Instalador.exe" `
        -Arguments "/S /y /install" `
        -TimeoutSeconds 300
    
    Install-Executable `
        -Name "PJE Office Pro" `
        -FileName "PJEOfficePro__Windows__x64__installer.exe" `
        -Arguments "/VERYSILENT /SUPPRESSMSGBOXES /ALLUSERS /NOCANCEL /NORESTART /RESTARTEXITCODE=999 /FORCECLOSEAPPLICATIONS" `
        -TimeoutSeconds 300
    
    # =========================================================
    # SOFTWARES VIA MSI LOCAL
    # =========================================================
    
    Write-Host "`n"
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  INSTALACOES VIA MSI (Arquivos locais)" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    
    Install-MSI `
        -Name "Assinador Livre com MobileID" `
        -MsiFile "AssinadorLivreComMobileID.msi" `
        -TransformFile "AssinadorLivreComMobileID_Transform.mst" `
        -TimeoutSeconds 300
    
    # FortiClient VPN
    Install-MSI `
        -Name "FortiClient VPN" `
        -MsiFile "FortiClientVPN.msi" `
        -TimeoutSeconds 300
    
    # =========================================================
    # CONFIGURACAO DO FORTICLIENT VPN
    # =========================================================
    
    Write-Host "`n"
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  CONFIGURACAO VPN" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    
    Configure-FortiClientVPN
    
    # =========================================================
    # TOKENS E MIDDLEWARE
    # =========================================================
    
    Write-Host "`n"
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  TOKENS E MIDDLEWARE (Arquivos locais)" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    
    Install-Executable `
        -Name "GD Starsign Token Driver" `
        -FileName "GDsetupStarsignCUTx64_Setup_v1.7.17.0.exe" `
        -Arguments "/S /v/qn" `
        -TimeoutSeconds 180
    
    Install-Executable `
        -Name "DXSafe Middleware" `
        -FileName "Instalador DXSafe Middleware -  1.0.34.exe" `
        -Arguments "/S /v/qn" `
        -TimeoutSeconds 120
    
    Install-MSI `
        -Name "SafeNet Authentication Client" `
        -MsiFile "SafeNet-10.6_Win10-64bits.msi" `
        -TimeoutSeconds 300
    
    Install-MSI `
        -Name "SafeSign IC Token Admin" `
        -MsiFile "SafeSignIC__W64__v3.5.3.0-AET.000__TokenAdmin.msi" `
        -TimeoutSeconds 300
    
    # =========================================================
    # CERTIFICADOS DIGITAIS
    # =========================================================
    
    Write-Host "`n"
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  CERTIFICADOS DIGITAIS" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    
    Install-Certificate `
        -Name "ICP-Brasil v5" `
        -CertFile "ICP-Brasilv5.crt" `
        -StoreLocation "LocalMachine" `
        -StoreName "Root"
    
    Install-Certificate `
        -Name "Autoridade Certificadora Raiz Brasileira v5" `
        -CertFile "Autoridade Certificadora Raiz Brasileira v5.cer" `
        -StoreLocation "LocalMachine" `
        -StoreName "Root"
    
    Install-Certificate `
        -Name "AC SAFEWEB RFB v5" `
        -CertFile "AC SAFEWEB RFB v5.cer" `
        -StoreLocation "LocalMachine" `
        -StoreName "CA"
    
    Install-Certificate `
        -Name "AC Secretaria da Receita Federal v4" `
        -CertFile "AC Secretaria da Receita Federal do Brasil v4.cer" `
        -StoreLocation "LocalMachine" `
        -StoreName "CA"
    
    Install-Certificate `
        -Name "AC Soluti Multipla v5" `
        -CertFile "ac-soluti-multipla-v5.crt" `
        -StoreLocation "LocalMachine" `
        -StoreName "CA"
    
    Install-Certificate `
        -Name "AC Soluti v5" `
        -CertFile "ac-soluti-v5.crt" `
        -StoreLocation "LocalMachine" `
        -StoreName "CA"
    
    # =========================================================
    # ATALHOS NA AREA DE TRABALHO
    # =========================================================
    
    Write-Host "`n"
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  ATALHOS NA AREA DE TRABALHO" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    
    Create-WebShortcut `
        -Name "Atalho DAM" `
        -Url "https://dam.rio.rj.gov.br/" `
        -ShortcutName "DAM - Diretoria de Administracao de Materiais"
    
    Create-WebShortcut `
        -Name "Atalho PA Virtual" `
        -Url "http://pavirtual.pgm.rio.rj.gov.br/portal" `
        -ShortcutName "PA Virtual - PGM Rio"
    
    # =========================================================
    # RELATORIO FINAL
    # =========================================================
    
    Show-Report
    Restore-ExecutionPolicy
    
    Write-Host ""
    Write-Host "Pressione ENTER para reiniciar o computador..." -ForegroundColor Yellow
    Read-Host
    Restart-Computer -Force
}

# =========================================================
# EXECUTAR SCRIPT
# =========================================================

try {
    Main
}
catch {
    Write-Host "`n`n[ERRO CRITICO] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Restore-ExecutionPolicy
    Read-Host "`nPressione Enter para sair"
    exit 1
}
