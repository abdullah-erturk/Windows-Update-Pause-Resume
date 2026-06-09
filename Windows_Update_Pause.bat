<# : hybrid batch + powershell script
@echo off
chcp 65001 >nul
title Windows Update Pause / Resume ^| by Abdullah ERTÜRK
:: Eski konsol açıldıktan sonra ekranı temizle ve boyutu ayarla
::mode con: cols=70 lines=4
for %%A in ("%~f0") do set "BAT_PATH=%%~A"
powershell.exe -noprofile -ExecutionPolicy Bypass -Command "& { $scriptPath = '%BAT_PATH%'; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; if (Test-Path $scriptPath) { $content = [System.IO.File]::ReadAllText($scriptPath, $utf8NoBom); Invoke-Expression $content } else { Write-Host 'Dosya bulunamadı: ' $scriptPath } }"
exit /b
#>

# =====================================================================
#  Windows Update Pause / Resume  (PowerShell + hybrid .bat launcher)
#  Copyright (C) 2024-2026 Abdullah ERTÜRK
#  https://github.com/abdullah-erturk/Windows-Update-Pause-Resume
#
#  [TR]
#  - Windows 10 / 11 (build >= 15063): yerleşik "Pause Updates" API'si ile
#    güncelleştirmeler gerçek bitiş tarihine kadar (UTC) duraklatılır.
#  - Windows Server 2016 ve eski sürümler (build < 15063): Pause API'si
#    bulunmadığından otomatik OS güncelleştirmeleri "NoAutoUpdate" ilkesiyle
#    durdurulur. wuauserv hizmeti ÇALIŞMAYA DEVAM EDER; böylece Microsoft
#    Defender imza (tanım) güncellemeleri kesilmez. N hafta sonra ilkeyi
#    geri açan zamanlanmış görev kurulur.
#  - İşletim sistemi dili Türkçe ise arayüz Türkçe, değilse İngilizce.
#
#  [EN]
#  - Windows 10 / 11 (build >= 15063): pauses updates until a real expiry
#    date (UTC) using the built-in "Pause Updates" API.
#  - Windows Server 2016 and older (build < 15063): the Pause API is not
#    available, so automatic OS updates are stopped via the "NoAutoUpdate"
#    policy. The wuauserv service KEEPS RUNNING, so Microsoft Defender
#    signature (definition) updates are not interrupted. A scheduled task
#    is created to turn the policy back off after N weeks.
#  - The UI is Turkish if the OS language is Turkish, otherwise English.
# =====================================================================

# ---- Yönetici olarak çalıştır / Self-elevate ----
$principal = New-Object Security.Principal.WindowsPrincipal(
    [Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    try { Start-Process -FilePath $scriptPath -Verb RunAs } catch {}
    exit
}

Add-Type -AssemblyName Microsoft.VisualBasic

# ---- Dil tespiti / Language detection ----
function Test-Turkish {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        if ($os.MUILanguages -and $os.MUILanguages.Count -ge 1) {
            return $os.MUILanguages[0].ToLower().StartsWith('tr')   # aktif görüntüleme dili
        }
        if ($os.OSLanguage -eq 1055) { return $true }               # 1055 = Türkçe
    } catch {}
    return ((Get-Culture).LCID -eq 1055)
}
$isTR = Test-Turkish

# ---- Metinler / Strings ----
if ($isTR) {
    $L_Title       = 'Windows Updates Pause / Resume | by Abdullah ERTÜRK'
    $L_ResumeMsg   = "Windows güncelleştirmeleri yeniden etkinleştirildi.`r`n`r`nWindows Update hizmeti çalışır duruma getirildi."
    $L_PausePrompt = "Windows güncelleştirmelerini kaç hafta duraklatmak istiyorsunuz?`r`n`r`n(1 - {0} hafta arası)`r`n`r`nÖrnek: 4, 10, 20, 30, 40 ..."
    $L_PromptTitle = 'Güncelleştirmeleri Duraklatma Süresi'
    $L_PausedTmpl  = 'Windows güncelleştirmeleri {0} hafta ({1} gün) süreyle duraklatıldı.'
    $L_Invalid     = 'Lütfen 1 ile {0} arasında geçerli bir tam sayı girin.'
} else {
    $L_Title       = 'Windows Updates Pause / Resume | by Abdullah ERTÜRK'
    $L_ResumeMsg   = "Windows Updates have been re-enabled.`r`n`r`nThe Windows Update service is running again."
    $L_PausePrompt = "Enter the number of weeks you want to pause Windows Updates:`r`n`r`n(between 1 and {0} weeks)`r`n`r`nExample: 4, 10, 20, 30, 40 ..."
    $L_PromptTitle = 'Duration to Pause Updates'
    $L_PausedTmpl  = 'Windows Updates paused for {0} weeks ({1} days).'
    $L_Invalid     = 'Please enter a valid whole number between 1 and {0}.'
}

# ---- Sabitler ve yardımcılar / Constants & helpers ----
$MaxWeeks = 10000         # üst sınır - tarih taşmasını önler / upper limit (prevents date overflow)
$RegKey   = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings'
$AUKey    = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
$TaskName = 'WindowsUpdatePause_AutoResume'
$PauseValues = @(
    'PauseFeatureUpdatesStartTime','PauseFeatureUpdatesEndTime',
    'PauseQualityUpdatesStartTime','PauseQualityUpdatesEndTime',
    'PauseUpdatesStartTime','PauseUpdatesExpiryTime'
)
$InfoModal = [Microsoft.VisualBasic.MsgBoxStyle]::Information  -bor [Microsoft.VisualBasic.MsgBoxStyle]::SystemModal
$WarnModal = [Microsoft.VisualBasic.MsgBoxStyle]::Exclamation -bor [Microsoft.VisualBasic.MsgBoxStyle]::SystemModal

# Yerleşik Pause API'si destekleniyor mu? (Win10 1703 / build 15063 ve sonrası)
$buildNum = 0
[int]::TryParse((Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue).CurrentBuildNumber, [ref]$buildNum) | Out-Null
$supportsPauseApi = $buildNum -ge 15063

function Show-Msg($text, $style) {
    [Microsoft.VisualBasic.Interaction]::MsgBox($text, $style, $L_Title) | Out-Null
}
function Restart-UpdateService {
    try { Stop-Service  -Name wuauserv -Force -ErrorAction SilentlyContinue } catch {}
    try { Start-Service -Name wuauserv       -ErrorAction SilentlyContinue } catch {}
}
function Refresh-Settings {
    Stop-Process -Name SystemSettings -Force -ErrorAction SilentlyContinue
    # ms-settings (UWP) yükseltilmiş süreçten doğrudan açılamaz (0x80040905);
    # explorer.exe üzerinden başlatmak kullanıcı bağlamına devreder.
    if ($supportsPauseApi) {
        try { Start-Process -FilePath 'explorer.exe' -ArgumentList 'ms-settings:windowsupdate' } catch {}
    }
}
# Otomatik OS güncelleştirme ilkesi (hizmeti durdurmaz -> Defender güncel kalır)
function Set-NoAutoUpdate([int]$val) {
    if (-not (Test-Path $AUKey)) { New-Item -Path $AUKey -Force | Out-Null }
    Set-ItemProperty -Path $AUKey -Name 'NoAutoUpdate' -Value $val -Type DWord
}
# Eski sürümlerden kalmış olabilecek devre dışı hizmeti tekrar etkinleştir
function Enable-UpdateService {
    try { Set-Service  -Name wuauserv -StartupType Manual -ErrorAction SilentlyContinue } catch {}
    try { Start-Service -Name wuauserv -ErrorAction SilentlyContinue } catch {}
}
function Remove-ResumeTask {
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch {}
}
function Register-ResumeTask([int]$days) {
    $resumeCmd = "Set-ItemProperty -Path '$AUKey' -Name 'NoAutoUpdate' -Value 0 -Type DWord; Restart-Service -Name wuauserv -Force -ErrorAction SilentlyContinue; Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false"
    $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -Command `"$resumeCmd`""
    $trigger   = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddDays($days))
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
}

# ---- Mevcut durum / Current state ----
$expiry = (Get-ItemProperty -Path $RegKey -Name 'PauseUpdatesExpiryTime'      -ErrorAction SilentlyContinue).PauseUpdatesExpiryTime
$startT = (Get-ItemProperty -Path $RegKey -Name 'PauseFeatureUpdatesStartTime' -ErrorAction SilentlyContinue).PauseFeatureUpdatesStartTime
$svc    = Get-CimInstance Win32_Service -Filter "Name='wuauserv'" -ErrorAction SilentlyContinue
$svcDisabled = ($svc -and $svc.StartMode -eq 'Disabled')
$auNoUpd = ((Get-ItemProperty -Path $AUKey -Name 'NoAutoUpdate' -ErrorAction SilentlyContinue).NoAutoUpdate -eq 1)
$isPaused = ([bool]$expiry) -or ([bool]$startT) -or $auNoUpd -or $svcDisabled

if ($isPaused) {
    # ---- RESUME / DEVAM (her iki yöntemi de geri al) ----
    if (-not (Test-Path $RegKey)) { New-Item -Path $RegKey -Force | Out-Null }
    foreach ($v in $PauseValues) {
        Remove-ItemProperty -Path $RegKey -Name $v -ErrorAction SilentlyContinue
    }
    New-ItemProperty -Path $RegKey -Name 'FlightSettingsMaxPauseDays' -Value 35 -PropertyType DWord -Force | Out-Null

    Set-NoAutoUpdate 0      # otomatik OS güncelleştirmelerini tekrar aç
    Remove-ResumeTask
    Enable-UpdateService    # eski sürümlerde devre dışı bırakılmış olabilir
    Restart-UpdateService

    Show-Msg $L_ResumeMsg $InfoModal
    Refresh-Settings
}
else {
    # ---- PAUSE / DURAKLAT ----
    # Geçerli bir değer girilene (ya da İptal'e basılana) kadar tekrar sor
    $weeks = 0
    while ($true) {
        $inp = [Microsoft.VisualBasic.Interaction]::InputBox(($L_PausePrompt -f $MaxWeeks), $L_PromptTitle, '4')
        if ([string]::IsNullOrWhiteSpace($inp)) { exit }   # İptal / Cancelled
        if ([int]::TryParse($inp.Trim(), [ref]$weeks) -and $weeks -ge 1 -and $weeks -le $MaxWeeks) { break }
        Show-Msg ($L_Invalid -f $MaxWeeks) $WarnModal
    }
    $days = $weeks * 7

        if ($supportsPauseApi) {
            # Modern Windows 10/11: yerleşik Pause API'si
            if (-not (Test-Path $RegKey)) { New-Item -Path $RegKey -Force | Out-Null }
            $now    = Get-Date
            $nowUTC = $now.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            $endUTC = $now.AddDays($days).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

            Set-ItemProperty -Path $RegKey -Name 'PauseFeatureUpdatesStartTime' -Value $nowUTC -Type String
            Set-ItemProperty -Path $RegKey -Name 'PauseFeatureUpdatesEndTime'   -Value $endUTC -Type String
            Set-ItemProperty -Path $RegKey -Name 'PauseQualityUpdatesStartTime' -Value $nowUTC -Type String
            Set-ItemProperty -Path $RegKey -Name 'PauseQualityUpdatesEndTime'   -Value $endUTC -Type String
            Set-ItemProperty -Path $RegKey -Name 'PauseUpdatesStartTime'        -Value $nowUTC -Type String
            Set-ItemProperty -Path $RegKey -Name 'PauseUpdatesExpiryTime'       -Value $endUTC -Type String

            $maxDays = [Math]::Max($days, 35)
            New-ItemProperty -Path $RegKey -Name 'FlightSettingsMaxPauseDays' -Value $maxDays -PropertyType DWord -Force | Out-Null

            Restart-UpdateService
        }
        else {
            # Windows Server 2016 / eski sürümler: otomatik OS güncelleştirmelerini
            # ilkeyle durdur (hizmet çalışmaya devam eder -> Defender güncel kalır)
            Set-NoAutoUpdate 1
            Remove-ResumeTask
            try { Register-ResumeTask -days $days } catch {}
            Restart-UpdateService   # ilkenin uygulanması için yenile
        }

        Show-Msg ($L_PausedTmpl -f $weeks, $days) $InfoModal
        Refresh-Settings
}
