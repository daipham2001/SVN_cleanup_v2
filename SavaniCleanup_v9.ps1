param(
    [string]$ConfigPath = "C:\IT_Scripts\cleanup_config.json",
    [string]$LogFile    = "C:\IT_Scripts\Cleanup_Log.txt",
    [string]$ExternalToken = $null,
    [string]$ExternalChatID = $null
)
$CurrentVersion = 11
# ===========================================================
#  SAVANI IT CLEANUP V9.1 FINAL - Production Ready
#  100 chi nhanh ban hang - Offline Ready
# ===========================================================

# -- 0. LOG ROTATION -----------------------------------------
try {
    if (Test-Path $LogFile -ErrorAction SilentlyContinue) {
        $logItem = Get-Item $LogFile -ErrorAction SilentlyContinue
        if ($logItem.Length -gt 5MB) {
            $oldLog = $LogFile.Replace(".txt", "_old.txt")
            if (Test-Path $oldLog -ErrorAction SilentlyContinue) {
                Remove-Item $oldLog -Force -ErrorAction SilentlyContinue
            }
            Rename-Item -Path $LogFile -NewName (Split-Path $oldLog -Leaf) -Force -ErrorAction SilentlyContinue
        }
    }
} catch { }

# -- 1. MUTEX ------------------------------------------------
$mutexName = "Global\SavaniITCleanupMutex"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)
if (-not $mutex.WaitOne(0, $false)) { exit }

$scriptStartTime = [DateTime]::Now
$timeoutMinutes  = 10  # Giam tu 25 xuong 10 phut de bao ve gio lam viec

# -- HAM TIEN ICH (DI CHUYEN LEN TRUOC) ----------------------

function Write-Log {
    param([string]$msg, [string]$level = "INFO")
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')][$level] $msg"
    $line | Out-File -FilePath $LogFile -Append -Encoding utf8
}

function Test-Timeout {
    $elapsed = ([DateTime]::Now - $scriptStartTime).TotalMinutes
    if ($elapsed -ge $timeoutMinutes) {
        Write-Log "TIMEOUT sau $([math]::Round($elapsed,1)) phut. Dung de bao ve hieu nang." "WARN"
        return $true
    }
    return $false
}

function Get-FolderSize {
    param([string]$path)
    if (-not (Test-Path $path -ErrorAction SilentlyContinue)) { return [long]0 }
    $s = Get-ChildItem $path -Recurse -File -ErrorAction SilentlyContinue |
         Measure-Object -Property Length -Sum
    if ($null -eq $s.Sum) { return [long]0 }
    return [long]$s.Sum
}

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Add-Type -AssemblyName Microsoft.VisualBasic
    # =========================================================
    # --- MODULE AUTO UPDATE (ZERO-AGENT) ---
    # =========================================================
    $UpdateUrl = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("aHR0cHM6Ly9yYXcuZ2l0aHVidXNlcmNvbnRlbnQuY29tL2RhaXBoYW0yMDAxL1NWTl9jbGVhbnVwL21haW4vU2F2YW5pQ2xlYW51cF92OS5wczE="))
    
    try {
        Write-Log "AUTO-UPDATE: Bat dau kiem tra phien ban moi..."
        
        $cacheBuster = [guid]::NewGuid().ToString()
        $fetchUrl = "$($UpdateUrl.Trim())?t=$cacheBuster"

        # Đâm thẳng vào tải code. Nếu rớt mạng tự nó nhảy xuống catch, không cần test mạng!
        $resp = Invoke-WebRequest -Uri $fetchUrl -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
        $remoteScript = $resp.Content

        if ($remoteScript -match '(?im)^\s*\$CurrentVersion\s*=\s*["'']?([0-9]+(\.[0-9]+)*)') {
            $remoteVersion = [double]$matches[1]

            if ($remoteVersion -gt $CurrentVersion) {
                Write-Log "AUTO-UPDATE: Co ban moi V$remoteVersion (Hien tai: V$CurrentVersion). Dang cap nhat..."

                # Fix triệt để vụ file tàng hình của Task Scheduler
                $scriptPath = $PSCommandPath
                if ([string]::IsNullOrEmpty($scriptPath)) { 
                    $scriptPath = "C:\IT_Scripts\SavaniCleanup_v9.ps1" 
                }

                $remoteScript | Out-File -FilePath $scriptPath -Encoding utf8 -Force
                Write-Log "AUTO-UPDATE: Ghi de thanh cong. Dang reboot kich ban..."

                if ($null -ne $mutex) { try { $mutex.ReleaseMutex() } catch { } $mutex.Dispose() }

                Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$scriptPath`""
                exit
            } else {
                Write-Log "AUTO-UPDATE: Dang dung ban moi nhat (V$CurrentVersion)."
            }
        } else {
            Write-Log "AUTO-UPDATE: Khong tim thay bien CurrentVersion tren GitHub" "WARN"
        }
    } catch {
        Write-Log "AUTO-UPDATE Bo qua (Loi mang/Link): $($_.Exception.Message)" "WARN"
    }
    # =========================================================
    # =========================================================

    # -- 2. DOC CONFIG ----------------------------------------
    if (-not (Test-Path $ConfigPath -ErrorAction SilentlyContinue)) { exit }
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

    # -- 3. DOC BOT TOKEN ------------------------------------
    $botToken = $null
    
    # Uu tien 1: Lay Token tu file .bat truyen vao (Neu co)
    if (-not [string]::IsNullOrEmpty($ExternalToken)) {
        $botToken = $ExternalToken
        Write-Log "Dung Bot Token tu tham so truyen vao (.bat)"
    } 
    # Uu tien 2: Neu khong co (chay tu dong), thi moi doc file ma hoa
    else {
        $encFile = "C:\IT_Scripts\.tg_token.enc"
        if (Test-Path $encFile) {
            try {
                $aesKey = [byte[]](1..32)
                $encryptedData = Get-Content $encFile -Raw
                $secureString = ConvertTo-SecureString $encryptedData -Key $aesKey
                $botToken = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString))
            } catch { 
                $botToken = $null 
                Write-Log "Loi giai ma Token: $($_.Exception.Message)" "ERROR"
            }
        }
    }


    # -- BIEN TONG HOP ----------------------------------------
    $globalDeletedSize   = [long]0
    $global:zaloTotal    = [long]0
    $global:downTotal    = [long]0
    $global:tempTotal    = [long]0
    $global:rbTotal      = [long]0
    $global:desktopSize  = [long]0
    $global:docsSize     = [long]0
    $global:skippedCount = 0
    $global:timeoutHit   = $false

    if ($config.Options.DryRun) {
        $globalStatus = "DRY-RUN (Mo phong)"
    } else {
        $globalStatus = "LIVE (Xoa that)"
    }

    function Log-And-Sum {
        param($files, [string]$label, [string]$category)
        $count = ($files | Measure-Object).Count
        $sumResult = ($files | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $sumResult) { $sumResult = [long]0 }
        $sum = [long]$sumResult

        $global:globalDeletedSize += $sum
        switch ($category) {
            "Zalo"       { $global:zaloTotal += $sum }
            "Downloads"  { $global:downTotal += $sum }
            "Temp"       { $global:tempTotal += $sum }
            "RecycleBin" { $global:rbTotal   += $sum }
        }
        $sizeMB = [math]::Round($sum / 1MB, 2)
        if ($count -gt 0) { Write-Log "    |-- $label : $count file (~$sizeMB MB)" }
    }

    function Remove-Smart {
        param($item, [bool]$useRecycleBin = $true)
        try {
            if ($config.Options.DryRun) { return }
            if ($useRecycleBin) {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                    $item.FullName, 'OnlyErrorDialogs', 'SendToRecycleBin'
                )
            } else {
                $safe = "\\?\$($item.FullName)"
                Remove-Item -LiteralPath $safe -Force -Recurse -ErrorAction SilentlyContinue
            }
        } catch {
            $global:skippedCount++
            Write-Log "Bo qua: $($item.FullName) - $($_.Exception.Message)" "WARN"
        }
    }

    # -- HAM GUI TELEGRAM ------------------------------------
    function Send-Telegram {
        param([string]$message, [string]$context = "REPORT")

        if (-not $config.Telegram.Enabled) { return }
        if ([string]::IsNullOrEmpty($botToken)) {
            Write-Log "Telegram: Khong tim thay Bot Token." "WARN"
            return
        }

        $maxRetries = 6
        $retryCount = 0
        $connected  = $false

        # Test ket noi bang HTTP thay vi ping (tranh bi chan ICMP)
        while (-not $connected -and $retryCount -lt $maxRetries) {
            try {
                Invoke-WebRequest -Uri "https://api.telegram.org" -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop | Out-Null
                $connected = $true
            } catch {
                $retryCount++
                Write-Log "Telegram: Chua co mang, thu lai $retryCount/$maxRetries..." "WARN"
                Start-Sleep -Seconds 15
            }
        }

        if (-not $connected) {
            Write-Log "Telegram: Khong co mang sau $maxRetries lan thu." "ERROR"
            return
        }

        try {
            $url  = "https://api.telegram.org/bot$botToken/sendMessage"
            $body = @{
                chat_id    = $config.Telegram.ChatID
                text       = $message
                parse_mode = "HTML"
            }
            Invoke-RestMethod -Uri $url -Method Post -Body $body -TimeoutSec 30 -ErrorAction Stop | Out-Null
            Write-Log "Telegram [$context]: OK"
        } catch {
            Write-Log "Telegram [$context]: Loi - $($_.Exception.Message)" "ERROR"
        }
    }

    function Send-ErrorAlert {
        param([string]$errorMsg, [string]$section = "UNKNOWN")
        $alert = "LOI - SAVANI CLEANUP`n" +
                 "May: <b>$($env:COMPUTERNAME)</b>`n" +
                 "Muc: $section`n" +
                 "Chi tiet: <code>$errorMsg</code>`n" +
                 "<i>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</i>"
        Send-Telegram $alert "ERROR-ALERT"
    }

    # -- KIEM TRA O C ----------------------------------------
    $diskC = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    if ($null -eq $diskC) {
        Send-ErrorAlert "Khong lay duoc thong tin o C:" "DISK-CHECK"
        exit
    }

    $freeGBBefore   = [math]::Round($diskC.FreeSpace / 1GB, 2)
    $totalGB        = [math]::Round($diskC.Size / 1GB, 2)
    $aggressiveMode = $freeGBBefore -le $config.Threshold.AggressiveFreeGB

    if ($freeGBBefore -gt $config.Threshold.MinFreeGBToRun) { exit }

    Write-Log "========================================="
    Write-Log "SAVANI CLEANUP V$CurrentVersion - START"
    Write-Log "O C: $freeGBBefore GB / $totalGB GB"
    if ($aggressiveMode) {
        Write-Log "MODE: AGGRESSIVE" "WARN"
    } else {
        Write-Log "MODE: NORMAL"
    }
    Write-Log "STATUS: $globalStatus"
    Write-Log "========================================="

    # Tinh nguong ngay theo mode
    if ($aggressiveMode) {
        $daysTemp = [math]::Max(1, [int]($config.Days.Temp      / 3))
        $daysDown = [math]::Max(3, [int]($config.Days.Downloads / 3))
    } else {
        $daysTemp = $config.Days.Temp
        $daysDown = $config.Days.Downloads
    }

    $limitTemp = (Get-Date).AddDays(-$daysTemp)
    $limitDown = (Get-Date).AddDays(-$daysDown)

    if ($aggressiveMode) {
        Write-Log "AGGRESSIVE: O C con $freeGBBefore GB. Temp: $daysTemp ngay, Down: $daysDown ngay." "WARN"
        Send-Telegram ("AGGRESSIVE MODE`n" +
                       "May: $($env:COMPUTERNAME)`n" +
                       "O C con $freeGBBefore GB / $totalGB GB`n" +
                       "Dang dung nguong xoa manh de giai phong khong gian.") "AGGRESSIVE-ALERT"
    }

    $processCount = 0

    # -- 2. SYSTEM TEMP --------------------------------------
    try {
        $sysTemp = "$env:windir\Temp"
        if (Test-Path $sysTemp -ErrorAction SilentlyContinue) {
            $files = Get-ChildItem $sysTemp -File -Recurse -ErrorAction SilentlyContinue |
                     Where-Object { $_.LastWriteTime -lt $limitTemp }
            Log-And-Sum $files "System Temp" "Temp"
            foreach ($f in $files) {
                if (Test-Timeout) { $global:timeoutHit = $true; break }
                Remove-Smart $f $false
                $processCount++
                if ($processCount % 50 -eq 0) { Start-Sleep -Milliseconds 10 }
            }
        }
    } catch {
        Write-Log "Loi System Temp: $($_.Exception.Message)" "ERROR"
        Send-ErrorAlert $_.Exception.Message "SYSTEM-TEMP"
    }

    # -- 3. USER PROFILES ------------------------------------
    try {
        $users = Get-CimInstance Win32_UserProfile | Where-Object { $_.Special -eq $false }
        foreach ($profile in $users) {
            if ($global:timeoutHit) { break }

            $userDir  = $profile.LocalPath
            $userName = Split-Path $userDir -Leaf
            if ($config.ExcludeUsers -contains $userName) { continue }

            Write-Log "[*] Profile: $userName"

            try { $global:desktopSize += Get-FolderSize (Join-Path $userDir "Desktop") } catch { }
            try { $global:docsSize    += Get-FolderSize (Join-Path $userDir "Documents") } catch { }

            # Zalo
            foreach ($zPath in $config.Paths.Zalo) {
                if ($global:timeoutHit) { break }
                try {
                    $fullPath = Join-Path $userDir $zPath
                    if (-not (Test-Path $fullPath -ErrorAction SilentlyContinue)) { continue }
                    $targets = Get-ChildItem $fullPath -Directory -Recurse -Depth 2 -ErrorAction SilentlyContinue |
                               Where-Object { $config.JunkFolders -contains $_.Name }
                    foreach ($t in $targets) {
                        if ($global:timeoutHit) { break }
                        $files = Get-ChildItem $t.FullName -File -Recurse -ErrorAction SilentlyContinue |
                                 Where-Object { $_.LastWriteTime -lt $limitTemp }
                        Log-And-Sum $files "Zalo ($($t.Name))" "Zalo"
                        foreach ($f in $files) {
                            if (Test-Timeout) { $global:timeoutHit = $true; break }
                            Remove-Smart $f $false
                            $processCount++
                            if ($processCount % 50 -eq 0) { Start-Sleep -Milliseconds 10 }
                        }
                    }
                } catch {
                    Write-Log "Loi Zalo ($userName): $($_.Exception.Message)" "WARN"
                }
            }

            # Downloads
            try {
                $downDir = Join-Path $userDir "Downloads"
                if (Test-Path $downDir -ErrorAction SilentlyContinue) {
                    $files = Get-ChildItem $downDir -File -ErrorAction SilentlyContinue |
                             Where-Object {
                                 $_.LastWriteTime -lt $limitDown -and
                                 $config.SafeExtensions -notcontains $_.Extension.ToLower()
                             }
                    Log-And-Sum $files "Downloads" "Downloads"
                    foreach ($f in $files) {
                        if (Test-Timeout) { $global:timeoutHit = $true; break }
                        Remove-Smart $f $true
                        $processCount++
                        if ($processCount % 50 -eq 0) { Start-Sleep -Milliseconds 10 }
                    }
                }
            } catch {
                Write-Log "Loi Downloads ($userName): $($_.Exception.Message)" "WARN"
            }

            # User Temp
            try {
                $tempDir = Join-Path $userDir "AppData\Local\Temp"
                if (Test-Path $tempDir -ErrorAction SilentlyContinue) {
                    $files = Get-ChildItem $tempDir -File -Recurse -ErrorAction SilentlyContinue |
                             Where-Object { $_.LastWriteTime -lt $limitTemp }
                    Log-And-Sum $files "User Temp" "Temp"
                    foreach ($f in $files) {
                        if (Test-Timeout) { $global:timeoutHit = $true; break }
                        Remove-Smart $f $false
                        $processCount++
                        if ($processCount % 50 -eq 0) { Start-Sleep -Milliseconds 10 }
                    }
                }
            } catch {
                Write-Log "Loi UserTemp ($userName): $($_.Exception.Message)" "WARN"
            }
        }
    } catch {
        Write-Log "Loi USER PROFILES: $($_.Exception.Message)" "ERROR"
        Send-ErrorAlert $_.Exception.Message "USER-PROFILES"
    }

    # -- 4. RECYCLE BIN --------------------------------------
    try {
        $rbSizeBefore = [long]0
        $rbCount      = 0

        $sa = New-Object -ComObject Shell.Application
        $rb = $sa.Namespace(0x0a)
        if ($null -ne $rb) {
            foreach ($item in $rb.Items()) {
                try {
                    if ($item.ModifyDate -lt $limitTemp) {
                        $rbSizeBefore += $item.Size
                        $rbCount++
                    }
                } catch { }
            }
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sa) | Out-Null

        $global:rbTotal           += $rbSizeBefore
        $global:globalDeletedSize += $rbSizeBefore

        $rbMB = [math]::Round($rbSizeBefore / 1MB, 2)
        if ($rbCount -gt 0) { Write-Log "    |-- Recycle Bin: $rbCount item (~$rbMB MB)" }

        if (-not $config.Options.DryRun -and $rbCount -gt 0) {
            Clear-RecycleBin -Force -ErrorAction SilentlyContinue
            Write-Log "Recycle Bin: Da xoa $rbCount item."
        }
    } catch {
        Write-Log "Loi Recycle Bin: $($_.Exception.Message)" "WARN"
        Send-ErrorAlert $_.Exception.Message "RECYCLE-BIN"
    }

    # -- 5. TELEGRAM REPORT ----------------------------------
    $finalDisks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
    $diskInfo   = ""
    foreach ($d in $finalDisks) {
        $f = [math]::Round($d.FreeSpace / 1GB, 2)
        $t = [math]::Round($d.Size / 1GB, 2)
        $p = [math]::Round(($f / $t) * 100, 1)
        $diskInfo += "  $($d.DeviceID) : $f GB / $t GB ($p% Free)`n"
    }

    $totMB   = [math]::Round($global:globalDeletedSize / 1MB, 2)
    $elapsed = [math]::Round(([DateTime]::Now - $scriptStartTime).TotalMinutes, 1)

    $extraTag = ""
    if ($global:timeoutHit)         { $extraTag += "`nCANH BAO: Timeout $timeoutMinutes phut - dung som." }
    if ($global:skippedCount -gt 0) { $extraTag += "`nBo qua: $($global:skippedCount) file loi (xem log)." }
    if ($aggressiveMode)            { $extraTag += "`nAGGRESSIVE MODE - O C can, nguong xoa manh hon." }

    $msg = "🚀 SAVANI CLEANUP V$CurrentVersion - $($env:COMPUTERNAME)`n" +
           "------------------------`n" +
           "Trang thai : $globalStatus$extraTag`n" +
           "Tong don   : $totMB MB`n" +
           "Thoi gian  : $elapsed phut`n" +
           "------------------------`n" +
           "Chi tiet xoa:`n" +
           "  Zalo      : $([math]::Round($global:zaloTotal /1MB,2)) MB`n" +
           "  Downloads : $([math]::Round($global:downTotal /1MB,2)) MB`n" +
           "  Temp      : $([math]::Round($global:tempTotal /1MB,2)) MB`n" +
           "  Recycle   : $([math]::Round($global:rbTotal   /1MB,2)) MB`n" +
           "------------------------`n" +
           "Giam sat:`n" +
           "  Desktop   : $([math]::Round($global:desktopSize/1GB,2)) GB`n" +
           "  Documents : $([math]::Round($global:docsSize   /1GB,2)) GB`n" +
           "------------------------`n" +
           "O dia sau don:`n$diskInfo" +
           "------------------------`n" +
           "Savani Operations - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

    Send-Telegram $msg "REPORT"
    Write-Log "DONE. Tong: $totMB MB | $elapsed phut."

} catch {
    $fatalMsg = $_.Exception.Message
    Write-Log "FATAL: $fatalMsg" "ERROR"
    try {
        if ($null -ne $config -and $config.Telegram.Enabled -and -not [string]::IsNullOrEmpty($botToken)) {
            $crashMsg = "LOI NGHIEM TRONG - SAVANI CLEANUP`n" +
                        "May: $($env:COMPUTERNAME)`n" +
                        "Loi: $fatalMsg`n" +
                        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $url  = "https://api.telegram.org/bot$botToken/sendMessage"
            $body = @{ chat_id = $config.Telegram.ChatID; text = $crashMsg; parse_mode = "HTML" }
            Invoke-RestMethod -Uri $url -Method Post -Body $body -TimeoutSec 20 -ErrorAction SilentlyContinue | Out-Null
        }
    } catch { }

} finally {
    if ($null -ne $mutex) {
        try { $mutex.ReleaseMutex() } catch { }
        $mutex.Dispose()
    }
}

