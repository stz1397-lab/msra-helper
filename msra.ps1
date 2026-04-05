# =============================================
# Умное подключение MSRA v2.18
# Поддержка: trueconf, pacs, ping, история и т.д.
# =============================================

$scriptVersion = "2.18"# --- Подключение по IP ---

# Настройки истории
$historyFile = Join-Path $PSScriptRoot "msra_history.log"
$maxHistoryEntries = 1000

# Список известных подсетей
$knownSubnets = @(2, 3, 11, 13, 14, 15, 17, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 31, 32, 33, 35, 36, 37, 38, 48, 56, 57, 60, 61, 63, 91, 92, 96, 97, 98, 102, 103, 104, 105, 106, 107, 110, 111, 112, 113, 114, 122, 123, 124, 125, 126, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 214, 215)

# --- Функции ---

function Show-Help {
    Clear-Host
    Write-Host "╔══════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║           Справка по командам        ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Дополнительные команды:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "• update    - Обновить скрипт до последней версии" -ForegroundColor Gray
    Write-Host "• trueconf  - Открыть админку TrueConf" -ForegroundColor Gray
    Write-Host "• pacs      - Открыть админку PACS" -ForegroundColor Gray
    Write-Host "• glpi      - Открыть GLPI" -ForegroundColor Gray
    Write-Host "• scanpass  - Пароль от учётки Scan" -ForegroundColor Gray
    Write-Host "• distr     - Открыть папку дистрибутивов" -ForegroundColor Gray
    Write-Host "• restart   - Перезагрузка компьютера" -ForegroundColor Gray
    Write-Host "• exit      - Выход из программы" -ForegroundColor Gray
    Write-Host ""
    Read-Host "Нажмите Enter, чтобы вернуться в меню"
}

function Show-History {
    if (Test-Path $historyFile) {
        $history = Get-Content $historyFile
        if ($history.Count -gt 0) {
            Write-Host "`nПоследние подключения:" -ForegroundColor Yellow
            Write-Host ""
            $history | Select-Object -Last 10 | ForEach-Object {
                $line = $_
                $baseColor = if ($line -match "Успешно") { "Green" } else { "Red" }
                
                if ($line -match '(.+)(\s+\|\s+Подключение\s+#\d+)$') {
                    $mainPart = $matches[1]
                    $suffix  = $matches[2]
                    Write-Host " $mainPart" -ForegroundColor $baseColor -NoNewline
                    Write-Host $suffix -ForegroundColor Yellow
                } else {
                    Write-Host " $line" -ForegroundColor $baseColor
                }
            }
            Write-Host ""
        }
    }
}

function Get-HistoryCount {
    if (Test-Path $historyFile) {
        return (Get-Content $historyFile).Count
    }
    return 0
}

function Clear-HistoryFile {
    if (Test-Path $historyFile) {
        Clear-Content $historyFile
        Write-Host "История подключений очищена." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
    } else {
        Write-Host "Файл истории не найден." -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

function Remove-FailedHistoryEntries {
    if (Test-Path $historyFile) {
        $history = Get-Content $historyFile
        $filteredHistory = $history | Where-Object { $_ -notmatch "Нет пинга" }

        if ($filteredHistory.Count -eq $history.Count) {
            Write-Host "Нет неудачных записей для удаления." -ForegroundColor Yellow
        } else {
            $removedCount = $history.Count - $filteredHistory.Count
            $filteredHistory | Set-Content $historyFile -Encoding utf8
            Write-Host "Удалено $removedCount неудачных записей из истории." -ForegroundColor Red
        }
        Start-Sleep -Seconds 2
    } else {
        Write-Host "Файл истории не найден." -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

function Remove-DuplicateHistoryEntries {
    if (Test-Path $historyFile) {
        $history = @(Get-Content $historyFile)
        $filteredHistory = @()
        $prevTarget = ""
        $prevStatus = ""
        $removedCount = 0

        foreach ($line in $history) {
            $fields = $line.Trim() -split '\s*\|\s*'
            if ($fields.Count -ge 4) {
                $target = $fields[2].Trim()
                $status = $fields[3].Trim()
                
                # Оставляем только если адрес или статус отличаются от предыдущего
                if ($target -ne $prevTarget -or $status -ne $prevStatus) {
                    $filteredHistory += $line
                } else {
                    $removedCount++
                }
                $prevTarget = $target
                $prevStatus = $status
            } else {
                $filteredHistory += $line
                $prevTarget = ""
                $prevStatus = ""
            }
        }

        if ($removedCount -eq 0) {
            Write-Host "Нет подряд идущих дубликатов для удаления." -ForegroundColor Yellow
        } else {
            $filteredHistory | Set-Content $historyFile -Encoding utf8
            Write-Host "Удалено $removedCount подряд идущих дубликатов из истории." -ForegroundColor Green
        }
        Start-Sleep -Seconds 2
    } else {
        Write-Host "Файл истории не найден." -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

function Test-ConsecutiveDuplicates {
    if (Test-Path $historyFile) {
        # 🔑 @() гарантирует массив даже при 1 строке
        $history = @(Get-Content $historyFile)
        $prevTarget = ""
        $prevStatus = ""
        
        foreach ($line in $history) {
            $fields = $line.Trim() -split '\s*\|\s*'
            # ✅ Принимаем 4 или 5 полей (с номером подключения их 5)
            if ($fields.Count -ge 4) {
                $target = $fields[2].Trim()
                $status = $fields[3].Trim()
                
                # Сравниваем ТОЛЬКО адрес и статус, игнорируя номер подключения
                if ($target -eq $prevTarget -and $status -eq $prevStatus) {
                    return $true
                }
                $prevTarget = $target
                $prevStatus = $status
            } else {
                $prevTarget = ""
                $prevStatus = ""
            }
        }
    }
    return $false
}


function Show-FullHistory {
    if (Test-Path $historyFile) {
        $history = Get-Content $historyFile
        if ($history.Count -gt 0) {
            $history | ForEach-Object {
                $line = $_
                $baseColor = if ($line -match "Успешно") { "Green" } else { "Red" }
                
                if ($line -match '(.+)(\s+\|\s+Подключение\s+#\d+)$') {
                    $mainPart = $matches[1]
                    $suffix  = $matches[2]
                    Write-Host " $mainPart" -ForegroundColor $baseColor -NoNewline
                    Write-Host $suffix -ForegroundColor Yellow
                } else {
                    Write-Host " $line" -ForegroundColor $baseColor
                }
            }

            $total = $history.Count
            $success = ($history -match "Успешно").Count
            $failed = $total - $success
            Write-Host "`nВсего $total записей:" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Успешных: $success" -ForegroundColor Green -NoNewline
            Write-Host " | Неудачных: $failed" -ForegroundColor Red

            $hasFailed = $failed -gt 0
            $hasDuplicates = Test-ConsecutiveDuplicates

            Write-Host "`nДополнительные действия:" -ForegroundColor Yellow
            if ($hasFailed) { Write-Host "Del     - Удалить все неудачные подключения" }
            if ($hasDuplicates) { Write-Host "Deldbl  - Удалить подряд идущие дубликаты" }
            if ($hasFailed -and $hasDuplicates) { Write-Host "Delall  - Удалить и неудачные, и дубликаты" }
            Write-Host "Enter   - Вернуться в меню"

            $choice = Read-Host "`nВыбор"
            switch ($choice.Trim().ToLower()) {
                { $_ -in "del", "вуд" } { Remove-FailedHistoryEntries }
                "deldbl" { Remove-DuplicateHistoryEntries }
                { $_ -in "delall", "вудфдд" } {
                    if (Test-Path $historyFile) {
                        $historyBefore = Get-Content $historyFile
                        $failedCount = ($historyBefore -match "Нет пинга").Count
                        if ($failedCount -gt 0) { Remove-FailedHistoryEntries }
                        if (Test-ConsecutiveDuplicates) { Remove-DuplicateHistoryEntries }
                        $historyAfter = Get-Content $historyFile
                        $removedTotal = $historyBefore.Count - $historyAfter.Count
                        if ($removedTotal -gt 0) { Write-Host "Всего удалено записей: $removedTotal" -ForegroundColor Cyan }
                        else { Write-Host "Нечего удалять." -ForegroundColor Yellow }
                        Start-Sleep -Seconds 2
                    } else {
                        Write-Host "Файл истории не найден." -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
                default { }
            }
        } else {
            Write-Host "История пуста." -ForegroundColor Yellow
            Read-Host "`nНажмите Enter, чтобы вернуться в меню"
        }
    } else {
        Write-Host "Файл истории не найден." -ForegroundColor Red
        Read-Host "`nНажмите Enter, чтобы вернуться в меню"
    }
}

function Add-HistoryEntry {
    param([string]$target, [bool]$isHostname, [bool]$success, [int]$connectionCount = 0)
    
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $type = if ($isHostname) { "Хост" } else { "IP" }
    $st = if ($success) { "Успешно" } else { "Нет пинга" }
    $suf = if ($success -and $connectionCount -gt 0) { " | Подключение #$($connectionCount + 1)" } else { "" }
    $entry = "$ts | $type | $target | $st$suf"
    
    # 🔑 КРИТИЧНО: @() гарантирует, что $cur ВСЕГДА будет массивом, даже если в файле 1 строка или 0
    $cur = @(if (Test-Path $historyFile) { Get-Content $historyFile } else { })
    
    # Добавляем, убираем пустые, оставляем последние N
    $newHistory = ($cur + $entry) | Where-Object { $_ -ne "" } | Select-Object -Last $maxHistoryEntries
    $newHistory -join "`n" | Out-File $historyFile -Encoding utf8
}

function Get-ConnectionCount {
    param([string]$target)
    if (-not (Test-Path $historyFile)) { return 0 }
    
    $count = 0
    foreach ($line in Get-Content $historyFile) {
        # Пропускаем строки без статуса "Успешно"
        if ($line -notmatch 'Успешно') { continue }
        
        # Разбиваем и берём 3-е поле (адрес)
        $parts = $line.Trim() -split '\s*\|\s*'
        if ($parts.Count -ge 3) {
            $histTarget = $parts[2].Trim()
            # Используем .Contains() вместо -like: быстрее и надёжнее для подстрок
            if ($histTarget -eq $target -or $histTarget.Contains($target)) {
                $count++
            }
        }
    }
    return $count
}

function Get-HostnameAndIP {
    param([string]$target)
    try {
        $hostEntry = [System.Net.Dns]::GetHostEntry($target)
        return @{
            Hostname = $hostEntry.HostName.Split('.')[0]
            IP = $hostEntry.AddressList | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1 | ForEach-Object { $_.ToString() }
        }
    } catch {
        return @{ Hostname = $null; IP = $null }
    }
}

function Get-PossibleSubnets {
    param(
        [string]$number,
        [array]$knownSubnets
    )
    $possibleSubnets = @()
    foreach ($subnet in $knownSubnets | Sort-Object -Descending) {
        $subnetStr = $subnet.ToString()
        if ($number.StartsWith($subnetStr)) {
            $remainingDigits = $number.Substring($subnetStr.Length)
            if ($remainingDigits -match '^(0|[1-9]\d*)$' -and [int]$remainingDigits -le 255) {
                $possibleSubnets += $subnet
            }
        }
    }
    return $possibleSubnets | Sort-Object -Unique
}

function Update-KnownSubnetsInScript {
    param([int]$newSubnet)
    $scriptPath = $PSCommandPath
    if ([string]::IsNullOrEmpty($scriptPath)) {
        Write-Host "Автоматическое обновление возможно только при запуске скрипта из файла! Создайте ярлык на .ps1 файл и попробуйте снова" -ForegroundColor Red
        return
    }
    $scriptContent = Get-Content $scriptPath -Raw
    $pattern = '(?ms)^\$knownSubnets\s*=\s*@\((.*?)\)'
    if ($scriptContent -match $pattern) {
        $current = $matches[1]
        $arr = $current -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }
        if (-not ($arr -contains "$newSubnet")) {
            $arr += "$newSubnet"
            $arr = $arr | Sort-Object { [int]$_ }
            $newLine = '$knownSubnets = @(' + ($arr -join ', ') + ')'
            $scriptContent = [regex]::Replace($scriptContent, $pattern, $newLine)
            Set-Content -Path $scriptPath -Value $scriptContent -Encoding UTF8
            Write-Host "Скрипт обновлён: подсеть $newSubnet добавлена в список известных подсетей." -ForegroundColor Cyan
            Start-Sleep -Seconds 3
        }
    } else {
        Write-Host "Не удалось найти knownSubnets в скрипте!" -ForegroundColor Red
    }
}


function Test-TCPPortWithFeedback {
    param(
        [string]$Target,
        [int[]]$Ports = @(135, 445),
        [int]$TimeoutMs = 1000
    )
    Write-Host "`nПроверка доступности через TCP..." -ForegroundColor Magenta

    foreach ($port in $Ports) {
        Write-Host "  Проверка порта $port..." -ForegroundColor Gray -NoNewline

        $tcp = $null
        $result = $false

        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $asyncResult = $tcp.BeginConnect($Target, $port, $null, $null)
            $waitHandle = $asyncResult.AsyncWaitHandle

            try {
                if ($waitHandle.WaitOne($TimeoutMs, $false)) {
                    $tcp.EndConnect($asyncResult)
                    $result = $true
                }
            } finally {
                $waitHandle.Close()
            }
        } catch {
            $result = $false
        } finally {
            if ($tcp) {
                try {
                    if ($tcp.Connected) { $tcp.Client.Close() }
                    $tcp.Close()
                    $tcp.Dispose()
                } catch {}
            }
        }

        if ($result) {
            Write-Host " Успешно!" -ForegroundColor Green
            return @{ Success = $true; Port = $port }
        } else {
            Write-Host " Недоступен" -ForegroundColor DarkGray
        }
    }

    Write-Host "  → Устройство недоступно по TCP (порты $($Ports -join ', '))." -ForegroundColor Red
    return @{ Success = $false; Port = $null }
}

function Test-TCPPortWithRetry {
    param(
        [string]$Target,
        [int[]]$Ports = @(135, 445),
        [int]$TimeoutMs = 1000,
        [int]$Retries = 1,      # ← уменьшено до 1 (итого 2 попытки)
        [int]$DelayMs = 1500
    )
    Write-Host "`nПроверка доступности через TCP..." -ForegroundColor Magenta

    for ($attempt = 0; $attempt -le $Retries; $attempt++) {
        if ($attempt -gt 0) {
            Write-Host "  → Повторная попытка ($attempt/$Retries)..." -ForegroundColor Yellow
            Start-Sleep -Milliseconds $DelayMs
        }

        foreach ($port in $Ports) {
            Write-Host "  Проверка порта $port..." -ForegroundColor Gray -NoNewline

            $tcp = $null
            $result = $false

            try {
                $tcp = New-Object System.Net.Sockets.TcpClient
                $asyncResult = $tcp.BeginConnect($Target, $port, $null, $null)
                $waitHandle = $asyncResult.AsyncWaitHandle

                try {
                    if ($waitHandle.WaitOne($TimeoutMs, $false)) {
                        $tcp.EndConnect($asyncResult)
                        $result = $true
                    }
                } finally {
                    $waitHandle.Close()
                }
            } catch {
                $result = $false
            } finally {
                if ($tcp) {
                    try {
                        if ($tcp.Connected) { $tcp.Client.Close() }
                        $tcp.Close()
                        $tcp.Dispose()
                    } catch {}
                }
            }

            if ($result) {
                Write-Host " Успешно!" -ForegroundColor Green
                return @{ Success = $true; Port = $port }
            } else {
                Write-Host " Недоступен" -ForegroundColor DarkGray
            }
        }
    }

    Write-Host "  → Устройство недоступно по TCP (порты $($Ports -join ', '))." -ForegroundColor Red
    return @{ Success = $false; Port = $null }
}

function Start-Ping {
    param(
        [string]$Target,
        [switch]$Continuous
    )
    Clear-Host
    Write-Host "── Запуск пинга ──" -ForegroundColor Cyan
    Write-Host ""

    if ($Continuous) {
        Write-Host "Бесконечный пинг $Target (Нажмите Q для остановки)" -ForegroundColor Yellow
        Write-Host ""
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "ping"
        $psi.Arguments = "$Target -t"
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.CreateNoWindow = $true
        $pingProcess = New-Object System.Diagnostics.Process
        $pingProcess.StartInfo = $psi

        $pingProcess.Start() | Out-Null

        while (-not $pingProcess.HasExited) {
            $line = $pingProcess.StandardOutput.ReadLine()
            if ($line -ne $null) {
                Write-Host $line
            }

            if ($Host.UI.RawUI.KeyAvailable) {
                $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                if ($key.Character -in 'q','Q','й','Й') {
                    $pingProcess.Kill()
                    Write-Host "`nПинг остановлен" -ForegroundColor Yellow
                    break
                }
            }
            Start-Sleep -Milliseconds 10
        }
    } else {
        Write-Host "Пинг $Target (4 запроса). Нажмите Q для отмены" -ForegroundColor Yellow
        Write-Host ""

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "ping"
        $psi.Arguments = "$Target -n 4"
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.CreateNoWindow = $true
        $pingProcess = New-Object System.Diagnostics.Process
        $pingProcess.StartInfo = $psi

        $pingProcess.Start() | Out-Null

        while (-not $pingProcess.StandardOutput.EndOfStream) {
            $line = $pingProcess.StandardOutput.ReadLine()
            if ($line -ne $null) {
                Write-Host $line
            }

            if ($Host.UI.RawUI.KeyAvailable) {
                $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                if ($key.Character -in 'q','Q','й','Й') {
                    $pingProcess.Kill()
                    Write-Host "`nПинг остановлен пользователем" -ForegroundColor Yellow
                    break
                }
            }
            Start-Sleep -Milliseconds 10
        }
    }

    Write-Host ""
    Read-Host "Нажмите Enter, чтобы вернуться в меню"
}

# --- Генерация примеров ---
$randomHostNum = Get-Random -Minimum 11 -Maximum 150
$hostExample = "nc-{0:d2}" -f $randomHostNum

$randomSubnet = $knownSubnets | Get-Random
$randomTailDotted = Get-Random -Minimum 10 -Maximum 99
$randomTailJoined = Get-Random -Minimum 100 -Maximum 199

$exampleDotted = "$randomSubnet.$randomTailDotted"
$exampleJoined = "$randomSubnet$randomTailJoined"
$exampleIPDotted = "192.168.$randomSubnet.$randomTailDotted"
$exampleIPJoined = "192.168.$randomSubnet.$randomTailJoined"

$randomTailPing = Get-Random -Minimum 10 -Maximum 99
$examplePing = "$randomSubnet$randomTailPing"
$examplePingIP = "192.168.$randomSubnet.$randomTailPing"

# --- Основной цикл ---
while ($true) {
    try {
        Clear-Host
        Write-Host "╔══════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║     Умное подключение MSRA v$scriptVersion     ║" -ForegroundColor Cyan
        Write-Host "╚══════════════════════════════════════╝" -ForegroundColor Cyan
        Show-History
        $allConnections = Get-HistoryCount
        Write-Host ("Всего подключений: {0}" -f $allConnections) -ForegroundColor DarkCyan
        Write-Host "`nФорматы ввода:" -ForegroundColor Yellow
        Write-Host "• $exampleDotted → $exampleIPDotted"
        Write-Host "• $exampleJoined → $exampleIPJoined"
        Write-Host "• Имя хоста (например $hostExample)"
        Write-Host "• ping hostname / $examplePing (-t)"
        Write-Host "• help → Справка по дополнительным командам"
        Write-Host "• his → Показать всю историю подключений"
        Write-Host "• clear → Очистить историю подключений"
        
        $userInput = Read-Host "Введите данные (Enter для выхода)"
        if ([string]::IsNullOrEmpty($userInput)) { break }

        # === Команды ===
        if ($userInput -eq "help" -or $userInput -eq "рудз") { Show-Help; continue }
        if ($userInput -eq "his" -or $userInput -eq "ршы") { Clear-Host; Write-Host "── Полная история подключений ──" -ForegroundColor Cyan; Show-FullHistory; continue }
        if ($userInput -eq "clear" -or $userInput -eq "сдефк") { Clear-HistoryFile; continue }

        if ($userInput -ieq "trueconf") {
            Start-Process "http://ssrv.lan.smclinic.ru/admin/general/info/"
            Write-Host "`nОткрывается страница входа TrueConf..." -ForegroundColor Green
            Set-Clipboard -Value "trueconftech"
            Write-Host "Логин скопирован." -ForegroundColor Green
            if ((Read-Host "Скопировать пароль? (y/n)") -match '^[yYдД]') { Set-Clipboard -Value "Beif8uwi"; Write-Host "Пароль скопирован." -ForegroundColor Green }
            Write-Host "Возвращаемся в меню..." -ForegroundColor Green; Start-Sleep -Seconds 2; continue
        }

        if ($userInput -ieq "update") {
            $currentScript = $PSCommandPath
            if ([string]::IsNullOrEmpty($currentScript)) { Write-Host "Обновление невозможно: скрипт запущен не из файла." -ForegroundColor Red; Read-Host "Нажмите Enter"; continue }
            $repoUrl = "https://raw.githubusercontent.com/stz1397-lab/msra-helper/main/msra.ps1"
            $tempFile = [System.IO.Path]::GetTempFileName()
            try {
                Write-Host "Загрузка новой версии..." -ForegroundColor Cyan
                $proxy = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy($repoUrl)
                if ($proxy -eq $repoUrl) { Invoke-WebRequest -Uri $repoUrl -OutFile $tempFile -UseBasicParsing -TimeoutSec 15 }
                else { Invoke-WebRequest -Uri $repoUrl -OutFile $tempFile -UseBasicParsing -TimeoutSec 15 -Proxy $proxy -ProxyUseDefaultCredentials }
                
                if ((Get-Content $currentScript -Raw) -eq (Get-Content $tempFile -Raw)) { Write-Host "У вас уже последняя версия." -ForegroundColor Green }
                else {
                    Copy-Item $currentScript "$currentScript.bak" -Force
                    Copy-Item $tempFile $currentScript -Force
                    Write-Host "Скрипт обновлён! Перезапуск..." -ForegroundColor Green; Start-Sleep -Seconds 1
                    & $currentScript; exit
                }
            } catch { Write-Host "Ошибка обновления: $_" -ForegroundColor Red }
            finally { if (Test-Path $tempFile) { Remove-Item $tempFile -Force } }
            Read-Host "Нажмите Enter, чтобы вернуться в меню"; continue
        }

        if ($userInput -ieq "glpi") { Start-Process "https://glpi.smclinic.ru/front/ticket.php"; Write-Host "GLPI запущен" -ForegroundColor Green; Start-Sleep -Seconds 2; continue }
        if ($userInput -ieq "restart") {
            if ((Get-CimInstance Win32_ComputerSystem).NumberOfUsers -gt 1) {
                Write-Warning "Есть активные пользователи. Продолжить?"
                if ((Read-Host "(y/n)") -notmatch '^[yYдД]') { Write-Host "Отмена." -ForegroundColor Yellow; Start-Sleep 2; continue }
            }
            try { Restart-Computer -Force; Write-Host "Перезагрузка..." -ForegroundColor Green } catch { Write-Host "Ошибка: $_" -ForegroundColor Red }
            Start-Sleep 2; continue
        }
        if ($userInput -ieq "exit") { exit }
        if ($userInput -ieq "scanpass") { "53807553QaZ" | Set-Clipboard; Write-Host "Пароль Scan скопирован"; Read-Host "Enter"; continue }
        if ($userInput -ieq "sigur")   { "rt54de1z"      | Set-Clipboard; Write-Host "Пароль Sigur скопирован"; Read-Host "Enter"; continue }
        if ($userInput -ieq "distr")   { Invoke-Item "\\fileserver\distr$"; continue }
        if ($userInput -ieq "pacs") {
            Start-Process "http://pacs-2.lan.smclinic.ru/pacs/login.php"
            Write-Host "PACS открыт." -ForegroundColor Green
            if ((Read-Host "Открыть инструкцию? (y/n)") -match '^[yYдД]') { Start-Process "https://conf.smclinic.ru/spaces/WTS/pages/121275145/%D0%90%D0%B4%D0%BC%D0%B8%D0%BD%D0%BA%D0%B0+PACS" }
            Start-Sleep 2; continue
        }

        # === Ping ===
        $isPing = $false; $contPing = $false
        if ($userInput -match '^ping\s+(.+?)(\s+-t)?$') {
            $isPing = $true; $contPing = $matches[2] -ne $null; $userInput = $matches[1].Trim()
        }
        if ($isPing) {
            $target = $null
            if ($userInput -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') { $target = $userInput }
            elseif ($userInput -match '^[\d\.]+$') {
                if ($userInput -match '^(\d{1,3})\.(\d{1,3})$') {
                    $o3=[int]$matches[1]; $o4=[int]$matches[2]
                    if ($o3 -gt 255 -or $o4 -gt 255) { Write-Host "Ошибка IP!" -ForegroundColor Red; Start-Sleep 2; continue }
                    $target = "192.168.$o3.$o4"
                } elseif ($userInput -match '^(\d+)$') {
                    $num = [int]$userInput.Trim()
                    if ($num -le 19) { Write-Host "Ошибка ввода" -ForegroundColor Red; Start-Sleep 2; continue }
                    $subs = Get-PossibleSubnets -number $num -knownSubnets $knownSubnets
                    if ($subs.Count -gt 1) {
                        Write-Host "`nНесколько подсетей:" -ForegroundColor Yellow
                        for($i=0;$i -lt $subs.Count;$i++) {
                            $rem = $num.ToString().Substring($subs[$i].ToString().Length)
                            Write-Host "$($i+1). 192.168.$($subs[$i]).$([int]$rem)" -ForegroundColor Gray
                        }
                        Write-Host "$($subs.Count+1). Отмена" -ForegroundColor DarkGray
                        $ch = Read-Host "Выбор"
                        if ($ch -match '^\d+$' -and [int]$ch -ge 1 -and [int]$ch -le $subs.Count) {
                            $sel = $subs[[int]$ch-1]; $rem = $num.ToString().Substring($sel.ToString().Length)
                            $target = "192.168.$sel.$([int]$rem)"
                        } elseif ([int]$ch -eq $subs.Count+1) { continue }
                        else { continue }
                    } elseif ($subs.Count -eq 1) {
                        $rem = $num.ToString().Substring($subs[0].ToString().Length)
                        $target = "192.168.$($subs[0]).$([int]$rem)"
                    } else { Write-Host "Подсеть не найдена!" -ForegroundColor Red; Start-Sleep 2; continue }
                } else { Write-Host "Ошибка ввода!" -ForegroundColor Red; Start-Sleep 2; continue }
            } else {
                $res = Get-HostnameAndIP -target $userInput
                if ($res.IP) { $target = $res.IP } else { Write-Host "Хост не найден!" -ForegroundColor Red; Start-Sleep 2; continue }
            }
            Start-Ping -Target $target -Continuous:$contPing
            continue
        }

        # === Преобразование в IP ===
        $ip = $null
        if ($userInput -match '^[\d\.]+$') {
            if ($userInput -match '^(\d{1,3})\.(\d{1,3})$') {
                $o3=[int]$matches[1]; $o4=[int]$matches[2]
                if ($o3 -gt 255 -or $o4 -gt 255) { Write-Host "Ошибка IP!" -ForegroundColor Red; Start-Sleep 2; continue }
                $ip = "192.168.$o3.$o4"
            } elseif ($userInput -match '^(\d+)$') {
                $num = [int]$userInput.Trim()
                if ($num -le 19) { Write-Host "Ошибка ввода" -ForegroundColor Red; Start-Sleep 2; continue }
                $subs = Get-PossibleSubnets -number $num -knownSubnets $knownSubnets
                if ($subs.Count -gt 1) {
                    Write-Host "`nНесколько подсетей:" -ForegroundColor Yellow
                    for($i=0;$i -lt $subs.Count;$i++) {
                        $rem = $num.ToString().Substring($subs[$i].ToString().Length)
                        Write-Host "$($i+1). 192.168.$($subs[$i]).$([int]$rem)" -ForegroundColor Gray
                    }
                    Write-Host "$($subs.Count+1). Отмена" -ForegroundColor DarkGray
                    $ch = Read-Host "Выбор"
                    if ($ch -match '^\d+$' -and [int]$ch -ge 1 -and [int]$ch -le $subs.Count) {
                        $sel = $subs[[int]$ch-1]; $rem = $num.ToString().Substring($sel.ToString().Length)
                        $ip = "192.168.$sel.$([int]$rem)"
                    } elseif ([int]$ch -eq $subs.Count+1) { continue }
                    else { continue }
                } elseif ($subs.Count -eq 1) {
                    $rem = $num.ToString().Substring($subs[0].ToString().Length)
                    $ip = "192.168.$($subs[0]).$([int]$rem)"
                } else { Write-Host "Подсеть не найдена!" -ForegroundColor Red; Start-Sleep 2; continue }
            } else { Write-Host "Ошибка ввода!" -ForegroundColor Red; Start-Sleep 2; continue }
        }

        # === Подключение по IP ===
        if ($ip) {
            $check = Test-TCPPortWithRetry -Target $ip
            if ($check.Success) {
                $oct3 = [int]($ip -split '\.')[2]
                if (-not ($knownSubnets -contains $oct3)) {
                    if ((Read-Host "Новая подсеть $oct3. Добавить? (y/n)") -match '^[yYдД]') {
                        Update-KnownSubnetsInScript -newSubnet $oct3
                        $knownSubnets += $oct3; $knownSubnets = $knownSubnets | Sort-Object
                        Write-Host "Подсеть добавлена." -ForegroundColor Cyan; Start-Sleep 3
                    }
                }
                $res = Get-HostnameAndIP -target $ip
                $hostN = $res.Hostname
                $disp = if ($hostN) { "$ip ($hostN)" } else { $ip }
                $cnt = Get-ConnectionCount -target $ip
                Write-Host "Успешно через порт $($check.Port)! Подключаемся..." -ForegroundColor Green
                Add-HistoryEntry -target $disp -isHostname $false -success $true -connectionCount $cnt
                Start-Process "msra.exe" -ArgumentList "/offerra $ip"
                Start-Sleep -Seconds 2
            } else {
                Add-HistoryEntry -target $ip -isHostname $false -success $false
                Start-Sleep -Seconds 2
            }
            continue
        }

        # === Подключение по имени хоста ===
        $check = Test-TCPPortWithRetry -Target $userInput
        if ($check.Success) {
            $res = Get-HostnameAndIP -target $userInput
            $ipRes = $res.IP
            $disp = if ($ipRes) { "$userInput ($ipRes)" } else { $userInput }
            $cnt = Get-ConnectionCount -target $userInput
            Write-Host "Успешно через порт $($check.Port)! Подключаемся..." -ForegroundColor Green
            Add-HistoryEntry -target $disp -isHostname $true -success $true -connectionCount $cnt
            Start-Process "msra.exe" -ArgumentList "/offerra $userInput"
            Start-Sleep -Seconds 2
        } else {
            Add-HistoryEntry -target $userInput -isHostname $true -success $false
            Start-Sleep -Seconds 2
        }
    } catch {
        Write-Host "Ошибка: $_" -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}