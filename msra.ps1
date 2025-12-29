# =============================================
# Умное подключение MSRA v2.1
# Поддержка: trueconf, pacs, ping, история и т.д.
# =============================================

$scriptVersion = "2.1"

# Настройки истории
$historyFile = Join-Path $PSScriptRoot "msra_history.log"
$maxHistoryEntries = 1000

# Список известных подсетей
$knownSubnets = @(2, 3, 11, 13, 14, 15, 17, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 31, 32, 33, 35, 36, 37, 38, 48, 56, 57, 60, 61, 63, 91, 92, 96, 97, 98, 102, 103, 104, 105, 106, 107, 110, 111, 112, 113, 114, 122, 123, 124, 125, 126, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 214, 215)

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
    Write-Host "• sigur     - Пароль от СКУДа" -ForegroundColor Gray
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
                $color = if ($_ -match "Успешно") { "Green" } else { "Red" }
                Write-Host " $_" -ForegroundColor $color
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
            Write-Host "Удалено $removedCount неудачных записей из истории." -ForegroundColor Green
        }
        Start-Sleep -Seconds 2
    } else {
        Write-Host "Файл истории не найден." -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

function Remove-DuplicateHistoryEntries {
    if (Test-Path $historyFile) {
        $history = Get-Content $historyFile
        $filteredHistory = @()
        $prevTarget = ""
        $prevStatus = ""
        $removedCount = 0

        foreach ($line in $history) {
            $fields = $line.Trim() -split '\s*\|\s*'
            if ($fields.Count -eq 4) {
                $target = $fields[2]
                $status = $fields[3]
                if ($target -ne $prevTarget -or $status -ne $prevStatus) {
                    $filteredHistory += $line
                    } else {
                        $removedCount++

                }
                # Запоминаем только для сравнения следующих подряд
                $prevTarget = $target
                $prevStatus = $status
            } else {
                # Строки не по формату не удаляем и сбрасываем сравнение
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
        $history = Get-Content $historyFile
        $prevTarget = ""
        $prevStatus = ""
        foreach ($line in $history) {
            $fields = $line.Trim() -split '\s*\|\s*'
            if ($fields.Count -eq 4) {
                $target = $fields[2]
                $status = $fields[3]
                if ($target -eq $prevTarget -and $status -eq $prevStatus) {
                    return $true
                }
                $prevTarget = $target
                $prevStatus = $status
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
                $color = if ($_ -match "Успешно") { "Green" } else { "Red" }
                Write-Host " $_" -ForegroundColor $color
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
            if ($hasFailed) {
                Write-Host "Del     - Удалить все неудачные подключения"
            }
            if ($hasDuplicates) {
                Write-Host "Deldbl  - Удалить подряд идущие дубликаты"
            }
            if ($hasFailed -and $hasDuplicates) {
                Write-Host "Delall  - Удалить и неудачные, и дубликаты"
            }

            Write-Host "Enter   - Вернуться в меню"

            $choice = Read-Host "`nВыбор"
            switch ($choice.Trim().ToLower()) {
                { $_ -in "del", "вуд" } {
                    Remove-FailedHistoryEntries
                }
                "deldbl" {
                    Remove-DuplicateHistoryEntries
                }
                { $_ -in "delall", "вудфдд" } {
                    # Удаляем всё: сначала неудачи, потом дубликаты
                    if (Test-Path $historyFile) {
                        # Удаляем неудачи
                        $historyBefore = Get-Content $historyFile
                        $failedCount = ($historyBefore -match "Нет пинга").Count
                        if ($failedCount -gt 0) {
                            Remove-FailedHistoryEntries
                        }

                        # Теперь удаляем дубликаты (уже из обновлённого файла)
                        if (Test-ConsecutiveDuplicates) {
                            Remove-DuplicateHistoryEntries
                        }

                        # Считаем итоговое удаление
                        $historyAfter = Get-Content $historyFile
                        $removedTotal = $historyBefore.Count - $historyAfter.Count
                        if ($removedTotal -gt 0) {
                            Write-Host "Всего удалено записей: $removedTotal" -ForegroundColor Cyan
                        } else {
                            Write-Host "Нечего удалять." -ForegroundColor Yellow
                        }
                        Start-Sleep -Seconds 2
                    } else {
                        Write-Host "Файл истории не найден." -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
                default {
                    # Возвращаемся в меню
                }
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
    param(
        [string]$target,
        [bool]$isHostname,
        [bool]$success
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $type = if ($isHostname) { "Хост" } else { "IP" }
    $status = if ($success) { "Успешно" } else { "Нет пинга" }
    $entry = "$timestamp | $type | $target | $status"
    $currentHistory = @()
    if (Test-Path $historyFile) {
        $currentHistory = Get-Content $historyFile
    }
    $newHistory = ($currentHistory -join "`n") + "`n" + $entry
    ($newHistory -split "`n" | Where-Object { $_ -ne "" } | Select-Object -Last $maxHistoryEntries) -join "`n" | Out-File $historyFile -Encoding utf8
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
        Write-Host "Автоматическое обновление возможно только при запуске скрипта из файла!" -ForegroundColor Red
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

        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $connect = $tcp.BeginConnect($Target, $port, $null, $null)
            $wait = $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)

            if ($wait) {
                $tcp.EndConnect($connect)
                $tcp.Close()
                Write-Host " Успешно!" -ForegroundColor Green
                return @{
                    Success = $true
                    Port    = $port
                }
            } else {
                $tcp.Close()
                Write-Host " Закрыт/недоступен" -ForegroundColor DarkGray
            }
        } catch {
            Write-Host " Ошибка" -ForegroundColor Red
        }
    }

    Write-Host "  → Устройство недоступно по TCP (порты $($Ports -join ', '))." -ForegroundColor Red
    return @{
        Success = $false
        Port    = $null
    }
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
        Write-Host "║     Умное подключение MSRA v$scriptVersion      ║" -ForegroundColor Cyan
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
        $input = Read-Host "Введите данные (Enter для выхода)"
        if ([string]::IsNullOrEmpty($input)) { break }

        # === Обработка специальных команд (ПОРЯДОК ВАЖЕН!) ===

        if ($input -eq "help" -or $input -eq "рудз") {
            Show-Help
            continue
        }

        if ($input -eq "his" -or $input -eq "ршы") {
            Clear-Host
            Write-Host "── Полная история подключений ──" -ForegroundColor Cyan
            Show-FullHistory
            continue
        }

        if ($input -eq "clear" -or $input -eq "сдефк") {
            Clear-HistoryFile
            continue
        }

        if ($input -ieq "trueconf") {
    $urlTrueConf = "http://ssrv.lan.smclinic.ru/admin/general/info/"
    $loginTrueConf = "trueconftech"
    $passwordTrueConf = "Beif8uwi"

    # 1. Сразу открываем страницу
    Start-Process $urlTrueConf
    Write-Host "`nОткрывается страница входа TrueConf..." -ForegroundColor Green

    # 2. Копируем логин
    Set-Clipboard -Value $loginTrueConf
    Write-Host "Логин '$loginTrueConf' скопирован в буфер обмена." -ForegroundColor Green
    Write-Host ""

    # 3. Спрашиваем про пароль (без таймаута)
    $copyPassword = Read-Host "Скопировать пароль в буфер обмена? (y/n)"
    if ($copyPassword -match '^[yYдД]') {
        Set-Clipboard -Value $passwordTrueConf
        Write-Host "`nПароль скопирован в буфер обмена." -ForegroundColor Green
    } else {
        Write-Host "`nПароль не был скопирован." -ForegroundColor Yellow
    }

    Write-Host "`nГотово. Возвращаемся в меню..." -ForegroundColor Green
    Start-Sleep -Seconds 2
    continue
}

if ($input -ieq "update") {
    $currentScript = $PSCommandPath
    if ([string]::IsNullOrEmpty($currentScript)) {
        Write-Host "Обновление невозможно: скрипт запущен не из файла." -ForegroundColor Red
        Read-Host "Нажмите Enter"
        continue
    }

    $repoUrl = "https://raw.githubusercontent.com/stz1397-lab/msra-helper/main/msra.ps1"
    $tempFile = [System.IO.Path]::GetTempFileName()

    try {
        Write-Host "Загрузка новой версии..." -ForegroundColor Cyan

        # Определяем, используется ли прокси для этого URL
        $systemProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $proxyUri = $systemProxy.GetProxy($repoUrl)

        if ($proxyUri -eq $repoUrl) {
            # Прямое подключение (без прокси)
            Invoke-WebRequest -Uri $repoUrl -OutFile $tempFile -UseBasicParsing -TimeoutSec 15
        } else {
            # Используем прокси с текущими учётными данными Windows
            Write-Host "Используется прокси: $proxyUri" -ForegroundColor DarkGray
            Invoke-WebRequest -Uri $repoUrl -OutFile $tempFile -UseBasicParsing -TimeoutSec 15 -Proxy $proxyUri -ProxyUseDefaultCredentials
        }

        # Сравниваем содержимое
        $currentContent = Get-Content $currentScript -Raw
        $newContent = Get-Content $tempFile -Raw

        if ($currentContent -eq $newContent) {
            Write-Host "У вас уже установлена последняя версия." -ForegroundColor Green
        } else {
            $backup = "$currentScript.bak"
            Copy-Item $currentScript $backup -Force
            Write-Host "Создана резервная копия: $backup" -ForegroundColor Yellow

            Copy-Item $tempFile $currentScript -Force
            Write-Host "Скрипт обновлён! Перезапуск..." -ForegroundColor Green
            Start-Sleep -Seconds 1
            & $currentScript
            exit
        }
    } catch {
        Write-Host "Ошибка при обновлении: $_" -ForegroundColor Red
    } finally {
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
    }

    Read-Host "Нажмите Enter, чтобы вернуться в меню"
    continue
}

        if ($input -ieq "glpi") {
        $urlGLPI = "https://glpi.smclinic.ru/front/ticket.php"
        Start-Process $urlGLPI
        Write-Host "GLPI запущен" -ForegroundColor Green
        Start-Sleep -Seconds 2
        continue
        }

        if ($input -ieq "restart") {
    # Проверяем, есть ли активные сессии
    $sessions = (Get-CimInstance -ClassName Win32_ComputerSystem).NumberOfUsers
    if ($sessions -gt 1) {
        Write-Warning "В системе есть активные пользователи. Перезагрузка может быть отклонена."
        $confirm = Read-Host "Всё равно продолжить? (y/n)"
        if ($confirm -notmatch '^(y|yes|д|да)$') {
            Write-Host "Отмена перезагрузки." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            continue
        }
    }

    try {
        Write-Host "Инициируется перезагрузка..." -ForegroundColor Magenta
        Restart-Computer -Force -ErrorAction Stop
        Write-Host "Перезагрузка выполнена." -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Host "Ошибка при перезагрузке: $($_.Exception.Message)" -ForegroundColor Red
        Start-Sleep -Seconds 3
    }
    continue
}

        if ($input -ieq "exit") {
        exit
        }

        if ($input -ieq "scanpass") {
        $scanPass = "53807553QaZ"
        Write-Host ""
        Write-Host $scanPass
        $scanPass | Set-Clipboard
        Write-Host ""
        Write-Host "Пароль в буфере обмена"
        Read-Host "Нажмите Enter, чтобы вернуться в главное меню"
        continue
        }

        if ($input -ieq "sigur") {
        $sigurpass = "rt54de1z"
        Write-Host ""
        Write-Host $sigurpass
        $sigurpass | Set-Clipboard
        Write-Host ""
        Write-Host "Пароль в буфере обмена"
        Read-Host "Нажмите Enter, чтобы вернуться в главное меню"
        continue
        }

        if ($input -ieq "distr") {
        Invoke-Item "\\fileserver\distr$"
        continue
}

        
if ($input -ieq "pacs") {
    $urlPacs = "http://pacs-2.lan.smclinic.ru/pacs/login.php"
    $urlInstruction = "https://conf.smclinic.ru/spaces/WTS/pages/121275145/%D0%90%D0%B4%D0%BC%D0%B8%D0%BD%D0%BA%D0%B0+PACS"

    Start-Process $urlPacs
    Write-Host "`nОткрывается страница входа PACS..." -ForegroundColor Green

    $openInstruction = Read-Host "Открыть инструкцию по настройке PACS? (y/n)"
    if ($openInstruction -match '^[yYдД]') {
        Start-Process $urlInstruction
        Write-Host "`nИнструкция по настройке PACS открыта в браузере." -ForegroundColor Cyan
    } else {
        Write-Host "`nИнструкция не будет открыта." -ForegroundColor Yellow
    }

    Write-Host "`nГотово. Возвращаемся в меню..." -ForegroundColor Green
    Start-Sleep -Seconds 2
    continue
}

        # Обработка команды ping
        $isPingCommand = $false
        $continuousPing = $false
        $ip = $null
        if ($input -match '^ping\s+(.+?)(\s+-t)?$') {
            $isPingCommand = $true
            $continuousPing = $matches[2] -ne $null
            $input = $matches[1].Trim()
        }

        if ($isPingCommand) {
            $target = $null

            if ($input -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                $target = $input
            }
            elseif ($input -match '^[\d\.]+$') {
                if ($input -match '^(\d{1,3})\.(\d{1,3})$') {
                    $octet3 = [int]$matches[1]
                    $octet4 = [int]$matches[2]
                    if ($octet3 -gt 255 -or $octet4 -gt 255) {
                        Write-Host "Ошибка: неверный IP!" -ForegroundColor Red
                        Start-Sleep -Seconds 2
                        continue
                    }
                    $target = "192.168.$octet3.$octet4"
                }
                elseif ($input -match '^(\d+)$') {
                    $num = [int]$input.Trim()
                    if ($num -ge 0 -and $num -le 19) {
                        Write-Host "Ошибка ввода" -ForegroundColor Red
                        Start-Sleep -Seconds 2
                        continue
                    }
                    $possibleSubnets = Get-PossibleSubnets -number $num -knownSubnets $knownSubnets
                    if ($possibleSubnets.Count -gt 1) {
                        Write-Host "`nОбнаружено несколько подсетей:" -ForegroundColor Yellow
                        for ($i = 0; $i -lt $possibleSubnets.Count; $i++) {
                            $subnet = $possibleSubnets[$i]
                            $remainingDigits = $num.ToString().Substring($subnet.ToString().Length)
                            $exampleIP = "192.168.$subnet.$([int]$remainingDigits)"
                            Write-Host "$($i+1). $exampleIP" -ForegroundColor Gray
                        }
                        Write-Host "$($possibleSubnets.Count+1). Отмена" -ForegroundColor DarkGray
                        $choice = Read-Host "Выберите вариант (1-$($possibleSubnets.Count+1))"
                        if ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $possibleSubnets.Count+1) {
                            if ([int]$choice -eq $possibleSubnets.Count+1) { continue }
                            $selectedSubnet = $possibleSubnets[[int]$choice - 1]
                            $remainingDigits = $num.ToString().Substring($selectedSubnet.ToString().Length)
                            $target = "192.168.$selectedSubnet.$([int]$remainingDigits)"
                        } else { continue }
                    } elseif ($possibleSubnets.Count -eq 1) {
                        $subnet = $possibleSubnets[0]
                        $remainingDigits = $num.ToString().Substring($subnet.ToString().Length)
                        $target = "192.168.$subnet.$([int]$remainingDigits)"
                    } else {
                        Write-Host "Не найдено подходящих подсетей для '$num'!" -ForegroundColor Red
                        Start-Sleep -Seconds 2
                        continue
                    }
                } else {
                    Write-Host "Ошибка: неверный ввод!" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
            }
            else {
                $resolved = Get-HostnameAndIP -target $input
                if ($resolved.IP) {
                    $target = $resolved.IP
                } else {
                    Write-Host "Хост не найден! '$input'" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
            }

            Start-Ping -Target $target -Continuous:$continuousPing
            continue
        }

        # --- Основная логика: преобразование в IP ---
        $ip = $null
        if ($input -match '^[\d\.]+$') {
            if ($input -match '^(\d{1,3})\.(\d{1,3})$') {
                $octet3 = [int]$matches[1]
                $octet4 = [int]$matches[2]
                if ($octet3 -gt 255 -or $octet4 -gt 255) {
                    Write-Host "Ошибка: неверный IP!" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
                $ip = "192.168.$octet3.$octet4"
            }
            elseif ($input -match '^(\d+)$') {
                $num = [int]$input.Trim()
                if ($num -ge 0 -and $num -le 19) {
                    Write-Host "Ошибка ввода" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
                $possibleSubnets = Get-PossibleSubnets -number $num -knownSubnets $knownSubnets
                if ($possibleSubnets.Count -gt 1) {
                    Write-Host "`nОбнаружено несколько подсетей:" -ForegroundColor Yellow
                    for ($i = 0; $i -lt $possibleSubnets.Count; $i++) {
                        $subnet = $possibleSubnets[$i]
                        $remainingDigits = $num.ToString().Substring($subnet.ToString().Length)
                        $exampleIP = "192.168.$subnet.$([int]$remainingDigits)"
                        Write-Host "$($i+1). $exampleIP" -ForegroundColor Gray
                    }
                    Write-Host "$($possibleSubnets.Count+1). Отмена" -ForegroundColor DarkGray
                    $choice = Read-Host "Выберите вариант (1-$($possibleSubnets.Count+1))"
                    if ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $possibleSubnets.Count+1) {
                        if ([int]$choice -eq $possibleSubnets.Count+1) { continue }
                        $selectedSubnet = $possibleSubnets[[int]$choice - 1]
                        $remainingDigits = $num.ToString().Substring($selectedSubnet.ToString().Length)
                        $ip = "192.168.$selectedSubnet.$([int]$remainingDigits)"
                    } else { continue }
                } elseif ($possibleSubnets.Count -eq 1) {
                    $subnet = $possibleSubnets[0]
                    $remainingDigits = $num.ToString().Substring($subnet.ToString().Length)
                    $ip = "192.168.$subnet.$([int]$remainingDigits)"
                } else {
                    Write-Host "Не найдено подходящих подсетей для '$num'!" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
            } else {
                Write-Host "Ошибка: неверный ввод!" -ForegroundColor Red
                Start-Sleep -Seconds 2
                continue
            }
        }

# --- Подключение по IP ---
if ($ip) {
    $checkResult = Test-TCPPortWithFeedback -Target $ip
    $pingResult = $checkResult.Success
    if ($pingResult) {
        $octet3 = [int]($ip -split '\.')[2]
        if (-not ($knownSubnets -contains $octet3)) {
            $answer = Read-Host "Обнаружена новая подсеть $octet3. Добавить её в список известных? (y/n)"
            if ($answer -match '^[yYдД]') {
                Update-KnownSubnetsInScript -newSubnet $octet3
                $knownSubnets += $octet3
                $knownSubnets = $knownSubnets | Sort-Object
                Write-Host "Подсеть $octet3 добавлена." -ForegroundColor Cyan
                Start-Sleep -Seconds 3
            }
        }
        $resolved = Get-HostnameAndIP -target $ip
        $hostname = $resolved.Hostname
        $displayTarget = if ($hostname) { "$ip ($hostname)" } else { $ip }
        $viaPort = $checkResult.Port
        Write-Host "Успешно через порт $viaPort! Подключаемся..." -ForegroundColor Green
        Add-HistoryEntry -target $displayTarget -isHostname $false -success $true
        Start-Process "msra.exe" -ArgumentList "/offerra $ip"
        Start-Sleep -Seconds 2
    } else {
        Add-HistoryEntry -target $ip -isHostname $false -success $false
        Start-Sleep -Seconds 2
    }
    continue
}

# --- Подключение по имени хоста ---
$checkResult = Test-TCPPortWithFeedback -Target $input
$pingResult = $checkResult.Success
if ($pingResult) {
    $resolved = Get-HostnameAndIP -target $input
    $ip = $resolved.IP
    $hostname = $resolved.Hostname
    $displayTarget = if ($ip) { "$input ($ip)" } else { $input }
    $viaPort = $checkResult.Port
    Write-Host "Успешно через порт $viaPort! Подключаемся..." -ForegroundColor Green
    Add-HistoryEntry -target $displayTarget -isHostname $true -success $true
    Start-Process "msra.exe" -ArgumentList "/offerra $input"
    Start-Sleep -Seconds 2
} else {
            Add-HistoryEntry -target $input -isHostname $true -success $false
        Start-Sleep -Seconds 2
    }
} catch {
    Write-Host "Ошибка: $_" -ForegroundColor Red
    Start-Sleep -Seconds 2
}
}