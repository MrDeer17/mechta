param(
    [string]$Language = "ru-RU",
    [double]$MinConfidence = 0.55,
    [switch]$AutoEnter = $true,
    [string]$WakeWord = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Speech
Add-Type -AssemblyName System.Windows.Forms

function New-Recognizer {
    param(
        [string]$PreferredCulture
    )

    $installed = [System.Speech.Recognition.SpeechRecognitionEngine]::InstalledRecognizers()
    if ($installed.Count -gt 0) {
        $match = $installed | Where-Object { $_.Culture.Name -ieq $PreferredCulture } | Select-Object -First 1
        if ($null -ne $match) {
            return [pscustomobject]@{
                Engine = [System.Speech.Recognition.SpeechRecognitionEngine]::new($match.Culture)
                CultureName = $match.Culture.Name
            }
        }

        Write-Warning "Культура '$PreferredCulture' не найдена. Будет использована '$($installed[0].Culture.Name)'."
        return [pscustomobject]@{
            Engine = [System.Speech.Recognition.SpeechRecognitionEngine]::new($installed[0].Culture)
            CultureName = $installed[0].Culture.Name
        }
    }

    Write-Warning "Список установленных speech-движков пуст. Пробую движок по умолчанию."
    try {
        $engine = [System.Speech.Recognition.SpeechRecognitionEngine]::new()
        $cultureName = "default"
        if ($engine.RecognizerInfo -and $engine.RecognizerInfo.Culture) {
            $cultureName = $engine.RecognizerInfo.Culture.Name
        }
        return [pscustomobject]@{
            Engine = $engine
            CultureName = $cultureName
        }
    }
    catch {
        throw "Не удалось инициализировать распознавание речи. Установи Windows Speech language pack (например, ru-RU). Ошибка: $($_.Exception.Message)"
    }
}

function Send-TextToActiveWindow {
    param(
        [string]$Text,
        [bool]$PressEnter
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    Set-Clipboard -Value $Text
    Start-Sleep -Milliseconds 25
    [System.Windows.Forms.SendKeys]::SendWait("^v")
    if ($PressEnter) {
        Start-Sleep -Milliseconds 25
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    }
}

$recognizerData = New-Recognizer -PreferredCulture $Language
$recognizer = $recognizerData.Engine
try {
    $recognizer.SetInputToDefaultAudioDevice()
}
catch {
    throw "Не удалось подключиться к микрофону по умолчанию. Проверь, что микрофон подключён и выбран устройством ввода в Windows. Ошибка: $($_.Exception.Message)"
}
$recognizer.LoadGrammar([System.Speech.Recognition.DictationGrammar]::new())

$isPaused = $false
$wakeWordNorm = $WakeWord.Trim().ToLowerInvariant()
$autoEnterEnabled = $AutoEnter.IsPresent

Write-Host "Активная культура распознавания: $($recognizerData.CultureName)"
Write-Host "Порог confidence: $MinConfidence"
Write-Host "Режим отправки Enter: $autoEnterEnabled"
Write-Host "Голосовые команды: 'пауза диктовка', 'продолжай диктовка', 'стоп диктовка'"
if ($wakeWordNorm) {
    Write-Host "Wake word: '$WakeWord' (будут отправляться только фразы, начинающиеся с него)"
}
Write-Host "Важно: оставь фокус в поле ввода (чат/редактор), куда нужно печатать."

try {
    while ($true) {
        $result = $recognizer.Recognize()
        if ($null -eq $result) {
            continue
        }

        $text = $result.Text.Trim()
        $confidence = $result.Confidence

        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }
        if ($confidence -lt $MinConfidence) {
            continue
        }

        $normalized = $text.ToLowerInvariant()

        if ($normalized -eq "стоп диктовка") {
            Write-Host "[voice] остановка"
            break
        }
        if ($normalized -eq "пауза диктовка") {
            Write-Host "[voice] пауза"
            $isPaused = $true
            continue
        }
        if ($normalized -eq "продолжай диктовка") {
            Write-Host "[voice] продолжение"
            $isPaused = $false
            continue
        }
        if ($isPaused) {
            continue
        }

        $outText = $text
        if ($wakeWordNorm) {
            if (-not $normalized.StartsWith($wakeWordNorm)) {
                continue
            }
            $outText = $text.Substring($wakeWordNorm.Length).TrimStart()
            if ([string]::IsNullOrWhiteSpace($outText)) {
                continue
            }
        }

        Write-Host ("[voice {0:n2}] {1}" -f $confidence, $outText)
        Send-TextToActiveWindow -Text $outText -PressEnter $autoEnterEnabled
    }
}
finally {
    $recognizer.Dispose()
}



