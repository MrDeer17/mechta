# Voice Auto Dictation (Vosk)

Утилита для полностью автоматической диктовки с микрофона в выбранное окно.

Используется сторонний движок `Vosk`, без `Windows Speech Pack`.
По умолчанию включен профиль качества `best` (большая модель, точнее на разговорной речи и сленге).

## Что умеет

- Постоянно слушает микрофон.
- Распознаёт речь локально.
- Находит целевое окно по части заголовка.
- По умолчанию печатает текст напрямую (без буфера обмена) и нажимает `Enter`.
- Работает, даже когда консоль не в фокусе.

## Установка

```powershell
python -m pip install -r requirements.txt
```

## Быстрый запуск

```powershell
.\start-voice-auto.cmd --target-title "ChatGPT" --auto-download-model
```

Где `--target-title` это часть заголовка окна, куда отправлять текст.

Примеры:

```powershell
.\start-voice-auto.cmd --target-title "Telegram" --auto-download-model
.\start-voice-auto.cmd --target-title "Visual Studio Code" --auto-download-model --no-enter
```

Максимальное качество распознавания:

```powershell
.\start-voice-auto.cmd --target-title "C:\Program Files\PowerShell\7\pwsh.exe" --quality best --auto-download-model
```

Быстрый режим (меньше нагрузка, хуже точность):

```powershell
.\start-voice-auto.cmd --target-title "..." --quality fast --auto-download-model
```

## Голосовые команды

- `пауза диктовка`
- `продолжай диктовка`
- `стоп диктовка`

## Полезные параметры

- `--wake-word "мехта"`: отправлять только фразы, начинающиеся с wake word.
- `--no-enter`: вставлять текст без автоматической отправки.
- `--input-mode type|paste`: способ ввода (`type` по умолчанию, без буфера).
- `--quality best|fast`: профиль модели (`best` лучше понимает разговорную речь).
- `--min-chars 4`: фильтрация слишком коротких фраз.
- `--model-dir <path>`: путь к локальной модели.
- `--model-url <url>`: откуда скачивать модель.
- `--corrections-file corrections.ru.example.json`: авто-исправления ослышек/сленга.

## Сленг и авто-исправления

Можно подложить JSON-файл словаря замен:

```json
{
  "чо": "че",
  "щас": "сейчас",
  "всм": "в смысле"
}
```

Запуск:

```powershell
.\start-voice-auto.cmd --target-title "..." --quality best --corrections-file .\corrections.ru.example.json
```

## Важно

- Windows может запросить доступ к микрофону для Python.
- Если целевое окно закрыто или заголовок не совпадает, отправка не сработает.
- Утилита может переключать фокус на целевое окно перед вставкой текста.
- `best` скачивает большую модель (~1.9 GB), первый запуск может занять время.
