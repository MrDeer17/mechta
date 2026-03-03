# Voice Auto Dictation (Vosk + Sherpa-ONNX)

Утилита для автодиктовки с микрофона в выбранное окно Windows.

- Для классических моделей Vosk (`conf/graph`) используется `vosk`.
- Для `vosk-model-ru-0.54` (Zipformer ONNX) используется `sherpa-onnx`.

## Что умеет

- Слушает микрофон в фоне.
- Локально распознаёт речь.
- Отправляет текст в окно по `--target-title`.
- Поддерживает режим без смены фокуса (`--input-mode post`) для консоли (`pwsh.exe`/`cmd.exe`).
- Добавляет в конец каждой отправленной фразы `. ` (точка + пробел).
- Голосовые команды: пауза/продолжение/стоп, отправка Enter, очистка ввода.

## Установка

```powershell
python -m pip install -r requirements.txt
```

## Быстрый запуск

Классическая модель:

```powershell
.\start-voice-auto.cmd --target-title "ChatGPT" --quality best --auto-download-model
```

Zipformer `0.54` (уже скачана локально):

```powershell
.\start-voice-auto.cmd --target-title "pwsh.exe" --model-dir ".\models\vosk-model-ru-0.54" --input-mode post
```

## Голосовые команды

По умолчанию:

- Пауза: `pause dictation`, `пауза диктовка`
- Продолжить: `resume dictation`, `continue dictation`, `продолжай диктовка`
- Стоп: `stop dictation`, `стоп диктовка`
- Отправка (Enter): `send`, `submit`, `enter`, `отправка`, `отправить`, `энтер`
- Очистка ввода: `clear`, `очистить`, `очистка`, `очисти`

Поведение:

- В `post`-режиме Enter не жмётся автоматически при диктовке.
- Enter жмётся отдельной голосовой командой отправки.
- `очистить` удаляет надиктованный (не отправленный) текст через Backspace.

## Кастомные команды pause/resume/stop

Можно задать свои фразы:

```powershell
.\start-voice-auto.cmd `
  --target-title "pwsh.exe" `
  --model-dir ".\models\vosk-model-ru-0.54" `
  --input-mode post `
  --pause-command "пауза" `
  --resume-command "продолжай" `
  --stop-command "стоп"
```

Можно передавать несколько вариантов:

- повторяя аргумент (`--stop-command "стоп" --stop-command "хватит"`)
- или списком через `,` / `;` / `|` (`--resume-command "го,можно дальше|продолжай"`).

## Аргументы

- `--target-title` (обязательный): часть заголовка окна или имя процесса (`pwsh.exe`).
- `--model-dir`: путь к локальной модели.
- `--model-url`: URL модели для автоскачивания.
- `--quality {best,fast}`: профиль для классических моделей.
- `--auto-download-model`: скачать модель, если её нет.
- `--wake-word`: отправлять только фразы с wake word.
- `--min-chars`: минимум символов в распознанной фразе.
- `--no-enter`: не жать Enter после отправки текста (`type/paste` режимы).
- `--input-mode {type,paste,post}`:
  - `type`: ввод через `SendInput` (по умолчанию).
  - `paste`: через буфер обмена + `Ctrl+V`.
  - `post`: без смены фокуса через `PostMessage` (лучше для консоли).
- `--corrections-file`: JSON словарь замен.
- `--stop-command`, `--pause-command`, `--resume-command`: кастомные голосовые команды.

## Corrections JSON

Пример `corrections.ru.example.json`:

```json
{
  "чо": "че",
  "щас": "сейчас",
  "всм": "в смысле"
}
```

## Важно

- Windows может запросить доступ к микрофону для Python.
- Если цель не найдена по `--target-title`, отправка не выполняется.
- Для `post` лучше указывать `--target-title "pwsh.exe"` или `"cmd.exe"`.
- Модели Vosk добавлены в `.gitignore`, чтобы не коммитить большие файлы.
