#!/usr/bin/env python
"""
Continuous dictation to a target window using Vosk (third-party STT engine).

Features:
- Works without Windows Speech language packs.
- Runs in background; console window does not need to be focused.
- Sends recognized phrases to a target window selected by title substring.
"""

from __future__ import annotations

import argparse
import ctypes
import json
import queue
import re
import sys
import tempfile
import threading
import time
import zipfile
from ctypes import wintypes
from pathlib import Path
from typing import Iterable, Optional

import pyperclip
import requests
import sounddevice as sd
from vosk import KaldiRecognizer, Model


MODEL_PROFILES = {
    "fast": {
        "name": "vosk-model-small-ru-0.22",
        "url": "https://alphacephei.com/vosk/models/vosk-model-small-ru-0.22.zip",
    },
    "best": {
        "name": "vosk-model-ru-0.42",
        "url": "https://alphacephei.com/vosk/models/vosk-model-ru-0.42.zip",
    },
}
DEFAULT_MODEL_PROFILE = "best"
DEFAULT_SAMPLE_RATE = 16000
DEFAULT_SILENCE_RMS_THRESHOLD = 500.0
DEFAULT_END_OF_UTTERANCE_SILENCE_SEC = 0.8
DEFAULT_MIN_UTTERANCE_SEC = 0.35
DEFAULT_MAX_UTTERANCE_SEC = 20.0
DEFAULT_FOCUS_COOLDOWN_SEC = 0.8
DEFAULT_DUPLICATE_SUPPRESS_SEC = 1.5

# Keyboard constants
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
VK_CONTROL = 0x11
VK_RETURN = 0x0D
VK_V = 0x56
VK_BACK = 0x08
ULONG_PTR = wintypes.WPARAM

# Window constants
SW_RESTORE = 9
SW_SHOW = 5
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_CHAR = 0x0102
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)


def configure_stdio() -> None:
    """Avoid Unicode crashes in legacy Windows console encodings."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]


class _INPUTUNION(ctypes.Union):
    _fields_ = [
        ("mi", MOUSEINPUT),
        ("ki", KEYBDINPUT),
        ("hi", HARDWAREINPUT),
    ]


class INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [
        ("type", wintypes.DWORD),
        ("u", _INPUTUNION),
    ]


class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("hwndActive", wintypes.HWND),
        ("hwndFocus", wintypes.HWND),
        ("hwndCapture", wintypes.HWND),
        ("hwndMenuOwner", wintypes.HWND),
        ("hwndMoveSize", wintypes.HWND),
        ("hwndCaret", wintypes.HWND),
        ("rcCaret", wintypes.RECT),
    ]


EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def check_bool(result, func, args):
    if not result:
        err = ctypes.get_last_error()
        raise ctypes.WinError(err)
    return args


user32.EnumWindows.argtypes = [EnumWindowsProc, wintypes.LPARAM]
user32.EnumWindows.restype = wintypes.BOOL

user32.IsWindowVisible.argtypes = [wintypes.HWND]
user32.IsWindowVisible.restype = wintypes.BOOL

user32.IsWindow.argtypes = [wintypes.HWND]
user32.IsWindow.restype = wintypes.BOOL

user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = ctypes.c_int

user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int

user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND

user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.SetForegroundWindow.restype = wintypes.BOOL

user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
user32.ShowWindow.restype = wintypes.BOOL

user32.IsIconic.argtypes = [wintypes.HWND]
user32.IsIconic.restype = wintypes.BOOL

user32.BringWindowToTop.argtypes = [wintypes.HWND]
user32.BringWindowToTop.restype = wintypes.BOOL

user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD

user32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
user32.AttachThreadInput.restype = wintypes.BOOL

user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
user32.SendInput.restype = wintypes.UINT

user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostMessageW.restype = wintypes.BOOL

user32.GetGUIThreadInfo.argtypes = [wintypes.DWORD, ctypes.POINTER(GUITHREADINFO)]
user32.GetGUIThreadInfo.restype = wintypes.BOOL

kernel32.GetCurrentThreadId.argtypes = []
kernel32.GetCurrentThreadId.restype = wintypes.DWORD

kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE

kernel32.QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.LPWSTR,
    ctypes.POINTER(wintypes.DWORD),
]
kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL

kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CloseHandle.restype = wintypes.BOOL


def _download_model(model_url: str, models_root: Path) -> Path:
    models_root.mkdir(parents=True, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix="vosk_model_", suffix=".zip")
    try:
        # mkstemp returns an open file descriptor; close it before reopening by path.
        import os

        os.close(fd)
    except OSError:
        pass
    tmp_zip = Path(tmp_path)

    print(f"[init] Downloading model: {model_url}")
    with requests.get(model_url, stream=True, timeout=60) as resp:
        resp.raise_for_status()
        with tmp_zip.open("wb") as fh:
            for chunk in resp.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    fh.write(chunk)

    print(f"[init] Extracting model to: {models_root}")
    with zipfile.ZipFile(tmp_zip, "r") as zf:
        zf.extractall(models_root)
        roots = sorted(
            {
                name.split("/", 1)[0]
                for name in zf.namelist()
                if name and not name.startswith("__MACOSX/")
            }
        )
    tmp_zip.unlink(missing_ok=True)

    if not roots:
        raise RuntimeError("Downloaded archive does not contain model files.")
    model_dir = models_root / roots[0]
    return model_dir


def ensure_model(model_dir: Path, model_url: str, auto_download: bool) -> Path:
    if model_dir.exists():
        return model_dir
    if not auto_download:
        raise FileNotFoundError(
            f"Model directory not found: {model_dir}\n"
            "Pass --auto-download-model or provide an existing --model-dir."
        )
    return _download_model(model_url, model_dir.parent)


def _is_classic_vosk_model(model_dir: Path) -> bool:
    return (model_dir / "conf").exists() and (model_dir / "graph").exists()


def _get_zipformer_onnx_paths(
    model_dir: Path,
) -> Optional[tuple[Path, Path, Path, Path]]:
    am_onnx_dir = model_dir / "am-onnx"
    tokens = model_dir / "lang" / "tokens.txt"
    if not am_onnx_dir.exists() or not tokens.is_file():
        return None

    int8_paths = (
        am_onnx_dir / "encoder.int8.onnx",
        am_onnx_dir / "decoder.int8.onnx",
        am_onnx_dir / "joiner.int8.onnx",
    )
    if all(p.is_file() for p in int8_paths):
        return int8_paths[0], int8_paths[1], int8_paths[2], tokens

    fp32_paths = (
        am_onnx_dir / "encoder.onnx",
        am_onnx_dir / "decoder.onnx",
        am_onnx_dir / "joiner.onnx",
    )
    if all(p.is_file() for p in fp32_paths):
        return fp32_paths[0], fp32_paths[1], fp32_paths[2], tokens

    return None


def detect_model_backend(model_dir: Path) -> str:
    if _is_classic_vosk_model(model_dir):
        return "vosk"
    if _get_zipformer_onnx_paths(model_dir):
        return "sherpa-onnx"

    if (model_dir / "lang").exists() and (model_dir / "lm").exists() and (
        (model_dir / "am").exists() or (model_dir / "am-onnx").exists()
    ):
        raise RuntimeError(
            "Unsupported model format in "
            f"{model_dir}. This directory looks like Zipformer (am/lang/lm), "
            "but required ONNX files were not found in am-onnx/. "
            "Expected either encoder/decoder/joiner.int8.onnx or "
            "encoder/decoder/joiner.onnx plus lang/tokens.txt."
        )

    raise RuntimeError(
        "Invalid model directory: "
        f"{model_dir}. Expected either classic Vosk model folders "
        "('conf' and 'graph') or Zipformer ONNX files in 'am-onnx' and "
        "'lang/tokens.txt'."
    )


def build_sherpa_offline_recognizer(model_dir: Path):
    try:
        import sherpa_onnx  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Zipformer model requires sherpa-onnx. Install it with: "
            "python -m pip install sherpa-onnx"
        ) from exc

    onnx_paths = _get_zipformer_onnx_paths(model_dir)
    if onnx_paths is None:
        raise RuntimeError(
            f"Zipformer ONNX files are missing in: {model_dir / 'am-onnx'}"
        )

    encoder, decoder, joiner, tokens = onnx_paths
    print(f"[init] Tokens: {tokens}")
    print(f"[init] Encoder: {encoder.name}")
    print(f"[init] Decoder: {decoder.name}")
    print(f"[init] Joiner: {joiner.name}")

    return sherpa_onnx.OfflineRecognizer.from_transducer(
        tokens=str(tokens),
        encoder=str(encoder),
        decoder=str(decoder),
        joiner=str(joiner),
        num_threads=2,
        sample_rate=DEFAULT_SAMPLE_RATE,
        dither=3e-5,
        decoding_method="modified_beam_search",
        max_active_paths=10,
        provider="cpu",
    )


def iter_visible_windows() -> Iterable[tuple[int, str]]:
    windows: list[tuple[int, str]] = []

    @EnumWindowsProc
    def _enum(hwnd: int, _lparam: int) -> bool:
        if not user32.IsWindowVisible(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, len(buf))
        title = buf.value.strip()
        if title:
            windows.append((hwnd, title))
        return True

    user32.EnumWindows(_enum, 0)
    return windows


def find_window_by_title_substring(needle: str) -> Optional[int]:
    needle_norm = needle.lower()
    process_hint = needle_norm if needle_norm.endswith(".exe") and " " not in needle_norm else ""
    title_match: Optional[int] = None

    for hwnd, title in iter_visible_windows():
        if process_hint:
            proc_name = get_window_process_name(hwnd)
            if proc_name and proc_name == process_hint:
                return hwnd

        if title_match is None and needle_norm in title.lower():
            title_match = hwnd
    return title_match


def get_window_process_name(hwnd: int) -> Optional[str]:
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        return None

    hproc = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
    if not hproc:
        return None
    try:
        buf_size = wintypes.DWORD(1024)
        buf = ctypes.create_unicode_buffer(buf_size.value)
        if not kernel32.QueryFullProcessImageNameW(hproc, 0, buf, ctypes.byref(buf_size)):
            return None
        image_path = buf.value
        if not image_path:
            return None
        return Path(image_path).name.lower()
    finally:
        kernel32.CloseHandle(hproc)


def is_hwnd_alive(hwnd: int) -> bool:
    return bool(user32.IsWindow(hwnd))


class TargetWindowResolver:
    def __init__(self, target_title: str) -> None:
        self.target_title = target_title
        self.cached_hwnd: Optional[int] = None

    def get_hwnd(self) -> int:
        if self.cached_hwnd is not None and is_hwnd_alive(self.cached_hwnd):
            return self.cached_hwnd

        hwnd = find_window_by_title_substring(self.target_title)
        if hwnd is None:
            raise RuntimeError(
                f"Target window not found by title substring: '{self.target_title}'"
            )
        self.cached_hwnd = hwnd
        return hwnd

    def try_get_hwnd(self) -> Optional[int]:
        if self.cached_hwnd is not None and is_hwnd_alive(self.cached_hwnd):
            return self.cached_hwnd
        hwnd = find_window_by_title_substring(self.target_title)
        if hwnd is not None:
            self.cached_hwnd = hwnd
            return hwnd
        return None


def _send_key(vk: int, key_up: bool = False) -> None:
    flags = KEYEVENTF_KEYUP if key_up else 0
    inp = INPUT(
        type=INPUT_KEYBOARD,
        ki=KEYBDINPUT(
            wVk=vk,
            wScan=0,
            dwFlags=flags,
            time=0,
            dwExtraInfo=0,
        ),
    )
    sent = user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
    if sent != 1:
        raise ctypes.WinError(ctypes.get_last_error())


def send_ctrl_v() -> None:
    _send_key(VK_CONTROL, key_up=False)
    _send_key(VK_V, key_up=False)
    _send_key(VK_V, key_up=True)
    _send_key(VK_CONTROL, key_up=True)


def send_enter() -> None:
    _send_key(VK_RETURN, key_up=False)
    _send_key(VK_RETURN, key_up=True)


def send_backspaces(count: int) -> None:
    for _ in range(max(0, count)):
        _send_key(VK_BACK, key_up=False)
        _send_key(VK_BACK, key_up=True)


def send_unicode_text(text: str) -> None:
    # Send UTF-16 code units so input works without clipboard dependency.
    units = text.encode("utf-16-le")
    for idx in range(0, len(units), 2):
        code_unit = int.from_bytes(units[idx : idx + 2], "little")
        key_down = INPUT(
            type=INPUT_KEYBOARD,
            ki=KEYBDINPUT(
                wVk=0,
                wScan=code_unit,
                dwFlags=KEYEVENTF_UNICODE,
                time=0,
                dwExtraInfo=0,
            ),
        )
        key_up = INPUT(
            type=INPUT_KEYBOARD,
            ki=KEYBDINPUT(
                wVk=0,
                wScan=code_unit,
                dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
                time=0,
                dwExtraInfo=0,
            ),
        )
        sent_down = user32.SendInput(1, ctypes.byref(key_down), ctypes.sizeof(INPUT))
        sent_up = user32.SendInput(1, ctypes.byref(key_up), ctypes.sizeof(INPUT))
        if sent_down != 1 or sent_up != 1:
            raise ctypes.WinError(ctypes.get_last_error())


def _post_message(hwnd: int, msg: int, wparam: int = 0, lparam: int = 0) -> None:
    if not user32.PostMessageW(hwnd, msg, wparam, lparam):
        raise ctypes.WinError(ctypes.get_last_error())


def _get_input_hwnd_for_thread(top_hwnd: int) -> int:
    thread_id = user32.GetWindowThreadProcessId(top_hwnd, None)
    if not thread_id:
        return top_hwnd

    info = GUITHREADINFO()
    info.cbSize = ctypes.sizeof(GUITHREADINFO)
    if user32.GetGUIThreadInfo(thread_id, ctypes.byref(info)) and info.hwndFocus:
        return int(info.hwndFocus)
    return top_hwnd


def post_text_background(hwnd: int, text: str, send_newline: bool) -> None:
    target_hwnd = _get_input_hwnd_for_thread(hwnd)
    for ch in text:
        _post_message(target_hwnd, WM_CHAR, ord(ch), 1)
    if send_newline:
        _post_message(target_hwnd, WM_KEYDOWN, VK_RETURN, 0x001C0001)
        _post_message(target_hwnd, WM_KEYUP, VK_RETURN, 0xC01C0001)


def post_backspaces_background(hwnd: int, count: int) -> None:
    target_hwnd = _get_input_hwnd_for_thread(hwnd)
    for _ in range(max(0, count)):
        _post_message(target_hwnd, WM_KEYDOWN, VK_BACK, 0x000E0001)
        _post_message(target_hwnd, WM_CHAR, VK_BACK, 1)
        _post_message(target_hwnd, WM_KEYUP, VK_BACK, 0xC00E0001)


def force_foreground(hwnd: int) -> None:
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)
    else:
        user32.ShowWindow(hwnd, SW_SHOW)
    user32.BringWindowToTop(hwnd)

    current_thread = kernel32.GetCurrentThreadId()
    fg_hwnd = user32.GetForegroundWindow()
    fg_thread = user32.GetWindowThreadProcessId(fg_hwnd, None) if fg_hwnd else 0
    target_thread = user32.GetWindowThreadProcessId(hwnd, None)

    attached_fg = False
    attached_target = False

    try:
        if fg_thread and fg_thread != current_thread:
            user32.AttachThreadInput(fg_thread, current_thread, True)
            attached_fg = True
        if target_thread and target_thread != current_thread:
            user32.AttachThreadInput(target_thread, current_thread, True)
            attached_target = True
        if not user32.SetForegroundWindow(hwnd):
            time.sleep(0.03)
            user32.SetForegroundWindow(hwnd)
    finally:
        if attached_target:
            user32.AttachThreadInput(target_thread, current_thread, False)
        if attached_fg:
            user32.AttachThreadInput(fg_thread, current_thread, False)


def normalize_text(text: str) -> str:
    return " ".join(text.lower().strip().split())


def ensure_sentence_dot_space(text: str) -> str:
    clean = " ".join(text.strip().split())
    if not clean:
        return ""
    clean = re.sub(r"[.?!,:;…]+$", "", clean)
    return f"{clean}. "


def _normalize_token(token: str) -> str:
    return re.sub(r"[^\w\-]", "", token.lower().strip(), flags=re.UNICODE)


def load_corrections(path: Optional[str]) -> tuple[dict[str, str], dict[str, str]]:
    if not path:
        return {}, {}

    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError("Corrections file must be a JSON object: {\"wrong\":\"right\"}")

    phrase_map: dict[str, str] = {}
    word_map: dict[str, str] = {}
    for src, dst in raw.items():
        if not isinstance(src, str) or not isinstance(dst, str):
            continue
        src_norm = normalize_text(src)
        dst_clean = " ".join(dst.strip().split())
        if not src_norm or not dst_clean:
            continue
        if " " in src_norm:
            phrase_map[src_norm] = dst_clean
        else:
            word_map[src_norm] = dst_clean
    return phrase_map, word_map


def apply_corrections(text: str, phrase_map: dict[str, str], word_map: dict[str, str]) -> str:
    normalized = normalize_text(text)
    if normalized in phrase_map:
        return phrase_map[normalized]

    out_words: list[str] = []
    for token in text.split():
        token_key = _normalize_token(token)
        replacement = word_map.get(token_key)
        out_words.append(replacement if replacement else token)
    return " ".join(out_words)


class DictationAgent:
    def __init__(
        self,
        recognizer: object,
        recognizer_backend: str,
        target_title: str,
        auto_enter: bool,
        input_mode: str,
        wake_word: str,
        min_chars: int,
        phrase_corrections: dict[str, str],
        word_corrections: dict[str, str],
    ) -> None:
        self.recognizer = recognizer
        self.recognizer_backend = recognizer_backend
        self.target_title = target_title
        self.auto_enter = auto_enter
        self.input_mode = input_mode
        self.target_resolver = TargetWindowResolver(target_title)
        self.wake_word_norm = normalize_text(wake_word) if wake_word else ""
        self.min_chars = min_chars
        self.phrase_corrections = phrase_corrections
        self.word_corrections = word_corrections
        self.paused = False
        self.running = True
        self.last_focus_ts = 0.0
        self.last_action_key = ""
        self.last_action_ts = 0.0
        self.pending_input_chars = 0

    def _get_target_hwnd(self) -> int:
        return self.target_resolver.get_hwnd()

    def _send_to_target(self, text: str) -> None:
        hwnd = self._get_target_hwnd()
        sent_len = len(text)
        if self.input_mode == "post":
            # In background-post mode Enter is always manual via voice "отправка".
            post_text_background(hwnd, text, send_newline=False)
            self.pending_input_chars += sent_len
            return

        force_foreground(hwnd)
        if self.input_mode == "type":
            send_unicode_text(text)
        else:
            # Optional clipboard mode for apps where direct key injection is blocked.
            pyperclip.copy(text)
            time.sleep(0.03)
            send_ctrl_v()

        if self.auto_enter:
            time.sleep(0.03)
            send_enter()
            self.pending_input_chars = 0
        else:
            self.pending_input_chars += sent_len

    def _submit_target(self) -> None:
        hwnd = self._get_target_hwnd()
        if self.input_mode == "post":
            post_text_background(hwnd, "", send_newline=True)
            self.pending_input_chars = 0
            return
        force_foreground(hwnd)
        send_enter()
        self.pending_input_chars = 0

    def _clear_target(self) -> None:
        hwnd = self._get_target_hwnd()
        clear_count = self.pending_input_chars
        if clear_count <= 0:
            return
        if self.input_mode == "post":
            post_backspaces_background(hwnd, clear_count)
            self.pending_input_chars = 0
            return
        force_foreground(hwnd)
        send_backspaces(clear_count)
        self.pending_input_chars = 0

    def _is_duplicate_action(self, key: str) -> bool:
        now = time.monotonic()
        if key == self.last_action_key and (now - self.last_action_ts) < DEFAULT_DUPLICATE_SUPPRESS_SEC:
            return True
        self.last_action_key = key
        self.last_action_ts = now
        return False

    @staticmethod
    def _pcm_rms(data: bytes) -> float:
        pcm = memoryview(data).cast("h")
        if not pcm:
            return 0.0
        sum_sq = 0.0
        for sample in pcm:
            sum_sq += float(sample) * float(sample)
        return (sum_sq / len(pcm)) ** 0.5

    def _focus_target_on_speech(self, rms: float) -> None:
        if self.input_mode == "post":
            return
        if rms < DEFAULT_SILENCE_RMS_THRESHOLD:
            return

        now = time.monotonic()
        if now - self.last_focus_ts < DEFAULT_FOCUS_COOLDOWN_SEC:
            return

        hwnd = self.target_resolver.try_get_hwnd()
        if hwnd is None:
            return

        force_foreground(hwnd)
        self.last_focus_ts = now

    def _handle_phrase(self, phrase: str) -> None:
        clean_phrase = " ".join(phrase.strip().split())
        normalized = normalize_text(clean_phrase)
        if not normalized:
            return

        if normalized == "stop dictation" or normalized == "стоп диктовка":
            print("[voice] stop")
            self.running = False
            return
        if normalized == "pause dictation" or normalized == "пауза диктовка":
            print("[voice] paused")
            self.paused = True
            return
        if (
            normalized == "resume dictation"
            or normalized == "continue dictation"
            or normalized == "продолжай диктовка"
        ):
            print("[voice] resumed")
            self.paused = False
            return
        if normalized in {"send", "submit", "enter", "отправка", "отправить", "энтер"}:
            if self._is_duplicate_action("cmd:submit"):
                return
            print("[voice] submit")
            self._submit_target()
            return
        if normalized in {"clear", "очистить", "очистка", "очисти"}:
            if self._is_duplicate_action("cmd:clear"):
                return
            print("[voice] clear")
            self._clear_target()
            return

        if self.paused:
            return

        out_text = clean_phrase
        if self.wake_word_norm:
            if not normalized.startswith(self.wake_word_norm):
                return
            wake_len = len(self.wake_word_norm.split())
            out_tokens = clean_phrase.split()[wake_len:]
            out_text = " ".join(out_tokens).strip()
            if not out_text:
                return

        out_text = apply_corrections(out_text, self.phrase_corrections, self.word_corrections)
        if len(out_text) < self.min_chars:
            return
        if self._is_duplicate_action(f"text:{normalize_text(out_text)}"):
            return

        out_text = ensure_sentence_dot_space(out_text)
        if not out_text:
            return
        print(f"[send] {out_text.rstrip()}")
        self._send_to_target(out_text)

    def _run_vosk(self, audio_queue: queue.Queue[bytes], stop_event: threading.Event) -> None:
        while self.running and not stop_event.is_set():
            data = audio_queue.get()
            self._focus_target_on_speech(self._pcm_rms(data))
            if self.recognizer.AcceptWaveform(data):
                result = json.loads(self.recognizer.Result())
                text = result.get("text", "").strip()
                if text:
                    self._handle_phrase(text)

    def _decode_sherpa_offline(self, samples: list[float]) -> str:
        stream = self.recognizer.create_stream()
        stream.accept_waveform(DEFAULT_SAMPLE_RATE, samples)
        self.recognizer.decode_stream(stream)
        return stream.result.text.strip()

    def _run_sherpa(self, audio_queue: queue.Queue[bytes], stop_event: threading.Event) -> None:
        in_speech = False
        speech_samples: list[float] = []
        last_speech_ts = 0.0
        min_utt_samples = int(DEFAULT_MIN_UTTERANCE_SEC * DEFAULT_SAMPLE_RATE)
        max_utt_samples = int(DEFAULT_MAX_UTTERANCE_SEC * DEFAULT_SAMPLE_RATE)

        while self.running and not stop_event.is_set():
            data = audio_queue.get()
            pcm = memoryview(data).cast("h")
            if not pcm:
                continue

            rms = self._pcm_rms(data)
            self._focus_target_on_speech(rms)
            is_speech = rms >= DEFAULT_SILENCE_RMS_THRESHOLD

            now = time.monotonic()
            if is_speech and not in_speech:
                in_speech = True
                speech_samples = []
                last_speech_ts = now

            if not in_speech:
                continue

            speech_samples.extend(sample / 32768.0 for sample in pcm)
            if is_speech:
                last_speech_ts = now

            silence_elapsed = (not is_speech) and (
                now - last_speech_ts >= DEFAULT_END_OF_UTTERANCE_SILENCE_SEC
            )
            too_long = len(speech_samples) >= max_utt_samples
            if not silence_elapsed and not too_long:
                continue

            if len(speech_samples) >= min_utt_samples:
                text = self._decode_sherpa_offline(speech_samples)
                if text:
                    self._handle_phrase(text)

            in_speech = False
            speech_samples = []

    def run(self) -> None:
        audio_queue: queue.Queue[bytes] = queue.Queue()
        stop_event = threading.Event()

        def _audio_callback(indata, _frames, _time_info, status):
            if status:
                print(f"[audio] {status}", file=sys.stderr)
            audio_queue.put(bytes(indata))

        print("[init] Listening...")
        with sd.RawInputStream(
            samplerate=DEFAULT_SAMPLE_RATE,
            blocksize=8000,
            dtype="int16",
            channels=1,
            callback=_audio_callback,
        ):
            if self.recognizer_backend == "vosk":
                self._run_vosk(audio_queue, stop_event)
            elif self.recognizer_backend == "sherpa-onnx":
                self._run_sherpa(audio_queue, stop_event)
            else:
                raise RuntimeError(f"Unsupported recognizer backend: {self.recognizer_backend}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Always-on dictation to a target window using Vosk or Sherpa-ONNX."
    )
    parser.add_argument(
        "--target-title",
        required=True,
        help="Substring of target window title (e.g. 'ChatGPT').",
    )
    parser.add_argument(
        "--model-dir",
        default=None,
        help=(
            "Model directory. Supports classic Vosk models (conf/graph) and "
            "Zipformer ONNX models (am-onnx/lang), e.g. vosk-model-ru-0.54."
        ),
    )
    parser.add_argument(
        "--model-url",
        default=None,
        help="Model zip URL for auto download. Overrides --quality profile URL.",
    )
    parser.add_argument(
        "--quality",
        choices=sorted(MODEL_PROFILES.keys()),
        default=DEFAULT_MODEL_PROFILE,
        help="Recognition quality profile: fast=small model, best=large model.",
    )
    parser.add_argument(
        "--auto-download-model",
        action="store_true",
        help="Download model automatically if model-dir does not exist.",
    )
    parser.add_argument(
        "--wake-word",
        default="",
        help="Optional wake word. Only phrases starting with this text are sent.",
    )
    parser.add_argument(
        "--min-chars",
        type=int,
        default=2,
        help="Ignore recognized phrases shorter than this number of characters.",
    )
    parser.add_argument(
        "--no-enter",
        action="store_true",
        help="Do not press Enter after sending recognized text.",
    )
    parser.add_argument(
        "--input-mode",
        choices=["type", "paste", "post"],
        default="type",
        help=(
            "Text injection mode: direct typing (type), clipboard paste (paste), "
            "or PostMessage without focus switch (post, best for console windows; "
            "Enter only by voice command send/отправка)."
        ),
    )
    parser.add_argument(
        "--corrections-file",
        default=None,
        help="Path to JSON corrections file: {\"wrong\":\"right\"} for slang/oslyshki.",
    )
    return parser


def main() -> int:
    configure_stdio()
    parser = build_parser()
    args = parser.parse_args()

    profile = MODEL_PROFILES[args.quality]
    base_dir = Path(__file__).resolve().parent
    model_dir = Path(args.model_dir) if args.model_dir else (base_dir / "models" / profile["name"])
    model_url = args.model_url if args.model_url else profile["url"]

    model_dir = ensure_model(model_dir, model_url, auto_download=args.auto_download_model)
    backend = detect_model_backend(model_dir)
    print(f"[init] Model: {model_dir}")
    print(f"[init] Backend: {backend}")

    if backend == "vosk":
        model = Model(str(model_dir))
        recognizer = KaldiRecognizer(model, DEFAULT_SAMPLE_RATE)
        recognizer.SetWords(True)
    else:
        recognizer = build_sherpa_offline_recognizer(model_dir)

    phrase_corrections, word_corrections = load_corrections(args.corrections_file)

    agent = DictationAgent(
        recognizer=recognizer,
        recognizer_backend=backend,
        target_title=args.target_title,
        auto_enter=not args.no_enter,
        input_mode=args.input_mode,
        wake_word=args.wake_word,
        min_chars=args.min_chars,
        phrase_corrections=phrase_corrections,
        word_corrections=word_corrections,
    )

    print(f"[init] Target title contains: '{args.target_title}'")
    print(f"[init] Quality profile: {args.quality}")
    print(f"[init] Input mode: {args.input_mode}")
    print("[init] Voice commands are enabled (ru/en pause/resume/stop)")
    if args.wake_word:
        print(f"[init] Wake word: '{args.wake_word}'")

    try:
        agent.run()
    except KeyboardInterrupt:
        print("\n[exit] Interrupted by user")
    except Exception as exc:
        print(f"[error] {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
