"""Queue short spoken status announcements using the host OS voice engine."""

from __future__ import annotations

import platform
import queue
import shutil
import subprocess
import threading
import json

from src.audit_log import audit_directory

_messages: queue.Queue[str] = queue.Queue()
_macos_voice: str | None | bool = False
_preference_path = audit_directory() / "voice-settings.json"
_state_lock = threading.Lock()


def _load_enabled() -> bool:
    try:
        value = json.loads(_preference_path.read_text(encoding="utf-8"))
        return bool(value.get("enabled", True)) if isinstance(value, dict) else True
    except (OSError, json.JSONDecodeError):
        return True


_enabled = _load_enabled()


def voice_enabled() -> bool:
    with _state_lock:
        return _enabled


def set_voice_enabled(enabled: bool) -> None:
    """Persist the operator's voice announcement preference."""
    global _enabled
    with _state_lock:
        _enabled = bool(enabled)
        _preference_path.parent.mkdir(parents=True, exist_ok=True)
        _preference_path.write_text(
            json.dumps({"enabled": _enabled}, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )


def _indonesian_macos_voice() -> str | None:
    """Choose an installed Indonesian voice, if this Mac provides one."""
    global _macos_voice
    if _macos_voice is False:
        try:
            voices = subprocess.run(
                ["say", "-v", "?"], capture_output=True, text=True,
                timeout=5, check=False,
            ).stdout
            _macos_voice = next(
                (
                    line.split()[0] for line in voices.splitlines()
                    if len(line.split()) > 1 and line.split()[1] == "id_ID"
                ),
                None,
            )
        except (OSError, subprocess.TimeoutExpired):
            _macos_voice = None
    return _macos_voice or None


def _speech_worker() -> None:
    while True:
        message = _messages.get()
        try:
            if not voice_enabled():
                continue
            current_platform = platform.system()
            if current_platform == "Darwin" and shutil.which("say"):
                voice = _indonesian_macos_voice()
                command = ["say"]
                if voice:
                    command.extend(["-v", voice])
                command.append(message)
                subprocess.run(
                    command,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                )
            elif current_platform == "Windows":
                script = (
                    "Add-Type -AssemblyName System.Speech; "
                    "$voice = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
                    "$idVoice = $voice.GetInstalledVoices() | "
                    "Where-Object { $_.VoiceInfo.Culture.Name -like 'id*' } | "
                    "Select-Object -First 1; "
                    "if ($idVoice) { $voice.SelectVoice($idVoice.VoiceInfo.Name) }; "
                    f"$voice.Speak('{message}')"
                )
                subprocess.run(
                    ["powershell.exe", "-NoProfile", "-NonInteractive",
                     "-WindowStyle", "Hidden", "-Command", script],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                )
            else:
                speech_tool = shutil.which("spd-say") or shutil.which("espeak")
                if speech_tool:
                    subprocess.run(
                        [speech_tool, message],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        check=False,
                    )
        except OSError:
            pass
        finally:
            _messages.task_done()


_worker = threading.Thread(
    target=_speech_worker, name="recon-voice-announcements", daemon=True
)
_worker.start()


def announce(message: str) -> None:
    """Speak a message without blocking the interface; preserve message order."""
    if voice_enabled():
        _messages.put(message)
