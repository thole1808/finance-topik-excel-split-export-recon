"""Queue short spoken status announcements using the host OS voice engine."""

from __future__ import annotations

import platform
import queue
import shutil
import subprocess
import threading

_messages: queue.Queue[str] = queue.Queue()
_macos_voice: str | None | bool = False


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
    _messages.put(message)
