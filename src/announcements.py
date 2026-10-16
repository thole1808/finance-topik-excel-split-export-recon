"""Queue short spoken status announcements using the host OS voice engine."""

from __future__ import annotations

import platform
import queue
import shutil
import subprocess
import threading

_messages: queue.Queue[str] = queue.Queue()


def _speech_worker() -> None:
    while True:
        message = _messages.get()
        try:
            current_platform = platform.system()
            if current_platform == "Darwin" and shutil.which("say"):
                subprocess.run(
                    ["say", message],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                )
            elif current_platform == "Windows":
                script = (
                    "Add-Type -AssemblyName System.Speech; "
                    "$voice = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
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
