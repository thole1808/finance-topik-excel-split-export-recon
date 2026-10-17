"""Check and install public GitHub releases for Recon Builder."""

from __future__ import annotations

import hashlib
import json
import os
import platform
import shlex
import shutil
import ssl
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

REPOSITORY = "thole1808/finance-topik-excel-split-export-recon"
RELEASE_API = f"https://api.github.com/repos/{REPOSITORY}/releases/latest"
USER_AGENT = "ReconBuilder-Updater"
BUFFER_SIZE = 1024 * 1024


def _ssl_context() -> ssl.SSLContext:
    """Return an SSL context that works in both frozen (PyInstaller) and source runs.

    On macOS PyInstaller bundles, the default ssl context cannot find CA
    certificates, causing CERTIFICATE_VERIFY_FAILED for all HTTPS requests.
    We try certifi first (bundled by PyInstaller via --collect-all certifi),
    then fall back to the system CA store.
    """
    # 1. Try certifi — covers Windows, Linux, and macOS PyInstaller bundles
    try:
        import certifi  # type: ignore[import-untyped]
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        pass

    # 2. macOS: use the system keychain CA bundle directly
    if platform.system() == "Darwin":
        system_ca = "/etc/ssl/cert.pem"
        if Path(system_ca).is_file():
            try:
                return ssl.create_default_context(cafile=system_ca)
            except Exception:
                pass

    # 3. Fall back to Python default (works on most Linux/Windows installs)
    return ssl.create_default_context()


class UpdateError(RuntimeError):
    """An update could not be checked, verified, downloaded, or started."""


@dataclass(frozen=True)
class UpdateRelease:
    version: str
    tag: str
    asset: dict[str, object]
    checksum_asset: dict[str, object]
    html_url: str


def _version_tuple(value: str) -> tuple[int, ...]:
    try:
        return tuple(int(part) for part in value.removeprefix("v").split("."))
    except ValueError as exc:
        raise UpdateError(f"Format versi rilis tidak valid: {value}") from exc


def _open_url(url: str, timeout: int = 20):
    if not url.startswith("https://"):
        raise UpdateError("URL update harus menggunakan HTTPS.")
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": USER_AGENT,
            "X-GitHub-Api-Version": "2026-03-10",
        },
    )
    try:
        response = urllib.request.urlopen(request, timeout=timeout, context=_ssl_context())
    except urllib.error.HTTPError:
        raise
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise UpdateError(f"Tidak dapat terhubung ke GitHub: {exc}") from exc
    if not response.geturl().startswith("https://"):
        response.close()
        raise UpdateError("Unduhan update dialihkan keluar dari koneksi HTTPS.")
    return response


def latest_update(current_version: str) -> UpdateRelease | None:
    """Return a newer release with a platform installer, if one is published."""
    try:
        with _open_url(RELEASE_API) as response:
            release = json.load(response)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return None
        raise UpdateError(f"GitHub menolak pemeriksaan update (HTTP {exc.code}).") from exc
    tag = str(release.get("tag_name", ""))
    version = tag.removeprefix("v")
    if not tag or _version_tuple(version) <= _version_tuple(current_version):
        return None
    assets = release.get("assets", [])
    if not isinstance(assets, list):
        raise UpdateError("Daftar file pada release GitHub tidak valid.")
    checksum_asset = next(
        (asset for asset in assets if asset.get("name") == "SHA256SUMS.txt"),
        None,
    )
    if platform.system() == "Darwin":
        asset = next(
            (
                item for item in assets
                if str(item.get("name", "")).startswith(f"ReconBuilder-v{version}-")
                and str(item.get("name", "")).endswith(".dmg")
            ),
            None,
        )
    elif platform.system() == "Windows":
        asset = next(
            (item for item in assets if item.get("name") == "ReconBuilder-Setup.exe"),
            None,
        )
    else:
        asset = None
    if asset is None:
        raise UpdateError(
            f"Rilis {tag} belum menyediakan updater untuk {platform.system()}."
        )
    if checksum_asset is None:
        raise UpdateError(f"Rilis {tag} tidak memiliki file checksum.")
    return UpdateRelease(
        version=version,
        tag=tag,
        asset=asset,
        checksum_asset=checksum_asset,
        html_url=str(release.get("html_url", "")),
    )


def _download(
    url: str,
    destination: Path,
    progress_callback: Callable[[int, int], None] | None = None,
    expected_size: int = 0,
) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    partial = destination.with_suffix(destination.suffix + ".part")
    try:
        completed = 0
        with _open_url(url, timeout=60) as response, partial.open("wb") as output:
            total = expected_size or int(response.headers.get("Content-Length", 0))
            while chunk := response.read(BUFFER_SIZE):
                output.write(chunk)
                completed += len(chunk)
                if progress_callback:
                    progress_callback(completed, total)
        partial.replace(destination)
    except Exception:
        partial.unlink(missing_ok=True)
        raise


def download_update(
    release: UpdateRelease,
    progress_callback: Callable[[int, int], None] | None = None,
) -> Path:
    """Download the matching installer and verify it against its release hash."""
    cache = Path.home() / "Library" / "Caches" / "Recon Builder" / "updates"
    if platform.system() == "Windows":
        cache = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "Recon Builder" / "updates"
    filename = str(release.asset.get("name", ""))
    if not filename or Path(filename).name != filename:
        raise UpdateError("Nama file updater pada release tidak valid.")
    destination = cache / filename
    checksum_path = cache / "SHA256SUMS.txt"
    expected_size = int(release.asset.get("size", 0))
    _download(str(release.checksum_asset.get("browser_download_url", "")), checksum_path)
    _download(
        str(release.asset.get("browser_download_url", "")), destination,
        progress_callback=progress_callback, expected_size=expected_size,
    )
    if expected_size and destination.stat().st_size != expected_size:
        destination.unlink(missing_ok=True)
        raise UpdateError("Ukuran file updater tidak sesuai metadata release.")
    expected_hash = None
    for line in checksum_path.read_text(encoding="utf-8").splitlines():
        fields = line.strip().split(maxsplit=1)
        if len(fields) == 2 and fields[1].lstrip("* ") == filename:
            expected_hash = fields[0].lower()
            break
    if not expected_hash or len(expected_hash) != 64:
        raise UpdateError(f"Checksum untuk {filename} tidak ditemukan.")
    digest = hashlib.sha256()
    with destination.open("rb") as source:
        while chunk := source.read(BUFFER_SIZE):
            digest.update(chunk)
    if digest.hexdigest() != expected_hash:
        destination.unlink(missing_ok=True)
        raise UpdateError("Checksum updater tidak cocok; file tidak akan dipasang.")
    return destination


def _macos_app_bundle() -> Path:
    for parent in Path(sys.executable).resolve().parents:
        if parent.suffix == ".app":
            return parent
    raise UpdateError("Aplikasi harus dipasang dari ReconBuilder.app untuk update.")


def _start_macos_update(dmg_path: Path, process_id: int) -> None:
    app_bundle = _macos_app_bundle()
    if "/Volumes/" in str(app_bundle):
        raise UpdateError("Salin ReconBuilder.app ke Applications sebelum update.")
    cache = dmg_path.parent
    mountpoint = Path(tempfile.mkdtemp(prefix="recon-update-mount-", dir=cache))
    try:
        subprocess.run(
            ["/usr/bin/hdiutil", "attach", "-quiet", "-nobrowse", "-readonly",
             "-mountpoint", str(mountpoint), str(dmg_path)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
        )
    except (OSError, subprocess.CalledProcessError) as exc:
        shutil.rmtree(mountpoint, ignore_errors=True)
        detail = getattr(exc, "stderr", "") or str(exc)
        raise UpdateError(f"DMG update tidak dapat dibuka: {detail}") from exc
    source_bundle = next(mountpoint.glob("*.app"), None)
    source_package = next(mountpoint.glob("*.pkg"), None)
    target = app_bundle
    if source_package is not None:
        target = Path("/Applications/Recon Builder.app")
        command = " ".join((
            f"while /bin/kill -0 {int(process_id)} 2>/dev/null; do /bin/sleep 1; done;",
            f"trap '/usr/bin/hdiutil detach {shlex.quote(str(mountpoint))} -quiet >/dev/null 2>&1 || true' EXIT;",
            f"/usr/sbin/installer -pkg {shlex.quote(str(source_package))} -target / &&",
            f"/bin/launchctl asuser $(/usr/bin/stat -f%u /dev/console) "
            f"/usr/bin/open {shlex.quote(str(target))}",
        ))
        try:
            applescript_command = command.replace("\\", "\\\\").replace('"', '\\"')
            applescript = f'do shell script "{applescript_command}" with administrator privileges'
            subprocess.Popen(
                ["/usr/bin/osascript", "-e", applescript],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
        except OSError as exc:
            subprocess.run(["/usr/bin/hdiutil", "detach", str(mountpoint), "-quiet"], check=False)
            shutil.rmtree(mountpoint, ignore_errors=True)
            raise UpdateError(f"Updater macOS tidak dapat dijalankan: {exc}") from exc
        return
    if source_bundle is None:
        subprocess.run(["/usr/bin/hdiutil", "detach", str(mountpoint), "-quiet"], check=False)
        shutil.rmtree(mountpoint, ignore_errors=True)
        raise UpdateError("DMG tidak berisi ReconBuilder.app.")
    staging = target.with_name(f".{target.name}.update")
    backup = target.with_name(f".{target.name}.previous")
    command = " ".join((
        f"while /bin/kill -0 {int(process_id)} 2>/dev/null; do /bin/sleep 1; done;",
        f"trap '/usr/bin/hdiutil detach {shlex.quote(str(mountpoint))} -quiet >/dev/null 2>&1 || true' EXIT;",
        f"/bin/rm -rf {shlex.quote(str(staging))} {shlex.quote(str(backup))} &&",
        f"/usr/bin/ditto {shlex.quote(str(source_bundle))} {shlex.quote(str(staging))} &&",
        f"(/bin/mv {shlex.quote(str(target))} {shlex.quote(str(backup))} || exit 1) &&",
        f"(/bin/mv {shlex.quote(str(staging))} {shlex.quote(str(target))} || {{ /bin/mv {shlex.quote(str(backup))} {shlex.quote(str(target))}; exit 1; }}) &&",
        f"/bin/rm -rf {shlex.quote(str(backup))} && "
        f"/bin/launchctl asuser $(/usr/bin/stat -f%u /dev/console) "
        f"/usr/bin/open {shlex.quote(str(target))}",
    ))
    try:
        if os.access(target.parent, os.W_OK):
            subprocess.Popen(
                ["/bin/sh", "-c", command],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
        else:
            applescript_command = command.replace("\\", "\\\\").replace('"', '\\"')
            applescript = f'do shell script "{applescript_command}" with administrator privileges'
            subprocess.Popen(
                ["/usr/bin/osascript", "-e", applescript],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
    except OSError as exc:
        subprocess.run(["/usr/bin/hdiutil", "detach", str(mountpoint), "-quiet"], check=False)
        shutil.rmtree(mountpoint, ignore_errors=True)
        raise UpdateError(f"Updater macOS tidak dapat dijalankan: {exc}") from exc


def _powershell_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _start_windows_update(installer_path: Path, process_id: int) -> None:
    """Wait for this app to exit, install silently, then reopen the new app."""
    executable = Path(sys.executable).resolve()
    if executable.name.lower() != "reconbuilder.exe":
        raise UpdateError("Lokasi aplikasi Windows tidak dikenali untuk update.")
    script = "\n".join((
        f"$waitPid = {int(process_id)}",
        f"$installer = {_powershell_quote(str(installer_path.resolve()))}",
        f"$application = {_powershell_quote(str(executable))}",
        "while (Get-Process -Id $waitPid -ErrorAction SilentlyContinue) {",
        "  Start-Sleep -Seconds 1",
        "}",
        "try {",
        "  $setup = Start-Process -FilePath $installer -Verb RunAs -Wait -PassThru `",
        "    -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART'",
        "  if ($setup.ExitCode -ne 0) { throw \"Installer exit code $($setup.ExitCode)\" }",
        "} catch {",
        "  Start-Process -FilePath $application",
        "  exit 1",
        "}",
        "Start-Process -FilePath $application",
    ))
    try:
        subprocess.Popen(
            [
                "powershell.exe", "-NoProfile", "-NonInteractive",
                "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
                "-Command", script,
            ],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except OSError as exc:
        raise UpdateError(f"Updater Windows tidak dapat dijalankan: {exc}") from exc


def start_installer(installer_path: Path, process_id: int) -> None:
    """Launch the signed-platform installer and let it replace the running app."""
    if not getattr(sys, "frozen", False):
        raise UpdateError("Update otomatis hanya tersedia pada aplikasi terpasang.")
    if platform.system() == "Darwin":
        _start_macos_update(installer_path, process_id)
        return
    if platform.system() == "Windows":
        _start_windows_update(installer_path, process_id)
        return
    raise UpdateError("Update otomatis hanya tersedia untuk macOS dan Windows.")
