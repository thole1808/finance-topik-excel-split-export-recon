"""Check and install public GitHub releases for Auto Recon Biller."""

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
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

REPOSITORY = "thole1808/finance-topik-excel-split-export-recon"
RELEASE_API = f"https://api.github.com/repos/{REPOSITORY}/releases/latest"
USER_AGENT = "ReconBuilder-Updater"
BUFFER_SIZE = 1024 * 1024


def _find_ca_bundle() -> str | None:
    """Cari file CA bundle yang tersedia di sistem untuk verifikasi SSL.

    Urutan prioritas:
    1. certifi (bundled via PyInstaller --collect-all certifi)
    2. macOS: /etc/ssl/cert.pem, ~/Library/Application Support/certifi/cacert.pem
    3. Linux: /etc/ssl/certs/ca-certificates.crt, /etc/pki/tls/certs/ca-bundle.crt
    4. Windows: default Python SSL context (menggunakan Windows certificate store)
    """
    # 1. certifi — terbaik untuk frozen PyInstaller bundle di semua OS
    try:
        import certifi  # type: ignore[import-untyped]
        ca = certifi.where()
        if ca and Path(ca).is_file():
            return ca
    except Exception:
        pass

    system = platform.system()

    # 2. macOS system CA bundles
    if system == "Darwin":
        for candidate in (
            "/etc/ssl/cert.pem",
            "/etc/openssl/cert.pem",
            str(Path.home() / "Library" / "Application Support" / "certifi" / "cacert.pem"),
        ):
            if Path(candidate).is_file():
                return candidate

    # 3. Linux system CA bundles
    elif system == "Linux":
        for candidate in (
            "/etc/ssl/certs/ca-certificates.crt",
            "/etc/pki/tls/certs/ca-bundle.crt",
            "/etc/pki/tls/certs/ca-bundle.trust.crt",
            "/usr/lib/ssl/certs/ca-certificates.crt",
        ):
            if Path(candidate).is_file():
                return candidate

    # 4. Windows: biarkan Python default (mengakses Windows certificate store)
    return None


def _ssl_context() -> ssl.SSLContext:
    """Return an SSL context that works in both frozen (PyInstaller) and source runs.

    On macOS PyInstaller bundles, the default ssl context cannot find CA
    certificates, causing CERTIFICATE_VERIFY_FAILED for all HTTPS requests.
    We try certifi first (bundled by PyInstaller via --collect-all certifi),
    then fall back to the system CA store.
    """
    ca_path = _find_ca_bundle()
    if ca_path:
        try:
            return ssl.create_default_context(cafile=ca_path)
        except Exception:
            pass

    # Fallback ke default Python (biasanya cukup untuk Windows dan Linux
    # dengan CA store sistem yang lengkap)
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


@dataclass(frozen=True)
class InstallerDownloads:
    version: str
    macos_name: str
    macos_url: str
    windows_name: str
    windows_url: str
    release_url: str


def _version_tuple(value: str) -> tuple[int, ...]:
    try:
        return tuple(int(part) for part in value.removeprefix("v").split("."))
    except ValueError as exc:
        raise UpdateError(f"Format versi rilis tidak valid: {value}") from exc


def _retry_after_seconds(exc: urllib.error.HTTPError) -> int | None:
    """Hitung detik hingga rate-limit GitHub di-reset, jika header tersedia."""
    try:
        reset_header = exc.headers.get("X-RateLimit-Reset")
        if reset_header:
            reset_at = int(reset_header)
            now = int(time.time())
            return max(reset_at - now, 0)
    except (ValueError, TypeError, AttributeError):
        pass
    return None


def _open_url(url: str, timeout: int = 30, retries: int = 5):
    if not url.startswith("https://"):
        raise UpdateError("URL update harus menggunakan HTTPS.")
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": USER_AGENT,
        "X-GitHub-Api-Version": "2026-03-10",
    }
    token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    if token:
        headers["Authorization"] = f"Bearer {token}"
    request = urllib.request.Request(url, headers=headers)

    last_error: Exception | None = None
    for attempt in range(retries):
        try:
            response = urllib.request.urlopen(
                request, timeout=timeout, context=_ssl_context()
            )
        except urllib.error.HTTPError as exc:
            last_error = exc
            if exc.code in (403, 422, 500, 502, 503, 504) and attempt < retries - 1:
                delay = _retry_after_seconds(exc)
                time.sleep(delay if delay is not None else (1 + attempt * 2))
                continue
            raise
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            last_error = exc
            if attempt < retries - 1:
                time.sleep(1 + attempt * 2)
                continue
            raise UpdateError(f"Tidak dapat terhubung ke server update: {exc}") from exc
        if not response.geturl().startswith("https://"):
            response.close()
            raise UpdateError("Unduhan update dialihkan keluar dari koneksi HTTPS.")
        return response
    if last_error is not None:
        raise UpdateError(f"Gagal memuat update setelah {retries} percobaan: {last_error}")
    raise UpdateError("Gagal memuat update dari server setelah beberapa percobaan.")


def latest_update(current_version: str) -> UpdateRelease | None:
    """Return a newer release with a platform installer, if one is published."""
    try:
        with _open_url(RELEASE_API) as response:
            release = json.load(response)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return None
        reset = _retry_after_seconds(exc)
        hint = f" Coba lagi dalam {reset} detik." if reset else ""
        raise UpdateError(
            f"GitHub menolak pemeriksaan update (HTTP {exc.code}).{hint}"
        ) from exc
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


def latest_installers() -> InstallerDownloads:
    """Find the current macOS and Windows installers on the latest release."""
    try:
        with _open_url(RELEASE_API) as response:
            release = json.load(response)
    except urllib.error.HTTPError as exc:
        reset = _retry_after_seconds(exc)
        hint = f" Coba lagi dalam {reset} detik." if reset else ""
        raise UpdateError(
            f"GitHub menolak pemeriksaan file aplikasi (HTTP {exc.code}).{hint}"
        ) from exc
    tag = str(release.get("tag_name", ""))
    version = tag.removeprefix("v")
    if not tag or not isinstance(release.get("assets"), list):
        raise UpdateError("Rilis terbaru tidak berisi daftar file installer.")
    assets = release["assets"]
    macos = next((
        item for item in assets
        if str(item.get("name", "")).startswith(f"ReconBuilder-v{version}-")
        and str(item.get("name", "")).endswith(".dmg")
    ), None)
    windows = next((
        item for item in assets if item.get("name") == "ReconBuilder-Setup.exe"
    ), None)
    if macos is None or windows is None:
        raise UpdateError("Rilis terbaru belum menyediakan installer Mac dan Windows.")
    macos_url = str(macos.get("browser_download_url", ""))
    windows_url = str(windows.get("browser_download_url", ""))
    release_url = str(release.get("html_url", ""))
    if not all(url.startswith("https://github.com/") for url in (
        macos_url, windows_url, release_url,
    )):
        raise UpdateError("Tautan installer dari rilis tidak valid.")
    return InstallerDownloads(
        version=version,
        macos_name=str(macos["name"]),
        macos_url=macos_url,
        windows_name=str(windows["name"]),
        windows_url=windows_url,
        release_url=release_url,
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
    raise UpdateError("Aplikasi harus dipasang dari Auto Recon Biller.app untuk update.")


def _start_macos_update(dmg_path: Path, process_id: int) -> None:
    app_bundle = _macos_app_bundle()
    if "/Volumes/" in str(app_bundle):
        raise UpdateError("Salin Auto Recon Biller.app ke Applications sebelum update.")
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
        target = Path("/Applications/Auto Recon Biller.app")
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
        raise UpdateError("DMG tidak berisi Auto Recon Biller.app.")
    target = Path("/Applications") / source_bundle.name
    legacy_target = app_bundle if app_bundle != target else None
    staging = target.with_name(f".{target.name}.update")
    backup = target.with_name(f".{target.name}.previous")
    command_parts = [
        f"while /bin/kill -0 {int(process_id)} 2>/dev/null; do /bin/sleep 1; done;",
        f"trap '/usr/bin/hdiutil detach {shlex.quote(str(mountpoint))} -quiet >/dev/null 2>&1 || true' EXIT;",
        f"/bin/rm -rf {shlex.quote(str(staging))} {shlex.quote(str(backup))} &&",
        f"/usr/bin/ditto {shlex.quote(str(source_bundle))} {shlex.quote(str(staging))} &&",
        f"if [ -e {shlex.quote(str(target))} ]; then /bin/mv {shlex.quote(str(target))} {shlex.quote(str(backup))} || exit 1; fi &&",
        f"(/bin/mv {shlex.quote(str(staging))} {shlex.quote(str(target))} || {{ if [ -e {shlex.quote(str(backup))} ]; then /bin/mv {shlex.quote(str(backup))} {shlex.quote(str(target))}; fi; exit 1; }}) &&",
        f"/bin/rm -rf {shlex.quote(str(backup))} &&",
    ]
    if legacy_target is not None:
        command_parts.append(f"/bin/rm -rf {shlex.quote(str(legacy_target))} &&")
    command_parts.extend((
        f"/bin/launchctl asuser $(/usr/bin/stat -f%u /dev/console) "
        f"/usr/bin/open {shlex.quote(str(target))}",
    ))
    command = " ".join(command_parts)
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
