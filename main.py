"""
PSA-DIAG FREE
"""
import sys
from pathlib import Path
from PySide6 import QtCore, QtGui, QtWidgets
import psutil
import platform
import requests #type:ignore
import os
import time
import shutil
import ctypes
import glob
import subprocess
import logging
import re
import threading
from datetime import datetime
import json

# Determine base path for resources.
if getattr(sys, 'frozen', False):
    BASE = Path(sys._MEIPASS)
else:
    BASE = Path(__file__).resolve().parent

# Centralized configuration (moved to `config.py`)
from config import CONFIG_DIR, APP_VERSION, URL_LAST_VERSION_PSADIAG, URL_LAST_VERSION_DIAGBOX, URL_VERSION_OPTIONS, URL_REMOTE_MESSAGES, ARCHIVE_PASSWORD

# Translation system
class Translator:
    def __init__(self, language='en'):
        self.language = language
        self.translations = {}
        self.load_translations()
    
    def load_translations(self):
        """Load translation file for current language"""
        lang_file = BASE / "lang" / f"{self.language}.json"
        try:
            if lang_file.exists():
                with open(lang_file, 'r', encoding='utf-8') as f:
                    self.translations = json.load(f)
                # Logger may not be initialized yet during early import
                if 'logger' in globals():
                    logger.info(f"Loaded translations for language: {self.language}")
            else:
                if 'logger' in globals():
                    logger.warning(f"Translation file not found: {lang_file}")
        except Exception as e:
            if 'logger' in globals():
                logger.error(f"Error loading translations: {e}")
    
    def t(self, key, **kwargs):
        """Translate a key with optional formatting parameters"""
        keys = key.split('.')
        value = self.translations
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                if 'logger' in globals():
                    logger.warning(f"Translation key not found: {key}")
                return key
        
        # Format with kwargs if provided
        if kwargs:
            try:
                return value.format(**kwargs)
            except:
                return value
        return value
    
    def set_language(self, language):
        """Change the current language"""
        self.language = language
        self.load_translations()
        self.save_language_preference()
    
    def save_language_preference(self):
        """Save language preference to file"""
        try:
            prefs_file = CONFIG_DIR / "preferences.json"
            prefs_file.parent.mkdir(parents=True, exist_ok=True)
            with open(prefs_file, 'w', encoding='utf-8') as f:
                json.dump({'language': self.language}, f)
        except Exception as e:
            if 'logger' in globals():
                logger.error(f"Failed to save language preference: {e}")
    
    def load_language_preference(self):
        """Load language preference from file"""
        try:
            prefs_file = CONFIG_DIR / "preferences.json"
            if prefs_file.exists():
                with open(prefs_file, 'r', encoding='utf-8') as f:
                    prefs = json.load(f)
                    return prefs.get('language', 'fr')
        except Exception as e:
            if 'logger' in globals():
                logger.error(f"Failed to load language preference: {e}")
        return 'fr'  # Default to French

# Global translator instance
translator = Translator(Translator('en').load_language_preference())  # Load saved preference

# Configure logging
# Always write logs to the persistent config directory so the executable
# (and dev runs) both use the same location. This ensures logs persist
# across runs and are accessible from the packaged executable.
log_folder = CONFIG_DIR / "logs"
log_folder.mkdir(parents=True, exist_ok=True)
log_file = log_folder / f"psa_diag_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8')
        # StreamHandler removed - console will be integrated in UI
    ]
)

logger = logging.getLogger(__name__)
logger.info(f"PSA-DIAG v{APP_VERSION} starting...")
logger.info(f"Log file path: {log_file}")

class QTextEditLogger(logging.Handler):
    """Custom logging handler that writes to a QTextEdit widget"""
    def __init__(self, text_edit):
        super().__init__()
        self.text_edit = text_edit
    
    def emit(self, record):
        msg = self.format(record)
        # Use invokeMethod to ensure thread safety
        QtCore.QMetaObject.invokeMethod(
            self.text_edit,
            "append",
            QtCore.Qt.ConnectionType.QueuedConnection,
            QtCore.Q_ARG(str, msg)
        )

def hide_console():
    """Hide the console window on Windows"""
    if sys.platform == 'win32':
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            user32 = ctypes.windll.user32
            
            # Get console window handle
            hwnd = kernel32.GetConsoleWindow()
            if hwnd:
                # Hide the console window
                user32.ShowWindow(hwnd, 0)  # SW_HIDE = 0
        except Exception as e:
            logger.error(f"Failed to hide console: {e}")

def kill_updater_processes():
    """Terminate any leftover updater.exe processes from previous runs."""
    try:
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                name = (proc.info.get('name') or '').lower()
                exe = proc.info.get('exe') or ''
                if name == 'updater.exe' or (exe and os.path.basename(exe).lower() == 'updater.exe'):
                    logger.info(f"Terminating leftover updater.exe PID={proc.pid}")
                    try:
                        proc.terminate()
                        proc.wait(timeout=2)
                    except Exception:
                        try:
                            proc.kill()
                        except Exception:
                            logger.debug(f"Failed to kill updater PID={proc.pid}")
            except Exception as e:
                logger.debug(f"Error while checking process for updater: {e}")
    except Exception as e:
        logger.debug(f"Failed to enumerate processes to kill updater.exe: {e}")

def is_admin():
    """Check if the script is running with admin privileges"""
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        logger.info(f"Admin check: {is_admin}")
        return is_admin
    except Exception as e:
        logger.error(f"Admin check failed: {e}")
        return False

def run_as_admin():
    """Relaunch the script with admin privileges"""
    try:
        if sys.platform == 'win32':
            # Get the path to the Python executable and the script
            script = os.path.abspath(sys.argv[0])
            params = ' '.join([script] + sys.argv[1:])
            
            logger.info("Requesting admin elevation...")
            # Use ShellExecuteW to run as admin
            ctypes.windll.shell32.ShellExecuteW(
                None, 
                "runas", 
                sys.executable, 
                params, 
                None, 
                1  # SW_SHOWNORMAL
            )
            return True
    except Exception as e:
        logger.error(f"Failed to elevate privileges: {e}")
        return False
    return False

# Load style
def load_qss():
    qss_path = BASE / "style.qss"
    try:
        if qss_path.exists():
            return qss_path.read_text()
        else:
            logger.warning(f"QSS file not found at: {qss_path}")
    except Exception as e:
        logger.error(f"Error loading QSS: {e}")
    return ""

class SidebarButton(QtWidgets.QPushButton):
    def __init__(self, text, icon_path=None, parent=None):
        super().__init__(text, parent)
        self.setCheckable(True)
        self.setMinimumHeight(56)
        self.setMinimumWidth(60)
        self.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        if icon_path and icon_path.exists():
            # Load original pixmap
            orig_pix = QtGui.QPixmap(str(icon_path))
            # Create a white-tinted pixmap for the unchecked (Off) state
            try:
                white_pix = QtGui.QPixmap(orig_pix.size())
                white_pix.fill(QtCore.Qt.transparent)
                p = QtGui.QPainter(white_pix)
                p.setRenderHints(QtGui.QPainter.Antialiasing | QtGui.QPainter.SmoothPixmapTransform)
                p.drawPixmap(0, 0, orig_pix)
                p.setCompositionMode(QtGui.QPainter.CompositionMode_SourceIn)
                p.fillRect(white_pix.rect(), QtGui.QColor('white'))
                p.end()
            except Exception:
                # Fallback: use original if tinting fails
                white_pix = orig_pix

            icon = QtGui.QIcon()
            # Off = not checked -> white icon
            icon.addPixmap(white_pix, QtGui.QIcon.Normal, QtGui.QIcon.Off)
            # On = checked -> original (colored) icon
            icon.addPixmap(orig_pix, QtGui.QIcon.Normal, QtGui.QIcon.On)
            self.setIcon(icon)
            self.setIconSize(QtCore.QSize(28, 28))
        # Center the icon by removing text and ensuring proper alignment
        self.setText("")


import os

class DownloadThread(QtCore.QThread):
    progress = QtCore.Signal(int, float, str)  # value, speed_mbs, eta_str
    finished = QtCore.Signal(bool, str)

    def __init__(self, url, path, last_version_diagbox, total_size=0):
        super().__init__()
        self.url = url
        self.path = path
        self.last_version_diagbox = last_version_diagbox
        self.total_size = total_size
        self._is_cancelled = False
        self._is_paused = False

    def cancel(self):
        """Cancel the download"""
        self._is_cancelled = True

    def pause(self):
        """Pause the download"""
        self._is_paused = True

    def resume(self):
        """Resume the download"""
        self._is_paused = False

    def run(self):
        try:
            logger.info(f"Starting download: {self.url}")
            response = requests.get(self.url, stream=True)
            response.raise_for_status()
            
            # Try to get content length from response if not provided
            if self.total_size == 0:
                content_length = response.headers.get('content-length')
                if content_length:
                    self.total_size = int(content_length)
                    logger.info(f"File size from GET response: {self.total_size / (1024*1024):.2f} MB")
            
            downloaded = 0
            chunk_count = 0
            start_time = time.time()
            
            if self.total_size > 0:
                logger.info(f"Download initiated, total size: {self.total_size / (1024*1024):.2f} MB")
            else:
                logger.info(f"Download initiated, size unknown")
            with open(self.path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    # Check if paused
                    while self._is_paused and not self._is_cancelled:
                        time.sleep(0.1)
                    
                    if self._is_cancelled:
                        # Delete partial file
                        f.close()
                        if os.path.exists(self.path):
                            os.remove(self.path)
                        logger.warning(f"Download cancelled: Diagbox {self.last_version_diagbox}")
                        self.finished.emit(False, f"Download Diagbox {self.last_version_diagbox} cancelled")
                        return
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        chunk_count += 1
                        if chunk_count % 100 == 0:  # Flush every 100 chunks (~800KB)
                            f.flush()
                            elapsed = time.time() - start_time
                            speed = downloaded / elapsed / (1024 * 1024) if elapsed > 0 else 0  # MB/s
                            if self.total_size > 0:
                                progress = int((downloaded / self.total_size) * 1000)
                                remaining = self.total_size - downloaded
                                eta_seconds = remaining / (speed * 1024 * 1024) if speed > 0 else 0
                                eta_str = f"{int(eta_seconds // 60):02d}:{int(eta_seconds % 60):02d}"
                            else:
                                # Show downloaded MB when total size is unknown
                                progress = 0
                                eta_str = f"{downloaded / (1024 * 1024):.1f} MB"
                            self.progress.emit(progress, speed, eta_str)
            if self.total_size == 0 or downloaded >= self.total_size:
                self.progress.emit(1000, 0, "00:00")
            if os.path.exists(self.path):
                logger.info(f"Download completed successfully: Diagbox {self.last_version_diagbox}")
                self.finished.emit(True, f"Download Diagbox {self.last_version_diagbox} ok")
            else:
                logger.error(f"Download failed: File not found after download")
                self.finished.emit(False, f"Download Diagbox {self.last_version_diagbox} failed")
        except Exception as e:
            logger.error(f"Download exception: {e}", exc_info=True)
            self.finished.emit(False, f"Download Diagbox {self.last_version_diagbox} failed: {e}")

class InstallThread(QtCore.QThread):
    finished = QtCore.Signal(bool, str)
    progress = QtCore.Signal(int)  # progress percentage
    file_progress = QtCore.Signal(str)  # current file being extracted
    # Signals to inform UI about runtimes installer state
    runtimes_started = QtCore.Signal()
    runtimes_finished = QtCore.Signal(bool, str)  # success, message
    # Signal to report driver installation result when run inside the install thread
    driver_finished = QtCore.Signal(bool, str)  # success, message

    def __init__(self, path):
        super().__init__()
        self.path = path
        self.process = None  # Store process reference for cleanup

    def stop(self):
        """Stop the extraction process"""
        if self.process and self.process.poll() is None:
            try:
                self.process.terminate()
                self.process.wait(timeout=5)
            except:
                try:
                    self.process.kill()
                except:
                    pass

    def run(self):
        try:
            logger.info(f"Starting installation from: {self.path}")
            extraction_errors = []
            
            # Prefer bundled 7za, then fallback to system-installed 7za
            candidates = []
            bundled_7za = BASE / "tools" / "7za.exe"
            if bundled_7za.exists():
                candidates.append(str(bundled_7za))

            # Add system command '7za' as fallback
            try:
                if shutil.which("7za"):
                    candidates.append("7za")
            except Exception:
                pass

            if not candidates:
                logger.error("No 7za executable found (checked bundled tools and PATH)")
                self.finished.emit(False, "7za executable not found")
                return

            logger.info(f"Starting extraction using candidates: {candidates}")
            self.progress.emit(0)

            tried = []
            extraction_succeeded = False
            last_errors = []

            # Try each candidate until one succeeds
            for exe in candidates:
                tried.append(exe)
                logger.info(f"Attempting extraction with: {exe}")
                try:
                    # Build command with optional password
                    cmd = [exe, "x", self.path, "-oC:\\", "-y", "-bsp1"]
                    if ARCHIVE_PASSWORD:
                        cmd.append(f"-p{ARCHIVE_PASSWORD}")
                    
                    self.process = subprocess.Popen(
                        cmd,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        bufsize=1,
                        universal_newlines=True,
                        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
                    )

                    stdout_lines = []
                    # Read output in real-time to get progress
                    while True:
                        output = self.process.stdout.readline()
                        if output == '' and self.process.poll() is not None:
                            break
                        if output:
                            stdout_lines.append(output)
                            # Log the output for debugging
                            #logger.debug(f"7z output ({exe}): {output.strip()}")

                            # 7z outputs progress like "1% 2909 - filename"
                            line = output.strip()
                            if '%' in line and ' - ' in line:
                                try:
                                    percent_str = line.split('%')[0].strip()
                                    if percent_str.isdigit():
                                        percent = int(percent_str)
                                        logger.debug(f"Progress extracted: {percent}%")
                                        self.progress.emit(percent)

                                    filename = line.split(' - ', 1)[1] if ' - ' in line else ''
                                    if filename:
                                        self.file_progress.emit(filename)
                                except Exception as e:
                                    logger.warning(f"Error parsing progress: {e}")
                                    pass

                    # After process ends, collect stderr and decide
                    return_code = self.process.poll()
                    stderr = self.process.stderr.read() if self.process.stderr is not None else ''
                    combined_output = '\n'.join(stdout_lines) + '\n' + (stderr or '')

                    if return_code == 0 and "Can't open as archive" not in combined_output:
                        logger.info(f"Extraction succeeded with: {exe}")
                        extraction_succeeded = True
                        break
                    else:
                        logger.warning(f"Extraction failed with {exe}: return_code={return_code}, stderr={stderr[:200]}")
                        last_errors.append(f"{exe}: {stderr[:400]}")
                        # try next candidate
                        continue

                except Exception as e:
                    logger.error(f"Extraction attempt with {exe} raised exception: {e}")
                    last_errors.append(f"{exe}: {str(e)}")
                    continue

            # After trying all candidates
            if not extraction_succeeded:
                if any('permission' in (err or '').lower() for err in last_errors):
                    extraction_errors.append("Some files skipped due to permission errors")
                elif last_errors:
                    extraction_errors.append("7z extraction failed: " + ' | '.join(last_errors[:3]))
                else:
                    extraction_errors.append("Unknown extraction failure")
            
            self.progress.emit(100)

            # After extraction, attempt to run VCI driver installer BEFORE runtimes if present
            try:
                # Use DPInst for driver installation as requested
                dpinst_path = r"C:\AWRoot\extra\drivers\xsevo\amd64\DPInst.exe"
                dp_path_arg = r"C:\AWRoot\extra\drivers\xsevo\dp"
                ini_check = r"C:\Windows\System32\DriverStore\FileRepository\vcommusb.inf_amd64_0cb1ee01f7e64ab9.ini"
                if os.path.exists(dpinst_path):
                    logger.info(f"Found DPInst at: {dpinst_path}")
                    # If .ini already exists, consider driver already installed
                    if os.path.exists(ini_check):
                        msg = translator.t('messages.vci_driver.already_present')
                        logger.info("VCI .ini present before installation - already installed")
                        self.driver_finished.emit(True, msg)
                    else:
                        logger.info(f"Launching DPInst with /PATH {dp_path_arg} (interactive mode)")
                        try:
                            dp_cwd = os.path.dirname(dpinst_path)
                            drv_proc = subprocess.run(
                                [dpinst_path, '/PATH', dp_path_arg],
                                capture_output=True,
                                text=True,
                                cwd=dp_cwd
                            )
                            out = (drv_proc.stdout or '').strip()
                            err = (drv_proc.stderr or '').strip()
                            rc = drv_proc.returncode
                            logger.debug(f"DPInst returned code={rc}")
                            if out:
                                logger.debug(f"DPInst stdout: {out[:2000]}")
                            if err:
                                logger.warning(f"DPInst stderr: {err[:2000]}")
                            # If DPInst indicates success but requires reboot (e.g., return code 256), treat as success-with-reboot
                            if rc == 0:
                                if os.path.exists(ini_check):
                                    msg = translator.t('messages.vci_driver.success')
                                    self.driver_finished.emit(True, msg)
                                else:
                                    msg = translator.t('messages.vci_driver.warning', code=rc)
                                    self.driver_finished.emit(False, msg + "\n" + (err or ''))
                            elif rc == 256:
                                msg = translator.t('messages.vci_driver.success_reboot')
                                logger.info("DPInst returned reboot-required code; reporting success requiring reboot")
                                self.driver_finished.emit(True, msg)
                            else:
                                msg = translator.t('messages.vci_driver.warning', code=rc)
                                self.driver_finished.emit(False, msg + "\n" + (err or ''))
                        except Exception as e:
                            logger.error(f"DPInst exception inside InstallThread: {e}", exc_info=True)
                            self.driver_finished.emit(False, translator.t('messages.vci_driver.error', error=str(e)))
                else:
                    # DPInst not found - emit not found message
                    msg = translator.t('messages.vci_driver.not_found', path=dpinst_path)
                    self.driver_finished.emit(False, msg)
            except Exception as e:
                logger.debug(f"Driver install pre-runtimes check failed: {e}")

            # After driver install attempt, attempt to run bundled runtimes installer if present
            runtimes_path = Path(r"C:\AWRoot\Extra\runtimes\runtimes.exe")
            runtimes_warnings = []
            try:
                if runtimes_path.exists():
                    logger.info(f"Found runtimes installer at: {runtimes_path}, launching with /ai /gm2")
                    # Notify UI that runtimes installation is starting
                    try:
                        logger.info("[RUNTIMES] Emitting runtimes_started signal")
                        self.runtimes_started.emit()
                    except Exception as e:
                        logger.error(f"[RUNTIMES] Failed to emit runtimes_started: {e}")
                    # Run silently and wait for completion; creationflags hides window on Windows
                    logger.info("[RUNTIMES] About to execute subprocess.run...")
                    proc = subprocess.run(
                        [str(runtimes_path), '/ai', '/gm2'],
                        capture_output=True,
                        text=True,
                        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
                    )
                    logger.info(f"[RUNTIMES] subprocess.run completed with returncode={proc.returncode}")
                    # Log stdout/stderr for diagnostics (truncated)
                    try:
                        out = (proc.stdout or '').strip()
                        err = (proc.stderr or '').strip()
                        if out:
                            logger.info(f"Runtimes stdout: {out[:2000]}")
                        if err:
                            logger.warning(f"Runtimes stderr: {err[:2000]}")
                    except Exception as log_ex:
                        logger.warning(f"[RUNTIMES] Failed to log stdout/stderr: {log_ex}")
                    if proc.returncode == 0:
                        msg = translator.t('messages.install.runtimes.success')
                        logger.info(msg)
                        try:
                            logger.info(f"[RUNTIMES] Emitting runtimes_finished signal: success=True")
                            self.runtimes_finished.emit(True, msg)
                        except Exception as e:
                            logger.error(f"[RUNTIMES] Failed to emit runtimes_finished: {e}")
                    else:
                        msg = translator.t('messages.install.runtimes.failed', code=proc.returncode)
                        logger.warning(f"{msg}: {proc.stderr[:400]}")
                        runtimes_warnings.append(msg)
                        try:
                            # include stderr in the emitted message for UI/logging
                            logger.info(f"[RUNTIMES] Emitting runtimes_finished signal: success=False")
                            self.runtimes_finished.emit(False, msg + "\n" + (proc.stderr or ""))
                        except Exception as e:
                            logger.error(f"[RUNTIMES] Failed to emit runtimes_finished: {e}")
                else:
                    msg = translator.t('messages.install.runtimes.not_found', path=str(runtimes_path))
                    logger.debug(msg)
                    runtimes_warnings.append(msg)
                    try:
                        logger.info(f"[RUNTIMES] Emitting runtimes_finished signal: not found")
                        self.runtimes_finished.emit(False, msg)
                    except Exception as e:
                        logger.error(f"[RUNTIMES] Failed to emit runtimes_finished: {e}")
            except Exception as e:
                msg = translator.t('messages.install.runtimes.error', error=str(e))
                logger.error(msg, exc_info=True)
                runtimes_warnings.append(msg)
                try:
                    logger.info(f"[RUNTIMES] Emitting runtimes_finished signal: exception")
                    self.runtimes_finished.emit(False, msg)
                except Exception as emit_err:
                    logger.error(f"[RUNTIMES] Failed to emit runtimes_finished: {emit_err}")

            # Build result message (include any extraction or runtimes warnings)
            combined_warnings = extraction_errors[:] if extraction_errors else []
            if runtimes_warnings:
                combined_warnings.extend(runtimes_warnings)

            if combined_warnings:
                error_summary = "\n".join(combined_warnings)
                logger.warning(f"Installation completed with warnings: {error_summary}")
                message = translator.t('messages.install.warnings', warnings=error_summary)
                self.finished.emit(True, message)
            else:
                logger.info("Diagbox installed successfully to C:")
                self.finished.emit(True, translator.t('messages.install.success'))
                
        except Exception as e:
            logger.error(f"Installation failed: {e}", exc_info=True)
            self.finished.emit(False, f"Installation failed: {e}")


class CleanThread(QtCore.QThread):
    """Thread for cleaning Diagbox folders, shortcuts and driver items

    Accepts three lists: folders, shortcuts, and driver_items. The latter
    contains paths (folders or files) related to the VCI driver that should
    be deleted and reported separately in the confirmation dialog.
    """
    finished = QtCore.Signal(bool, str, int)  # success, message, success_count
    progress = QtCore.Signal(int, int)  # current, total
    item_progress = QtCore.Signal(str)  # current item being deleted

    def __init__(self, folders, shortcuts, driver_items=None):
        super().__init__()
        self.folders = folders or []
        self.shortcuts = shortcuts or []
        self.driver_items = driver_items or []
        self.failed_items = []

    def run(self):
        total_items = len(self.folders) + len(self.shortcuts) + len(self.driver_items)
        current_item = 0
        success_count = 0

        # Protect folders that may contain DPInst (we must not delete them before running DPInst)
        dpinst_path = r"C:\AWRoot\Extra\Drivers\xsevo\amd64\DPInst.exe"
        norm_dpinst = os.path.normcase(os.path.normpath(dpinst_path))
        deferred_folders = []
        active_folders = []
        for folder in self.folders:
            try:
                nf = os.path.normcase(os.path.normpath(folder))
                # If DPInst path starts with the folder path, defer deletion of that folder
                if norm_dpinst.startswith(nf):
                    deferred_folders.append(folder)
                else:
                    active_folders.append(folder)
            except Exception:
                active_folders.append(folder)

        # Delete folders that are safe to remove now (not containing DPInst)
        for folder in active_folders:
            try:
                self.item_progress.emit(translator.t('labels.deleting_folder', folder=os.path.basename(folder)))
                shutil.rmtree(folder)
                logger.info(f"Deleted folder: {folder}")
                success_count += 1
            except Exception as e:
                logger.error(f"Failed to delete folder {folder}: {e}")
                self.failed_items.append(f"{folder}: {str(e)}")

            current_item += 1
            self.progress.emit(current_item, total_items)

        # Delete shortcuts
        for shortcut in self.shortcuts:
            try:
                self.item_progress.emit(translator.t('labels.deleting_shortcut', shortcut=os.path.basename(shortcut)))
                os.remove(shortcut)
                logger.info(f"Deleted shortcut: {shortcut}")
                success_count += 1
            except Exception as e:
                logger.error(f"Failed to delete shortcut {shortcut}: {e}")
                self.failed_items.append(f"{os.path.basename(shortcut)}: {str(e)}")

            current_item += 1
            self.progress.emit(current_item, total_items)
        # Delete driver-related items (files or folders)
        # The driver must only be removed via DPInst. Do NOT delete any driver files/folders manually.
        dpinst_attempted = False
        dpinst_success = False
        try:
            dpinst_path = r"C:\AWRoot\Extra\Drivers\xsevo\amd64\DPInst.exe"
            # Use the single canonical INF path only. Do not attempt to infer from driver_items.
            inf_path = r"C:\Windows\System32\DriverStore\FileRepository\vcommusb.inf_amd64_0cb1ee01f7e64ab9\vcommusb.inf"
            if not os.path.exists(inf_path):
                # If the canonical INF is not present, do not attempt other heuristics.
                inf_path = None

            if inf_path and os.path.exists(dpinst_path):
                dpinst_attempted = True
                try:
                    self.item_progress.emit(translator.t('labels.deleting_shortcut', shortcut=os.path.basename(inf_path)))
                    logger.info(f"Attempting DPInst uninstall. DPInst path={dpinst_path}, INF={inf_path}")
                    # If the app is not running as admin, DPInst may fail silently or return non-zero
                    try:
                        elevated = is_admin()
                    except Exception:
                        elevated = False
                    if not elevated:
                        logger.warning("Current process is not elevated. DPInst may require admin rights to uninstall the driver.")

                    # Run DPInst from its own directory (some installers expect local working directory)
                    dpinst_cwd = os.path.dirname(dpinst_path)
                    dp_proc = subprocess.run(
                        [dpinst_path, '/U', inf_path, '/S'],
                        capture_output=True,
                        text=True,
                        cwd=dpinst_cwd,
                        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
                    )

                    rc = dp_proc.returncode
                    out = (dp_proc.stdout or '').strip()
                    err = (dp_proc.stderr or '').strip()
                    logger.debug(f"DPInst returncode={rc}")
                    if out:
                        logger.debug(f"DPInst stdout: {out[:4000]}")
                    if err:
                        logger.warning(f"DPInst stderr: {err[:4000]}")

                    if rc == 0:
                        dpinst_success = True
                        logger.info("DPInst reported success; driver uninstall requested")
                    else:
                        logger.warning(f"DPInst returned non-zero code {rc}")
                        self.failed_items.append(f"DPInst: returncode={rc} stdout={out[:1000]} stderr={err[:1000]}")
                except Exception as e:
                    logger.error(f"Failed to run DPInst: {e}", exc_info=True)
                    self.failed_items.append(f"DPInst exception: {str(e)}")
        except Exception:
            pass

        if self.driver_items:
            if dpinst_attempted and dpinst_success:
                for item in self.driver_items:
                    try:
                        name = os.path.basename(item)
                        # Report progress but do not delete files manually
                        self.item_progress.emit(translator.t('labels.deleting_folder', folder=name) if os.path.isdir(item) else translator.t('labels.deleting_shortcut', shortcut=name))
                        logger.info(f"Driver removal handled by DPInst: {item}")
                        success_count += 1
                    except Exception as e:
                        logger.error(f"Error while marking driver item processed {item}: {e}")
                        self.failed_items.append(f"{item}: {str(e)}")
                    current_item += 1
                    self.progress.emit(current_item, total_items)
            else:
                # DPInst was not run or failed — flag all driver items as failed and do not delete them
                for item in self.driver_items:
                    name = os.path.basename(item)
                    logger.warning(f"DPInst not successful; skipping manual deletion for: {item}")
                    self.failed_items.append(f"DPInst not run or failed: {name}")
                    current_item += 1
                    self.progress.emit(current_item, total_items)

        # If DPInst succeeded, now delete any deferred folders (those containing DPInst), because
        # DPInst executable may be inside them and needed to uninstall the driver.
        if deferred_folders:
            if dpinst_success:
                for folder in deferred_folders:
                    try:
                        self.item_progress.emit(translator.t('labels.deleting_folder', folder=os.path.basename(folder)))
                        shutil.rmtree(folder)
                        logger.info(f"Deleted deferred folder: {folder}")
                        success_count += 1
                    except Exception as e:
                        logger.error(f"Failed to delete deferred folder {folder}: {e}")
                        self.failed_items.append(f"{folder}: {str(e)}")

                    current_item += 1
                    self.progress.emit(current_item, total_items)
            else:
                for folder in deferred_folders:
                    logger.warning(f"DPInst not successful; skipping deletion of deferred folder: {folder}")
                    self.failed_items.append(f"DPInst not run or failed: {os.path.basename(folder)}")
                    current_item += 1
                    self.progress.emit(current_item, total_items)

            current_item += 1
            self.progress.emit(current_item, total_items)

        # Build result message
        if self.failed_items:
            error_list = "\n".join(self.failed_items)
            message = translator.t('messages.clean.partial', count=success_count, errors=error_list)
            self.finished.emit(False, message, success_count)
        else:
            message = translator.t('messages.clean.success', count=success_count)
            self.finished.emit(True, message, success_count)


class MainWindow(QtWidgets.QWidget):
    
    
    download_finished = QtCore.Signal(bool, str)  # success, message
    manual_runtimes_finished = QtCore.Signal(bool, str)  # success, message - for manual button

    def __init__(self):
        super().__init__()
        logger.info("Initializing MainWindow")
        
        self.setWindowTitle(translator.t('app.title'))
        self.setWindowFlags(QtCore.Qt.WindowType.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground)
        self.resize(900, 500)

        # Download variables
        self.download_folder = "C:\\INSTALL\\"
        self.last_version_diagbox = ""
        self.diagbox_path = ""
        self.auto_install = None
        self.download_thread = None
        self.install_thread = None
        self.cancel_button = None
        self.pause_button = None
        self.dragPos = QtCore.QPoint()
        self.log_widget = None
        self.log_handler = None
        
        # Fetch last version first
        logger.info("[STEP 1] -- Fetching last Diagbox version...")
        self.fetch_last_version_diagbox()
        
        # Version options: load from remote JSON (configured in `config.URL_VERSION_OPTIONS`)
        # Falls back to the built-in defaults if remote fetch fails.
        self.version_options = self.load_version_options()
        # Remote messages/banners (for homepage notifications)
        self.remote_messages = []
        try:
            # initial load and periodic refresh
            self.load_remote_messages()
            self.message_timer = QtCore.QTimer(self)
            self.message_timer.timeout.connect(self.load_remote_messages)
            self.message_timer.start(60 * 1000)  # refresh every 60s
        except Exception:
            pass

        # Connect signals
        self.download_finished.connect(self.on_download_finished)
        self.manual_runtimes_finished.connect(self._on_manual_runtimes_finished)

        self.setup_ui()
        
        # Check for app updates after UI is ready
        QtCore.QTimer.singleShot(1000, self.check_app_update)

    def load_version_options(self):
        """Load version options from remote JSON configured in `config.URL_VERSION_OPTIONS`.

        Expected JSON format: an array of objects or arrays, e.g.:
        [
          {"display_name": "Diagbox 09.180 (2024)", "version": "09.180_PSA_DIAG", "url": "https://..."},
          ["Diagbox 09.85", "09.85", "https://..."]
        ]

        Returns a list of tuples: (display_name, version, url)
        """
        defaults = [
            ("ERROR", "00.00", "ERROR")
        ]

        try:
            logger.info(f"Loading version options from: {URL_VERSION_OPTIONS}")
            resp = requests.get(URL_VERSION_OPTIONS, timeout=6)
            resp.raise_for_status()
            data = resp.json()
            options = []
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, dict):
                        display = item.get('display_name') or item.get('display') or item.get('name')
                        version = item.get('version')
                        url = item.get('url')
                        if display and version and url:
                            options.append((display, version, url))
                    elif isinstance(item, (list, tuple)) and len(item) >= 3:
                        options.append((item[0], item[1], item[2]))

            if options:
                logger.info(f"Loaded {len(options)} version options from remote JSON")
                return options
            else:
                logger.warning("Version options JSON did not contain valid entries, using defaults")
                return defaults
        except Exception as e:
            logger.warning(f"Failed to load version options from remote JSON: {e}")
            return defaults

    def load_remote_messages(self):
        """Load remote messages/banners JSON and update the homepage banner.

        Expected JSON: list of objects like:
        {
          "id": "promo1",
          "lang": {"en": {"text":"Hello","link":"https://..."}, "fr": {...}},
          "start": "2025-12-01T00:00:00Z",
          "end": "2025-12-31T23:59:59Z",
          "priority": 10
        }
        """
        try:
            # logger.info(f"Loading remote messages from: {URL_REMOTE_MESSAGES}")
            r = requests.get(URL_REMOTE_MESSAGES, timeout=6)
            r.raise_for_status()
            data = r.json()
            messages = []
            if isinstance(data, dict):
                # allow single-object root
                data = [data]
            if isinstance(data, list):
                for item in data:
                    try:
                        mid = item.get('id') or item.get('name')
                        langmap = item.get('lang') or item.get('texts') or {}
                        start = item.get('start')
                        end = item.get('end')
                        priority = int(item.get('priority', 0))
                        messages.append({'id': mid, 'lang': langmap, 'start': start, 'end': end, 'priority': priority, 'raw': item})
                    except Exception:
                        continue
            # keep in memory and update UI
            self.remote_messages = sorted(messages, key=lambda x: x.get('priority', 0), reverse=True)
            QtCore.QTimer.singleShot(50, self.update_global_banner)
        except Exception as e:
            logger.debug(f"Failed to load remote messages: {e}")

    def update_global_banner(self):
        """Create or update a homepage banner from `self.remote_messages`.

        Picks the highest-priority active message and displays it on the
        home/config page (stack index 0). Supports `lang` map with per-
        language text and optional `link`/`link_text`.
        """
        try:
            if not getattr(self, 'remote_messages', None):
                # hide existing banner if nothing to show
                if hasattr(self, 'home_banner'):
                    try:
                        self.home_banner.setVisible(False)
                    except Exception:
                        pass
                return

            # collect all messages that apply to the home page and are active
            now = QtCore.QDateTime.currentDateTimeUtc()
            candidates = []
            for msg in self.remote_messages:
                raw = msg.get('raw', {})
                display_on = raw.get('display_on') or raw.get('display') or []
                if display_on and 'home' not in display_on:
                    continue
                # time window
                start = raw.get('start')
                end = raw.get('end')
                if start:
                    try:
                        st = QtCore.QDateTime.fromString(start, QtCore.Qt.ISODate)
                        if st.isValid() and st > now:
                            continue
                    except Exception:
                        pass
                if end:
                    try:
                        et = QtCore.QDateTime.fromString(end, QtCore.Qt.ISODate)
                        if et.isValid() and et < now:
                            continue
                    except Exception:
                        pass
                candidates.append(msg)

            if not candidates:
                if hasattr(self, 'home_banner'):
                    try:
                        self.home_banner.setVisible(False)
                    except Exception:
                        pass
                return

            # Save candidates (ordered by priority already) and start rotation
            self.home_messages = candidates
            self.home_banner_index = 0

            # ensure banner exists / create if needed
            home_widget = self.stack.widget(0)
            if home_widget is None:
                return
            layout = home_widget.layout()
            if not layout:
                return

            if not hasattr(self, 'home_banner'):
                self.home_banner = QtWidgets.QFrame()
                self.home_banner.setObjectName('installBanner')
                main_banner_layout = QtWidgets.QVBoxLayout(self.home_banner)
                main_banner_layout.setContentsMargins(10,10,10,10)
                main_banner_layout.setSpacing(8)
                
                # Top row: arrows + message content
                bann_layout = QtWidgets.QHBoxLayout()
                bann_layout.setSpacing(8)
                
                # Left arrow in vertical container
                left_arrow_container = QtWidgets.QVBoxLayout()
                left_arrow_container.addStretch()
                left_arrow = QtWidgets.QPushButton('\u276E')
                left_arrow.setObjectName('bannerArrow')
                left_arrow.setFixedSize(32, 32)
                left_arrow.clicked.connect(self._prev_home_banner)
                left_arrow_container.addWidget(left_arrow)
                left_arrow_container.addStretch()
                bann_layout.addLayout(left_arrow_container)
                self.home_banner_left_arrow = left_arrow
                
                # Center content (message + link + dots)
                center_layout = QtWidgets.QVBoxLayout()
                center_layout.setSpacing(8)
                
                lbl = QtWidgets.QLabel("")
                lbl.setObjectName('installBannerLabel')
                lbl.setWordWrap(True)
                lbl.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                center_layout.addWidget(lbl)
                
                # Link button row (centered)
                link_layout = QtWidgets.QHBoxLayout()
                link_layout.addStretch()
                self.home_banner_link_layout = link_layout
                link_layout.addStretch()
                center_layout.addLayout(link_layout)
                
                # Pagination dots row (centered)
                dots_layout = QtWidgets.QHBoxLayout()
                dots_layout.setSpacing(6)
                dots_layout.addStretch()
                self.home_banner_dots_layout = dots_layout
                self.home_banner_dots = []
                dots_layout.addStretch()
                center_layout.addLayout(dots_layout)
                
                bann_layout.addLayout(center_layout, 1)
                
                # placeholder for link
                self.home_banner_link = None
                
                # Right arrow in vertical container
                right_arrow_container = QtWidgets.QVBoxLayout()
                right_arrow_container.addStretch()
                right_arrow = QtWidgets.QPushButton('\u276F')
                right_arrow.setObjectName('bannerArrow')
                right_arrow.setFixedSize(32, 32)
                right_arrow.clicked.connect(self._next_home_banner)
                right_arrow_container.addWidget(right_arrow)
                right_arrow_container.addStretch()
                bann_layout.addLayout(right_arrow_container)
                self.home_banner_right_arrow = right_arrow
                
                main_banner_layout.addLayout(bann_layout)
                
                # Use opacity effect instead of drop shadow to allow fade animations
                try:
                    opacity_effect = QtWidgets.QGraphicsOpacityEffect(self.home_banner)
                    opacity_effect.setOpacity(1.0)
                    self.home_banner.setGraphicsEffect(opacity_effect)
                    self._banner_opacity_effect = opacity_effect
                except Exception:
                    pass
                layout.insertWidget(0, self.home_banner)
                self.home_banner_label = lbl

            # start a rotation timer if multiple messages
            try:
                if hasattr(self, 'home_banner_timer') and self.home_banner_timer.isActive():
                    self.home_banner_timer.stop()
                else:
                    self.home_banner_timer = QtCore.QTimer(self)
                    self.home_banner_timer.timeout.connect(self._advance_home_banner)
                if len(self.home_messages) > 1:
                    self.home_banner_timer.start(8000)
                else:
                    try:
                        self.home_banner_timer.stop()
                    except Exception:
                        pass
            except Exception:
                pass

            # Update pagination dots
            try:
                self._update_banner_dots()
            except Exception:
                pass
            
            # show first message now
            try:
                self._show_home_message(self.home_banner_index)
            except Exception as e:
                logger.debug(f"show_home_message error: {e}")
        except Exception as e:
            logger.debug(f"update_global_banner error: {e}")

    def _show_home_message(self, index):
        """Display message at `index` from `self.home_messages` with slide transition."""
        try:
            if not getattr(self, 'home_messages', None):
                return
            msg = self.home_messages[index]
            raw = msg.get('raw', {})
            langmap = msg.get('lang', {})
            lang_code = translator.language if hasattr(translator, 'language') else 'en'
            text_entry = langmap.get(lang_code) or langmap.get('en') or {}
            text = text_entry.get('text') if isinstance(text_entry, dict) else text_entry or str(text_entry)
            link = text_entry.get('link') if isinstance(text_entry, dict) else None
            link_text = text_entry.get('link_text') if isinstance(text_entry, dict) else None

            # Stop any running animation
            if hasattr(self, '_banner_animation') and self._banner_animation:
                try:
                    self._banner_animation.stop()
                except Exception:
                    pass

            # Create fade out effect first (if content exists)
            if hasattr(self, 'home_banner_label') and self.home_banner_label.text():
                try:
                    # Create or reuse opacity effect
                    if not hasattr(self, '_banner_opacity_effect'):
                        self._banner_opacity_effect = QtWidgets.QGraphicsOpacityEffect(self.home_banner)
                        self.home_banner.setGraphicsEffect(self._banner_opacity_effect)
                    
                    opacity_effect = self._banner_opacity_effect
                    
                    # Fade out animation
                    fade_out = QtCore.QPropertyAnimation(opacity_effect, b"opacity")
                    fade_out.setDuration(250)
                    fade_out.setStartValue(1.0)
                    fade_out.setEndValue(0.0)
                    fade_out.setEasingCurve(QtCore.QEasingCurve.Type.InOutQuad)
                    
                    # Update content at end of fade out
                    def update_content():
                        try:
                            self.home_banner_label.setText(text or '')
                            # Update link button
                            if link:
                                if not getattr(self, 'home_banner_link', None):
                                    btn = QtWidgets.QPushButton(link_text or translator.t('buttons.download'))
                                    btn.setObjectName('bannerDownload')
                                    btn._url = link
                                    btn.clicked.connect(self._open_button_url)
                                    # Insert in the centered link layout (between the two stretches)
                                    self.home_banner_link_layout.insertWidget(1, btn)
                                    self.home_banner_link = btn
                                else:
                                    self.home_banner_link._url = link
                                    self.home_banner_link.setText(link_text or translator.t('buttons.download'))
                                    self.home_banner_link.setVisible(True)
                            else:
                                if getattr(self, 'home_banner_link', None):
                                    self.home_banner_link.setVisible(False)
                            
                            # Fade back in
                            fade_in = QtCore.QPropertyAnimation(opacity_effect, b"opacity")
                            fade_in.setDuration(250)
                            fade_in.setStartValue(0.0)
                            fade_in.setEndValue(1.0)
                            fade_in.setEasingCurve(QtCore.QEasingCurve.Type.InOutQuad)
                            fade_in.start()
                            self._banner_animation = fade_in
                        except Exception as e:
                            logger.debug(f"Fade in error: {e}")
                    
                    fade_out.finished.connect(update_content)
                    fade_out.start()
                    self._banner_animation = fade_out
                except Exception as e:
                    logger.debug(f"Animation error: {e}")
                    # Fallback to instant update
                    self._update_banner_content_instant(text, link, link_text)
            else:
                # First message: no animation needed
                self._update_banner_content_instant(text, link, link_text)

            try:
                self.home_banner.setVisible(True)
            except Exception:
                pass
        except Exception as e:
            logger.debug(f"_show_home_message error: {e}")

    def _update_banner_content_instant(self, text, link, link_text):
        """Update banner content without animation (fallback)."""
        try:
            self.home_banner_label.setText(text or '')
            if link:
                if not getattr(self, 'home_banner_link', None):
                    btn = QtWidgets.QPushButton(link_text or translator.t('buttons.download'))
                    btn.setObjectName('bannerDownload')
                    btn._url = link
                    btn.clicked.connect(self._open_button_url)
                    # Insert in the centered link layout (between the two stretches)
                    self.home_banner_link_layout.insertWidget(1, btn)
                    self.home_banner_link = btn
                else:
                    self.home_banner_link._url = link
                    self.home_banner_link.setText(link_text or translator.t('buttons.download'))
                    self.home_banner_link.setVisible(True)
            else:
                if getattr(self, 'home_banner_link', None):
                    self.home_banner_link.setVisible(False)
        except Exception:
            pass

    def _advance_home_banner(self):
        try:
            if not getattr(self, 'home_messages', None):
                return
            self.home_banner_index = (self.home_banner_index + 1) % len(self.home_messages)
            self._show_home_message(self.home_banner_index)
            self._update_banner_dots()
        except Exception as e:
            logger.debug(f"_advance_home_banner error: {e}")
    
    def _next_home_banner(self):
        """Navigate to next banner message manually."""
        try:
            if not getattr(self, 'home_messages', None):
                return
            # Stop auto rotation timer when user manually navigates
            if hasattr(self, 'home_banner_timer') and self.home_banner_timer.isActive():
                self.home_banner_timer.stop()
            self.home_banner_index = (self.home_banner_index + 1) % len(self.home_messages)
            self._show_home_message(self.home_banner_index)
            self._update_banner_dots()
            # Restart timer
            if len(self.home_messages) > 1:
                self.home_banner_timer.start(8000)
        except Exception as e:
            logger.debug(f"_next_home_banner error: {e}")
    
    def _prev_home_banner(self):
        """Navigate to previous banner message manually."""
        try:
            if not getattr(self, 'home_messages', None):
                return
            # Stop auto rotation timer when user manually navigates
            if hasattr(self, 'home_banner_timer') and self.home_banner_timer.isActive():
                self.home_banner_timer.stop()
            self.home_banner_index = (self.home_banner_index - 1) % len(self.home_messages)
            self._show_home_message(self.home_banner_index)
            self._update_banner_dots()
            # Restart timer
            if len(self.home_messages) > 1:
                self.home_banner_timer.start(8000)
        except Exception as e:
            logger.debug(f"_prev_home_banner error: {e}")
    
    def _update_banner_dots(self):
        """Update pagination dots to reflect current message index."""
        try:
            if not hasattr(self, 'home_banner_dots_layout'):
                return
            num_messages = len(getattr(self, 'home_messages', []))
            if num_messages <= 1:
                # Hide all dots if only one message
                for dot in self.home_banner_dots:
                    dot.setVisible(False)
                # Hide arrows too
                if hasattr(self, 'home_banner_left_arrow'):
                    self.home_banner_left_arrow.setVisible(False)
                if hasattr(self, 'home_banner_right_arrow'):
                    self.home_banner_right_arrow.setVisible(False)
                return
            
            # Show arrows
            if hasattr(self, 'home_banner_left_arrow'):
                self.home_banner_left_arrow.setVisible(True)
            if hasattr(self, 'home_banner_right_arrow'):
                self.home_banner_right_arrow.setVisible(True)
            
            # Create or update dots
            current_index = getattr(self, 'home_banner_index', 0)
            
            # Remove extra dots if we have too many
            while len(self.home_banner_dots) > num_messages:
                dot = self.home_banner_dots.pop()
                self.home_banner_dots_layout.removeWidget(dot)
                dot.deleteLater()
            
            # Add missing dots
            while len(self.home_banner_dots) < num_messages:
                dot = QtWidgets.QLabel('\u25CF')
                dot.setObjectName('bannerDot')
                dot.setFixedSize(10, 10)
                dot.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                dot.setStyleSheet('font-size: 10px; color: #888;')
                # Insert before the second stretch (which is at index len(self.home_banner_dots) + 1)
                self.home_banner_dots_layout.insertWidget(len(self.home_banner_dots) + 1, dot)
                self.home_banner_dots.append(dot)
            
            # Update dot styles (active vs inactive)
            for i, dot in enumerate(self.home_banner_dots):
                dot.setVisible(True)
                if i == current_index:
                    dot.setStyleSheet('font-size: 12px; color: #fff; font-weight: bold;')
                else:
                    dot.setStyleSheet('font-size: 10px; color: #888;')
        except Exception as e:
            logger.debug(f"_update_banner_dots error: {e}")

    def _open_button_url(self):
        try:
            sender = self.sender()
            url = getattr(sender, '_url', None)
            if not url:
                return
            if sys.platform == 'win32':
                os.startfile(url)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', url])
            else:
                subprocess.Popen(['xdg-open', url])
        except Exception as e:
            logger.debug(f"_open_button_url error: {e}")

    def update_progress(self, value, speed, eta):
        # Update footer progress bar
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setValue(value)
            self.footer_progress.setFormat(f"{value / 10:.1f}% - {speed:.1f} MB/s - {eta}")
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(translator.t('labels.downloading'))
        
        QtWidgets.QApplication.processEvents()

    def on_download_finished(self, success, message):
        logger.info(f"Download finished: success={success}, message={message}")
        
        # Hide cancel and pause buttons
        if self.cancel_button:
            self.cancel_button.setVisible(False)
        if self.pause_button:
            self.pause_button.setVisible(False)
            self.pause_button.setText(translator.t('buttons.pause'))
        
        # Update footer
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setRange(0, 1000)
            self.footer_progress.setValue(1000 if success else 0)
            self.footer_progress.setFormat(translator.t('messages.download.complete') if success else translator.t('messages.download.failed_format'))
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(translator.t('labels.download_complete') if success else translator.t('labels.download_failed'))
        
        # Check if auto-install is enabled
        if success and self.auto_install and self.auto_install.isChecked():
            logger.info("Auto-install enabled, starting installation...")
            # Don't show download completion message, go directly to install
            QtCore.QTimer.singleShot(500, self.install_diagbox)
        else:
            # Re-enable all buttons and combo box only if not auto-installing
            self.set_buttons_enabled(True)
            QtWidgets.QMessageBox.information(self, translator.t('messages.download.title'), message)
            # Reset footer after message
            if hasattr(self, 'footer_label'):
                self.footer_label.setText(translator.t('labels.ready'))
            if hasattr(self, 'footer_progress'):
                self.footer_progress.setValue(0)
                self.footer_progress.setFormat("")

    def check_installed_version(self):
        version_file = r"C:\AWRoot\bin\fi\Version.ini"
        if os.path.exists(version_file):
            try:
                with open(version_file, 'r') as f:
                    content = f.read()
                    for line in content.splitlines():
                        if line.startswith("Version="):
                            version = line.split("=", 1)[1]
                            return version
            except:
                pass
        return None

    def get_diagbox_language(self):
        """Get current Diagbox language"""
        lang_file = r"C:\AWRoot\dtrd\Trans\Language.ini"
        if os.path.exists(lang_file):
            try:
                with open(lang_file, 'r') as f:
                    for line in f:
                        if '=' in line:
                            lang = line.split('=', 1)[1].strip()
                            logger.info(f"Detected Diagbox language: {lang}")
                            return lang
            except Exception as e:
                logger.error(f"Error reading language file: {e}")
        else:
            logger.info(f"Diagbox language file not found: {lang_file}")

        logger.info("No Diagbox language detected")
        return None

    def change_diagbox_language(self, new_lang_code):
        """Change Diagbox language"""
        lang_file = r"C:\AWRoot\dtrd\Trans\Language.ini"
        
        if not os.path.exists(lang_file):
            QtWidgets.QMessageBox.warning(
                self,
                translator.t('messages.language.title'),
                translator.t('messages.language.not_found', path=lang_file)
            )
            return
        
        try:
            logger.info(f"Changing Diagbox language to: {new_lang_code}")
            
            # Read current content
            with open(lang_file, 'r') as f:
                content = f.read()
            
            # Replace language
            lines = content.splitlines()
            new_lines = []
            for line in lines:
                if '=' in line:
                    key = line.split('=')[0]
                    new_lines.append(f"{key}={new_lang_code}")
                else:
                    new_lines.append(line)
            
            # Write back
            with open(lang_file, 'w') as f:
                f.write('\n'.join(new_lines))
            
            logger.info(f"Language changed successfully to {new_lang_code}")
            QtWidgets.QMessageBox.information(
                self,
                translator.t('messages.language.title'),
                translator.t('messages.language.success', lang=new_lang_code)
            )
        except Exception as e:
            logger.error(f"Failed to change language: {e}", exc_info=True)
            QtWidgets.QMessageBox.critical(
                self,
                translator.t('messages.language.title'),
                translator.t('messages.language.failed', error=str(e))
            )

    def check_downloaded_versions(self):
        """Check what versions are available in the download folder"""
        downloaded_versions = []
        if os.path.exists(self.download_folder):
            for file in os.listdir(self.download_folder):
                if file.endswith(".7z"):
                    version = None
                    # Support both naming formats:
                    # 1. "Diagbox_Install_09.180.7z" (old format)
                    # 2. "09.180.7z" (new format)
                    if file.startswith("Diagbox_Install_"):
                        version = file.replace("Diagbox_Install_", "").replace(".7z", "")
                    else:
                        # Assume filename is just version.7z
                        version = file.replace(".7z", "")
                    
                    if version:
                        file_path = os.path.join(self.download_folder, file)
                        file_size = os.path.getsize(file_path)
                        downloaded_versions.append({
                            'version': version,
                            'path': file_path,
                            'size': file_size,
                            'size_mb': file_size / (1024 * 1024)
                        })
        return downloaded_versions

    def set_buttons_enabled(self, enabled):
        """Enable or disable all buttons and combo box in the install page"""
        # Disable/enable all action buttons (except cancel and pause buttons)
        for child in self.stack.currentWidget().findChildren(QtWidgets.QPushButton):
            if child != self.cancel_button and child != self.pause_button:
                child.setEnabled(enabled)

        # Disable/enable combo box
        if hasattr(self, 'version_combo'):
            self.version_combo.setEnabled(enabled)

        # Disable/enable Diagbox language selector as well (avoid changing Diagbox language during install)
        if hasattr(self, 'language_combo'):
            self.language_combo.setEnabled(enabled)

        # Also ensure the standalone 'Install runtimes' button (if present) is toggled
        if hasattr(self, 'runtimes_btn') and self.runtimes_btn is not None:
            try:
                self.runtimes_btn.setEnabled(enabled)
            except Exception:
                pass

        # Disable/enable checkbox
        if self.auto_install:
            self.auto_install.setEnabled(enabled)

    def cancel_download(self):
        """Cancel the current download"""
        if self.download_thread and self.download_thread.isRunning():
            self.download_thread.cancel()

    def toggle_pause_download(self):
        """Pause or resume the current download"""
        if self.download_thread and self.download_thread.isRunning():
            if self.download_thread._is_paused:
                self.download_thread.resume()
                if self.pause_button:
                    self.pause_button.setText(translator.t('buttons.pause'))
            else:
                self.download_thread.pause()
                if self.pause_button:
                    self.pause_button.setText(translator.t('buttons.resume'))

    def on_install_finished(self, success, message, install_button, bar):
        # Ensure runtimes UI state is reset (re-enable runtimes button + footer)
        try:
            QtCore.QTimer.singleShot(0, lambda: self._set_runtimes_ui_running(False, message if message else translator.t('labels.ready')))
        except Exception:
            pass
        
        # Update footer
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setRange(0, 100)
            self.footer_progress.setValue(100 if success else 0)
            self.footer_progress.setFormat(translator.t('messages.install.complete') if success else translator.t('messages.install.failed_status'))
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(translator.t('labels.installation_complete') if success else translator.t('labels.installation_failed'))
        
        # Refresh install page if installation was successful
        if success:
            self.refresh_install_page()
            # Update the install banner state now that installation may have changed
            try:
                QtCore.QTimer.singleShot(100, self.on_enter_install_page)
            except Exception:
                try:
                    self.on_enter_install_page()
                except Exception:
                    pass
        
        # Build and show installation summary (Diagbox, Runtimes, Driver)
        try:
            # Diagbox result
            diag_success = bool(success)
            diag_msg = message or ""

            # Runtimes result (may be missing)
            runtimes_success = None
            runtimes_msg = None
            if hasattr(self, '_runtimes_install_result') and self._runtimes_install_result is not None:
                runtimes_success, runtimes_msg = self._runtimes_install_result

            # Driver result (may be missing)
            driver_success = None
            driver_msg = None
            if hasattr(self, '_driver_install_result') and self._driver_install_result is not None:
                driver_success, driver_msg = self._driver_install_result

            lines = []
            # Diagbox line
            diag_label = translator.t('messages.install.summary.diagbox')
            if diag_success:
                diag_status = translator.t('messages.install.summary.ok')
            else:
                diag_status = translator.t('messages.install.summary.error', msg=diag_msg)
            lines.append(f"-- {diag_label} : {diag_status}")

            # Runtimes line
            runt_label = translator.t('messages.install.summary.runtimes')
            if runtimes_success is None:
                runt_status = translator.t('messages.install.summary.not_run')
            else:
                if runtimes_success:
                    runt_status = translator.t('messages.install.summary.ok')
                else:
                    runt_status = translator.t('messages.install.summary.error', msg=(runtimes_msg or ''))
            lines.append(f"-- {runt_label} : {runt_status}")

            # Driver line
            drv_label = translator.t('messages.install.summary.driver')
            if driver_success is None:
                drv_status = translator.t('messages.install.summary.not_run')
            else:
                # If driver returned already_present message, show that text
                already_txt = translator.t('messages.vci_driver.already_present')
                if driver_success and driver_msg and driver_msg.strip() == already_txt:
                    drv_status = already_txt
                else:
                    if driver_success:
                        drv_status = translator.t('messages.install.summary.ok')
                    else:
                        drv_status = translator.t('messages.install.summary.error', msg=(driver_msg or ''))
            lines.append(f"-- {drv_label} : {drv_status}")

            summary = "\n".join(lines)
        except Exception as e:
            logger.error(f"Failed to build install summary: {e}", exc_info=True)
            summary = message or translator.t('messages.install.failed')

        # Show summary
        QtWidgets.QMessageBox.information(self, translator.t('messages.install.title'), summary)
        
        # Re-enable all buttons and combo box AFTER showing summary
        self.set_buttons_enabled(True)
        
        # Reset footer after a delay
        QtCore.QTimer.singleShot(3000, self.reset_footer)
    
    def reset_footer(self):
        """Reset footer to ready state"""
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(translator.t('labels.ready'))
        if hasattr(self, 'footer_progress'):
            # Reset range to determinate default and clear value/format
            try:
                self.footer_progress.setRange(0, 1000)
            except Exception:
                pass
            self.footer_progress.setValue(0)
            self.footer_progress.setFormat("")
    
    def refresh_install_page(self):
        """Refresh the install page to update version information"""
        # Update installed version label
        if hasattr(self, 'header_installed'):
            installed_version = self.check_installed_version()
            version_text = installed_version if installed_version else translator.t('labels.not_installed')
            self.header_installed.setText(translator.t('labels.installed_version', version=version_text))
        
        # Update downloaded versions label
        downloaded_versions = self.check_downloaded_versions()
        if downloaded_versions:
            downloaded_text = ", ".join([f"{v['version']} ({v['size_mb']:.1f} MB)" for v in downloaded_versions])
            if hasattr(self, 'header_downloaded') and self.header_downloaded:
                self.header_downloaded.setText(translator.t('labels.downloaded_versions', versions=downloaded_text))
                self.header_downloaded.setVisible(True)
            else:
                # Create downloaded label if it doesn't exist
                install_page = self.stack.widget(0)  # Install page is first (index 0)
                if install_page:
                    layout = install_page.layout()
                    self.header_downloaded = QtWidgets.QLabel(translator.t('labels.downloaded_versions', versions=downloaded_text))
                    self.header_downloaded.setObjectName("sectionHeader")
                    self.header_downloaded.setStyleSheet("color: #5cb85c;")
                    layout.insertWidget(2, self.header_downloaded)  # Insert after installed and online version
        else:
            # Hide downloaded label if no files exist
            if hasattr(self, 'header_downloaded') and self.header_downloaded:
                self.header_downloaded.setVisible(False)

            # Update install button state based on installed version
            try:
                installed_version = self.check_installed_version()
                if hasattr(self, 'install_button'):
                    if installed_version:
                        self.install_button.setEnabled(False)
                        self.install_button.setToolTip(translator.t('messages.install.must_clean_tooltip', installed=installed_version))
                    else:
                        self.install_button.setEnabled(True)
                        self.install_button.setToolTip("")
            except Exception:
                pass

    def update_install_progress(self, value):
        """Update installation progress bar"""
        # Update footer progress bar
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setRange(0, 100)
            self.footer_progress.setValue(value)
            self.footer_progress.setFormat(f"Extracting... {value}%")
        if hasattr(self, 'footer_label'):
            self.footer_label.setText("Installing Diagbox...")
        
        QtWidgets.QApplication.processEvents()
    
    def update_install_file(self, filename):
        """Update current file being extracted"""
        # Update footer label with truncated filename
        if hasattr(self, 'footer_label'):
            display_name = filename if len(filename) <= 60 else "..." + filename[-57:]
            self.footer_label.setText(f"Installing: {display_name}")
        
        QtWidgets.QApplication.processEvents()

    def _set_runtimes_ui_running(self, running: bool, message: str = None):
        """Enable/disable runtimes button and update footer with message.
        Should be called from the GUI thread (use QTimer.singleShot when called from worker threads)."""
        try:
            logger.info(f"[RUNTIMES UI] Setting running={running}, message={message[:100] if message else 'None'}")
            if hasattr(self, 'runtimes_btn') and self.runtimes_btn is not None:
                self.runtimes_btn.setEnabled(not running)
                logger.info(f"[RUNTIMES UI] Button enabled={not running}")
            else:
                logger.warning("[RUNTIMES UI] runtimes_btn not found or is None")
            if running:
                if hasattr(self, 'footer_label'):
                    self.footer_label.setText(message or translator.t('messages.install.runtimes.started'))
                # Optionally show an indeterminate progress while runtimes install runs
                if hasattr(self, 'footer_progress'):
                    self.footer_progress.setRange(0, 0)
                    logger.info("[RUNTIMES UI] Progress set to indeterminate (0, 0)")
            else:
                if hasattr(self, 'footer_label'):
                    # If a message was provided, display it briefly, otherwise reset
                    if message:
                        self.footer_label.setText(message)
                    else:
                        self.footer_label.setText(translator.t('labels.ready'))
                if hasattr(self, 'footer_progress'):
                    # Reset progress bar
                    self.footer_progress.setRange(0, 1000)
                    self.footer_progress.setValue(0)
                    self.footer_progress.setFormat("")
                    logger.info("[RUNTIMES UI] Progress reset to determinate (0, 1000)")
        except Exception as e:
            logger.error(f"_set_runtimes_ui_running error: {e}", exc_info=True)

    def _on_runtimes_started_from_installthread(self):
        try:
            logger.info("[RUNTIMES] Started callback invoked")
            msg = translator.t('messages.install.runtimes.started')
            QtCore.QTimer.singleShot(0, lambda m=msg: self._set_runtimes_ui_running(True, m))
        except Exception as e:
            logger.error(f"_on_runtimes_started_from_installthread error: {e}", exc_info=True)

    def _on_runtimes_finished_from_installthread(self, success: bool, message: str):
        try:
            logger.info(f"[RUNTIMES] Finished callback invoked: success={success}, message={message[:100]}")
            # Store runtimes result for summary
            try:
                self._runtimes_install_result = (success, message)
            except Exception:
                self._runtimes_install_result = (success, message)
            # Capture message immediately to avoid lambda late-binding issues
            msg = str(message) if message else translator.t('labels.ready')
            QtCore.QTimer.singleShot(0, lambda m=msg: self._set_runtimes_ui_running(False, m))
        except Exception as e:
            logger.error(f"_on_runtimes_finished_from_installthread error: {e}", exc_info=True)

    def _on_manual_runtimes_finished(self, success: bool, message: str):
        """Called when manual runtimes button finishes - runs in main thread"""
        try:
            logger.info(f"[MANUAL RUNTIMES] Finished: success={success}")
            self._set_runtimes_ui_running(False, message)
            if not success:
                QtWidgets.QMessageBox.warning(self, translator.t('messages.install.title'), message)
        except Exception as e:
            logger.error(f"_on_manual_runtimes_finished error: {e}", exc_info=True)

    def _on_driver_finished_from_installthread(self, success: bool, message: str):
        """Handler for driver installation result emitted by InstallThread."""
        try:
            logger.info(f"[DRIVER] Finished callback invoked: success={success}")
            # Store for final summary
            try:
                self._driver_install_result = (success, message)
            except Exception:
                self._driver_install_result = (success, message)
        except Exception as e:
            logger.error(f"_on_driver_finished_from_installthread error: {e}", exc_info=True)

    def install_diagbox(self):
        logger.info("Install Diagbox initiated")
    
        
        # If a Diagbox version is already installed, require cleaning first
        installed_version = self.check_installed_version()
        if installed_version:
            reply = QtWidgets.QMessageBox.question(
                self,
                translator.t('messages.install.title'),
                translator.t('messages.install.must_clean', installed=installed_version),
                QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                QtWidgets.QMessageBox.StandardButton.Yes
            )
            if reply == QtWidgets.QMessageBox.StandardButton.Yes:
                # Start the cleaning flow and return; user must run install after cleaning
                self.clean_diagbox()
            return
        # Get selected version from combo box
        if hasattr(self, 'version_combo'):
            selected_data = self.version_combo.currentData()
            if selected_data:
                version, url = selected_data
                self.last_version_diagbox = version
                # Try new format first (09.180.7z), then old format (Diagbox_Install_09.180.7z)
                new_format = os.path.join(self.download_folder, f"{version}.7z")
                old_format = os.path.join(self.download_folder, f"Diagbox_Install_{version}.7z")
                if os.path.exists(new_format):
                    self.diagbox_path = new_format
                else:
                    self.diagbox_path = old_format
                logger.info(f"Installing version: {version}, path: {self.diagbox_path}")
        
        if not os.path.exists(self.diagbox_path):
            logger.error(f"Diagbox file not found: {self.diagbox_path}")
            # Get the version being attempted
            version = self.last_version_diagbox if self.last_version_diagbox else "Unknown"
            QtWidgets.QMessageBox.warning(
                self, 
                translator.t('messages.install.title'), 
                translator.t('messages.install.file_not_found', version=version, path=self.diagbox_path)
            )
            return
        
        # Disable all buttons and combo box
        self.set_buttons_enabled(False)
        
        # Get install button reference
        install_button = None
        for child in self.stack.currentWidget().findChildren(QtWidgets.QPushButton):
            if child.text() == "Install":
                install_button = child
                break
        
        # Start installation in thread
        self.install_thread = InstallThread(self.diagbox_path)
        self.install_thread.progress.connect(self.update_install_progress)
        self.install_thread.file_progress.connect(self.update_install_file)
        # Connect runtimes signals so UI can reflect runtimes installer state
        try:
            self.install_thread.runtimes_started.connect(self._on_runtimes_started_from_installthread)
            self.install_thread.runtimes_finished.connect(self._on_runtimes_finished_from_installthread)
            # Connect driver finished signal so we can include it in the final summary
            try:
                self.install_thread.driver_finished.connect(self._on_driver_finished_from_installthread)
            except Exception:
                logger.debug("Failed to connect driver_finished signal")
            logger.info("[INSTALL] Runtimes signals connected successfully")
        except Exception as e:
            logger.error(f"[INSTALL] Failed to connect runtimes signals: {e}", exc_info=True)
        self.install_thread.finished.connect(lambda success, message: self.on_install_finished(success, message, install_button, None))
        self.install_thread.start()

    def clean_diagbox(self):
        """Clean Diagbox installation by removing C:\\APP, C:\\AWRoot, and C:\\APPLIC folders"""
        logger.info("Clean Diagbox initiated")
        # Check which folders exist
        folders_to_delete = []
        folders = [r"C:\APP", r"C:\AWRoot", r"C:\APPLIC"]
        
        for folder in folders:
            if os.path.exists(folder):
                folders_to_delete.append(folder)
        
        # Check for desktop shortcuts
        shortcuts_to_delete = []
        public_desktop = r"C:\Users\Public\Desktop"
        # Known public desktop shortcut filenames (include variants)
        shortcut_names = [
            "Diagbox Language Changer.lnk",
            "Change Diagbox Language.lnk",
            "Diagbox.lnk",
            "PSA Interface Checker.lnk",
            "PSA XS Evolution Interface Checker.lnk",
            "Terminate Diagbox Process.lnk",
            "Terminate Diagbox Processes.lnk"
        ]
        
        for shortcut in shortcut_names:
            shortcut_path = os.path.join(public_desktop, shortcut)
            if os.path.exists(shortcut_path):
                shortcuts_to_delete.append(shortcut_path)

        # Also detect specific VCI driver FileRepository folder and its .ini
        vcomm_folder = r"C:\Windows\System32\DriverStore\FileRepository\vcommusb.inf_amd64_0cb1ee01f7e64ab9"
        vcomm_ini = vcomm_folder + ".ini"
        try:
            if os.path.exists(vcomm_folder):
                folders_to_delete.append(vcomm_folder)
            if os.path.exists(vcomm_ini):
                shortcuts_to_delete.append(vcomm_ini)
        except Exception:
            pass

        if not folders_to_delete and not shortcuts_to_delete:
            QtWidgets.QMessageBox.information(
                self,
                translator.t('messages.clean.title'),
                translator.t('messages.clean.nothing_to_clean')
            )
            return
        # Prepare driver items separately so they are displayed in their own section
        driver_items = []
        try:
            if vcomm_folder in folders_to_delete:
                try:
                    folders_to_delete.remove(vcomm_folder)
                except Exception:
                    pass
                driver_items.append(vcomm_folder)
            if vcomm_ini in shortcuts_to_delete:
                try:
                    shortcuts_to_delete.remove(vcomm_ini)
                except Exception:
                    pass
                driver_items.append(vcomm_ini)
        except Exception:
            pass

        # Confirm deletion - build sections: Folders, Shortcuts, Driver
        items_list = []
        if folders_to_delete:
            items_list.append("Folders:")
            items_list.extend([f"- {folder}" for folder in folders_to_delete])

        if shortcuts_to_delete:
            if items_list:
                items_list.append("")
            items_list.append("Shortcuts:")
            items_list.extend([f"- {os.path.basename(s)}" for s in shortcuts_to_delete])

        if driver_items:
            if items_list:
                items_list.append("")
            items_list.append("Driver:")
            # Use a human-friendly label for the driver entry
            items_list.append(f"- PSA XS Evolution (vcommusb.inf)")

        items_text = "\n".join(items_list)
        reply = QtWidgets.QMessageBox.question(
            self,
            translator.t('messages.clean.title'),
            translator.t('messages.clean.confirm', items=items_text),
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No
        )
        
        if reply == QtWidgets.QMessageBox.StandardButton.No:
            return
        
        # Disable all buttons
        self.set_buttons_enabled(False)
        
        # Kill all Diagbox processes before cleaning
        self.kill_diagbox_processes_silent()
        
        # Initialize footer progress
        total_items = len(folders_to_delete) + len(shortcuts_to_delete)
        if hasattr(self, 'footer_progress'):
            # Show simulated determinate progress during cleaning (0..total_items)
            self.footer_progress.setRange(0, total_items if total_items > 0 else 1)
            self.footer_progress.setValue(0)
        if hasattr(self, 'footer_label'):
            self.footer_label.setText("Cleaning Diagbox...")
        
        # Start cleaning in thread (pass driver items separately)
        try:
            driver_items = driver_items  # defined earlier when building confirmation
        except Exception:
            driver_items = []

        self.clean_thread = CleanThread(folders_to_delete, shortcuts_to_delete, driver_items)
        self.clean_thread.progress.connect(self.update_clean_progress)
        self.clean_thread.item_progress.connect(self.update_clean_item)
        self.clean_thread.finished.connect(self.on_clean_finished)
        self.clean_thread.start()
    
    def update_clean_progress(self, current, total):
        """Update clean progress bar"""
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setValue(current)
    
    def update_clean_item(self, item_name):
        """Update current item being cleaned"""
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(item_name)
    
    def on_clean_finished(self, success, message, success_count):
        """Called when cleaning is finished"""
        # Re-enable buttons
        self.set_buttons_enabled(True)
        
        # Update footer to show completion
        if hasattr(self, 'footer_label'):
            self.footer_label.setText("Clean complete")
        if hasattr(self, 'footer_progress'):
            # If progress was indeterminate (maximum == 0), switch to determinate and show complete
            try:
                if self.footer_progress.maximum() == 0:
                    self.footer_progress.setRange(0, 100)
                    self.footer_progress.setValue(100)
                else:
                    total_items = self.footer_progress.maximum()
                    self.footer_progress.setValue(total_items)
            except Exception:
                try:
                    self.footer_progress.setRange(0, 100)
                    self.footer_progress.setValue(100)
                except Exception:
                    pass
        
        # Refresh install page
        self.refresh_install_page()
        
        # Show result
        if success:
            QtWidgets.QMessageBox.information(self, translator.t('messages.clean.title'), message)
        else:
            QtWidgets.QMessageBox.warning(self, translator.t('messages.clean.title'), message)
        
        # Reset footer after a delay
        QtCore.QTimer.singleShot(3000, self.reset_footer)

    def install_vci_driver(self):
        """Install VCI Driver using DPInst (/PATH ... /S)"""
        # Button handler: call unified installer and show result dialogs
        logger.info("Install VCI Driver initiated (button)")
        ok, msg = self._install_vci_driver_core()
        if ok:
            QtWidgets.QMessageBox.information(self, translator.t('messages.vci_driver.title'), msg)
        else:
            QtWidgets.QMessageBox.warning(self, translator.t('messages.vci_driver.title'), msg)

    def _install_vci_driver_auto(self):
        """Called automatically after runtimes installation completes.

        This runs the same core installer but does not show modal dialogs; it
        logs results and will show a non-blocking message only on failure.
        """
        logger.info("Automatic VCI Driver installation triggered")
        try:
            ok, msg = self._install_vci_driver_core()
            if not ok:
                # Show a non-blocking warning to the user
                QtWidgets.QMessageBox.warning(self, translator.t('messages.vci_driver.title'), msg)
            else:
                logger.info("Automatic VCI Driver install succeeded")
        except Exception as e:
            logger.error(f"Automatic VCI Driver installation error: {e}", exc_info=True)

    def _install_vci_driver_core(self):
        """Core installer used by both manual and automatic flows.

        Returns (success: bool, message: str).
        Also verifies installation by checking for the presence of the expected
        .ini file in the DriverStore FileRepository.
        """
        logger.info("VCI Driver core installer starting")
        # Use DPInst for driver installation per updated requirement
        dpinst_path = r"C:\AWRoot\extra\drivers\xsevo\amd64\DPInst.exe"
        dp_path_arg = r"C:\AWRoot\extra\drivers\xsevo\dp"
        ini_check = r"C:\Windows\System32\DriverStore\FileRepository\vcommusb.inf_amd64_0cb1ee01f7e64ab9.ini"

        if not os.path.exists(dpinst_path):
            logger.error(f"DPInst not found: {dpinst_path}")
            return False, translator.t('messages.vci_driver.not_found', path=dpinst_path)

        # Quick check: if .ini already present, consider driver already installed
        try:
            if os.path.exists(ini_check):
                logger.info("VCI .ini already present - driver considered installed")
                return True, translator.t('messages.vci_driver.already_present')
        except Exception:
            pass

        try:
            logger.info(f"Launching DPInst with /PATH {dp_path_arg} (interactive mode) from {os.path.dirname(dpinst_path)}")
            proc = subprocess.run(
                [dpinst_path, '/PATH', dp_path_arg],
                capture_output=True,
                text=True,
                cwd=os.path.dirname(dpinst_path)
            )
            out = (proc.stdout or '').strip()
            err = (proc.stderr or '').strip()
            rc = proc.returncode
            logger.debug(f"DPInst returned code={rc}")
            if out:
                logger.debug(f"DPInst stdout: {out[:2000]}")
            if err:
                logger.warning(f"DPInst stderr: {err[:2000]}")

            # Consider 0 as success; 256 indicates success but reboot required
            if rc == 0:
                # Verify by checking for the .ini file
                try:
                    if os.path.exists(ini_check):
                        msg = translator.t('messages.vci_driver.success')
                        logger.info("VCI Driver installation verified via .ini presence")
                        return True, msg
                    else:
                        logger.warning("DPInst returned success but .ini not found")
                        return False, translator.t('messages.vci_driver.warning', code=rc)
                except Exception as e:
                    logger.error(f"Error while verifying VCI .ini existence: {e}")
                    return False, translator.t('messages.vci_driver.error', error=str(e))
            elif rc == 256:
                # Some DPInst return codes indicate a reboot is required; treat as success but advise reboot
                msg = translator.t('messages.vci_driver.success_reboot')
                logger.info("DPInst reported reboot required; treating as success requiring reboot")
                return True, msg
            else:
                logger.warning(f"DPInst failed with code {rc}: {err[:400]}")
                return False, translator.t('messages.vci_driver.warning', code=rc) + "\n" + (err or '')
        except Exception as e:
            logger.error(f"DPInst installation exception: {e}", exc_info=True)
            return False, translator.t('messages.vci_driver.error', error=str(e))

    def launch_diagbox(self):
        """Launch Diagbox application"""
        logger.info("Launch Diagbox initiated")
        diagbox_exe = r"C:\AWRoot\bin\launcher\Diagbox.exe"
        
        # Check if Diagbox.exe exists
        if not os.path.exists(diagbox_exe):
            logger.error(f"Diagbox.exe not found: {diagbox_exe}")
            QtWidgets.QMessageBox.warning(
                self,
                translator.t('messages.launch.title'),
                translator.t('messages.launch.not_found', path=diagbox_exe)
            )
            return
        
        try:
            # Launch Diagbox.exe
            subprocess.Popen([diagbox_exe], cwd=r"C:\AWRoot\bin\launcher")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                translator.t('messages.launch.title'),
                translator.t('messages.launch.error', error=str(e))
            )

    def kill_diagbox_processes_silent(self):
        """Kill all Diagbox related processes silently (no message)"""
        # List of Diagbox process names to kill
        process_names = [
            "AWFInterpreter_vc80.exe",
            "LctPOLUX.exe",
            "AWRSrv.exe",
            "MCComm.exe",
            "fbguard.exe",
            "fbserver.exe",
            "httpd_ddc.exe",
            "diagnostic.exe",
            "awrsrv.exe",
            "awacscmd.exe",
            "awrcmd.exe",
            "AWACSserver.exe",
            "psaagent.exe",
            "psaSingleSignOnDaemon.exe",
            "psalance.exe",
            "mccomm.exe",
            "sim.exe",
            "firefoxportable.exe",
            "Ftspssrv.exe",
            "j9w.exe",
            "eclipse.exe",
            "Java.exe",
            "Jusched.exe",
            "Pg_ctl.exe",
            "Postgres.exe",
            "Sed.exe",
            "LCTPolux.exe",
            "DccFsmRunner.exe",
            "DdcECUReader.exe",
            "WSTransformer.exe",
            "partialtrace.exe",
            "psainterfaceservice.exe",
            "FirefoxPortable.exe",
            "psaAgent.exe",
            "Ground.exe",
            "instsvc.exe",
            "instreg.exe",
            "Psarefreshredwire.exe",
            "PSA-AUTH_Killer.exe",
            "Diagbox.exe"
        ]
        
        try:
            # Iterate through all running processes
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    # Check if process name matches any Diagbox process (case-insensitive)
                    proc_name = proc.info['name']
                    if any(proc_name.lower() == pname.lower() for pname in process_names):
                        proc.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        except:
            pass  # Silent failure

    def kill_diagbox(self):
        """Kill all Diagbox related processes with user feedback"""
        killed_count = 0
        
        # Get list before killing to count
        process_names = [
            "AWFInterpreter_vc80.exe", "LctPOLUX.exe", "AWRSrv.exe", "MCComm.exe",
            "fbguard.exe", "fbserver.exe", "httpd_ddc.exe", "diagnostic.exe",
            "awrsrv.exe", "awacscmd.exe", "awrcmd.exe", "AWACSserver.exe",
            "psaagent.exe", "psaSingleSignOnDaemon.exe", "psalance.exe", "mccomm.exe",
            "sim.exe", "firefoxportable.exe", "Ftspssrv.exe", "j9w.exe",
            "eclipse.exe", "Java.exe", "Jusched.exe", "Pg_ctl.exe",
            "Postgres.exe", "Sed.exe", "LCTPolux.exe", "DccFsmRunner.exe",
            "DdcECUReader.exe", "WSTransformer.exe", "partialtrace.exe",
            "psainterfaceservice.exe", "FirefoxPortable.exe", "psaAgent.exe",
            "Ground.exe", "instsvc.exe", "instreg.exe", "Psarefreshredwire.exe",
            "PSA-AUTH_Killer.exe", "Diagbox.exe"
        ]
        
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    proc_name = proc.info['name']
                    if any(proc_name.lower() == pname.lower() for pname in process_names):
                        proc.kill()
                        killed_count += 1
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            if killed_count > 0:
                QtWidgets.QMessageBox.information(
                    self,
                    translator.t('messages.kill_process.title'),
                    translator.t('messages.kill_process.success', count=killed_count)
                )
            else:
                QtWidgets.QMessageBox.information(
                    self,
                    translator.t('messages.kill_process.title'),
                    translator.t('messages.kill_process.none')
                )
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                translator.t('messages.kill_process.title'),
                translator.t('messages.kill_process.error', error=str(e))
            )

    def on_language_changed(self):
        """Handle language change from combo box"""
        lang_code = self.language_combo.currentData()
        if lang_code:
            self.change_diagbox_language(lang_code)
    
    def on_app_language_changed(self):
        """Handle application language change"""
        new_lang = self.app_language_combo.currentData()
        if new_lang and new_lang != translator.language:
            translator.set_language(new_lang)
            # Show restart dialog
            reply = QtWidgets.QMessageBox.question(
                self,
                "Language Changed / Langue Modifiée",
                "Please restart the application for the language change to take full effect.\n\n"
                "Veuillez redémarrer l'application pour que le changement de langue prenne pleinement effet.\n\n"
                "Close now? / Quitter maintenant ?",
                QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                QtWidgets.QMessageBox.StandardButton.Yes
            )
            
            if reply == QtWidgets.QMessageBox.StandardButton.Yes:
                # Restart application
                QtWidgets.QApplication.quit()

    def fetch_last_version_diagbox(self):
        try:
            logger.info(f"Fetching last version from: {URL_LAST_VERSION_DIAGBOX}")
            response = requests.get(URL_LAST_VERSION_DIAGBOX)
            response.raise_for_status()
            data = response.json()
            self.last_version_diagbox = data.get('version', '')
            logger.info(f"Last Diagbox version: {self.last_version_diagbox}")
            # Try new format first, then old format
            new_format = os.path.join(self.download_folder, f"{self.last_version_diagbox}.7z")
            old_format = os.path.join(self.download_folder, f"Diagbox_Install_{self.last_version_diagbox}.7z")
            if os.path.exists(new_format):
                self.diagbox_path = new_format
            else:
                self.diagbox_path = old_format
        except Exception as e:
            logger.error(f"Failed to fetch last version: {e}", exc_info=True)
            QtWidgets.QMessageBox.warning(self, "Error", f"Failed to fetch last version: {e}")

    def check_app_update(self):
        """Check if a newer version of PSA-DIAG is available"""
        try:
            logger.info("Checking for app updates...")
            response = requests.get(URL_LAST_VERSION_PSADIAG, timeout=5)
            response.raise_for_status()
            data = response.json()
            latest_version = data.get('version', '')
            logger.info(f"Latest app version: {latest_version}, Current: {APP_VERSION}")
            
            if latest_version and latest_version != APP_VERSION:
                # Compare versions (simple string comparison, assumes format like "2.0.0.0")
                current_parts = [int(x) for x in APP_VERSION.split('.')]
                latest_parts = [int(x) for x in latest_version.split('.')]
                
                # Check if latest is newer
                if latest_parts > current_parts:
                    logger.info("New version available, showing update dialog")
                    reply = QtWidgets.QMessageBox.question(
                        self,
                        translator.t('messages.update.title'),
                        translator.t('messages.update.available', current=APP_VERSION, latest=latest_version),
                        QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                        QtWidgets.QMessageBox.StandardButton.Yes
                    )
                    
                    if reply == QtWidgets.QMessageBox.StandardButton.Yes:
                        logger.info("User accepted update, performing automatic download and update")
                        # Perform automatic download of the latest release asset and run updater
                        try:
                            QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
                            self.perform_self_update(latest_version)
                        finally:
                            QtWidgets.QApplication.restoreOverrideCursor()
                else:
                    logger.info("App is up to date")
        except Exception as e:
            # Silently fail if update check fails (no internet, server down, etc.)
            logger.warning(f"Update check failed: {e}")

    def perform_self_update(self, latest_version):
        """Download the latest release .exe from GitHub and invoke the updater helper.

        Note: This implementation uses the GitHub Releases API to locate a release asset
        that looks like an executable (.exe). It downloads to `CONFIG_DIR/updates/`.
        It then launches the `updater.py` helper (bundled in the project) which will
        wait for this process to exit, replace the running exe, and optionally restart it.
        """
        try:
            api_url = "https://api.github.com/repos/RetroGameSets/PSA-DIAG/releases/latest"
            logger.info(f"Querying GitHub API for latest release: {api_url}")
            r = requests.get(api_url, timeout=10)
            r.raise_for_status()
            release = r.json()

            # Find a downloadable .exe asset
            exe_asset = None
            for asset in release.get('assets', []):
                name = asset.get('name', '')
                if name.lower().endswith('.exe'):
                    exe_asset = asset
                    break

            if not exe_asset:
                QtWidgets.QMessageBox.warning(self, translator.t('messages.update.title'), translator.t('messages.update.no_asset'))
                return

            download_url = exe_asset.get('browser_download_url')
            if not download_url:
                QtWidgets.QMessageBox.warning(self, translator.t('messages.update.title'), translator.t('messages.update.no_asset'))
                return

            updates_dir = CONFIG_DIR / 'updates'
            updates_dir.mkdir(parents=True, exist_ok=True)
            target_name = exe_asset.get('name') or f"PSA-DIAG-{latest_version}.exe"
            download_path = str(updates_dir / target_name)

            # If a previous download exists, remove it first to avoid replace conflicts
            try:
                if os.path.exists(download_path):
                    logger.info(f"Removing existing downloaded update before re-downloading: {download_path}")
                    os.remove(download_path)
            except Exception as e:
                logger.warning(f"Failed to remove previous download {download_path}: {e}")

            # Download with progress (simple)
            logger.info(f"Downloading update asset: {download_url} -> {download_path}")
            with requests.get(download_url, stream=True, timeout=30) as resp:
                resp.raise_for_status()
                with open(download_path, 'wb') as f:
                    for chunk in resp.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)

            logger.info(f"Download complete: {download_path}")

            # Launch updater helper and exit
            updater_script = BASE / 'updater.py'
            if not updater_script.exists():
                # Fallback: use embedded updater in CONFIG_DIR if present
                updater_script = CONFIG_DIR / 'updater.py'

            if not updater_script.exists():
                QtWidgets.QMessageBox.warning(self, translator.t('messages.update.title'), translator.t('messages.update.no_updater'))
                return

            # Determine current executable to replace
            current_exe = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])

            # Spawn updater: choose the correct launcher depending on frozen state
            logger.info("Preparing to launch updater helper")
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0

            try:
                if getattr(sys, 'frozen', False):
                    # When frozen, BASE points to a temp _MEIPASS folder that gets cleaned up.
                    # Copy the bundled updater to a persistent location (CONFIG_DIR) before invoking it.
                    # Priority: standalone updater.exe in tools/, then onedir updater/updater.exe
                    bundled_updater_standalone = BASE / 'tools' / 'updater.exe'
                    bundled_updater_dir = BASE / 'tools' / 'updater' / 'updater.exe'
                    persistent_updater = CONFIG_DIR / 'updater.exe'
                    
                    # Ensure CONFIG_DIR exists
                    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
                    
                    # Copy bundled updater to persistent location if found
                    if bundled_updater_standalone.exists():
                        logger.info(f"Copying standalone updater from {bundled_updater_standalone} to {persistent_updater}")
                        shutil.copy2(bundled_updater_standalone, persistent_updater)
                    elif bundled_updater_dir.exists():
                        logger.info(f"Copying onedir updater from {bundled_updater_dir} to {persistent_updater}")
                        shutil.copy2(bundled_updater_dir, persistent_updater)
                    
                    # Use the persistent copy if it exists
                    if persistent_updater.exists():
                        cmd = [str(persistent_updater), '--target', str(current_exe), '--new', str(download_path), '--wait-pid', str(os.getpid()), '--restart']
                    else:
                        # Fallback: Try system python to run updater.py (best-effort)
                        python_bin = shutil.which('python') or shutil.which('python3') or shutil.which('py')
                        if python_bin:
                            cmd = [python_bin, str(updater_script), '--target', str(current_exe), '--new', str(download_path), '--wait-pid', str(os.getpid()), '--restart']
                        else:
                            logger.error("No bundled updater and no python in PATH; cannot self-update from frozen exe")
                            QtWidgets.QMessageBox.warning(self, translator.t('messages.update.title'), translator.t('messages.update.no_updater'))
                            return
                else:
                    # Running from source/interpreter: use current python interpreter
                    cmd = [sys.executable, str(updater_script), '--target', str(current_exe), '--new', str(download_path), '--wait-pid', str(os.getpid()), '--restart']

                # Add timeout for handle release (reduced to 15s since process exits quickly)
                cmd.extend(['--timeout', '15'])
                logger.info(f"Launching updater helper: {cmd}")
                
                # Close the application FIRST to release all handles and temp folders
                # before launching the updater (prevents temp folder deletion warnings)
                try:
                    # Shutdown logging to help release any file handles held by this process
                    import logging as _logging
                    _logging.shutdown()
                except Exception:
                    pass
                
                # Show a brief message to user before closing
                QtWidgets.QMessageBox.information(
                    self, 
                    translator.t('messages.update.title'),
                    translator.t('messages.update.closing_app')
                )
                
                # Launch updater with a small delay to ensure clean shutdown
                # Ensure updater window is visible: do NOT use CREATE_NO_WINDOW for updater
                creationflags_for_updater = 0
                subprocess.Popen(cmd, creationflags=creationflags_for_updater)
                
                # Give subprocess time to start and PyInstaller time to finish cleanup
                # This prevents the temp folder warning
                QtCore.QThread.msleep(500)
            except Exception as e:
                logger.error(f"Failed to launch updater helper: {e}")
                QtWidgets.QMessageBox.warning(self, translator.t('messages.update.title'), translator.t('messages.update.launch_failed', error=str(e)))
                return

            # Close the application to allow updater to replace the executable
            QtWidgets.QApplication.quit()

        except Exception as e:
            logger.error(f"Self-update failed: {e}", exc_info=True)
            QtWidgets.QMessageBox.warning(self, translator.t('messages.update.title'), translator.t('messages.update.failed', error=str(e)))

    def download_diagbox(self):
        logger.info("Download Diagbox button clicked")
        
        # Check system requirements before proceeding
        if not self.check_system_requirements():
            logger.info("Download cancelled due to system requirements")
            return
        
        # Get selected version from combo box
        if hasattr(self, 'version_combo'):
            selected_data = self.version_combo.currentData()
            if selected_data:
                version, url = selected_data
                self.last_version_diagbox = version
                logger.info(f"Selected version: {version}")
            else:
                if not self.last_version_diagbox:
                    self.fetch_last_version_diagbox()
                if not self.last_version_diagbox:
                    return
                url = f"https://archive.org/download/psa-diag.fr/Diagbox_Install_{self.last_version_diagbox}.7z"
        else:
            if not self.last_version_diagbox:
                self.fetch_last_version_diagbox()
            if not self.last_version_diagbox:
                return
            url = f"https://archive.org/download/psa-diag.fr/Diagbox_Install_{self.last_version_diagbox}.7z"
        
        # Set path to new format (downloaded files will use new naming)
        self.diagbox_path = os.path.join(self.download_folder, f"{self.last_version_diagbox}.7z")
        
        if not os.path.exists(self.download_folder):
            os.makedirs(self.download_folder)
        
        file_path = self.diagbox_path
        
        # Get total size for progress
        try:
            # Follow redirects to get actual file size
            head = requests.head(url, allow_redirects=True, timeout=10)
            total_size = int(head.headers.get('content-length', 0))
            logger.info(f"File size from HEAD request: {total_size / (1024*1024):.2f} MB")
        except Exception as e:
            logger.warning(f"Could not get file size from HEAD request: {e}")
            total_size = 0
        
        # Check if file already exists and size matches
        if os.path.exists(file_path):
            if total_size > 0 and os.path.getsize(file_path) == total_size:
                QtWidgets.QMessageBox.information(self, translator.t('messages.download.title'), translator.t('messages.download.already_downloaded', version=self.last_version_diagbox))
                return
            else:
                # File exists but size doesn't match, delete it
                os.remove(file_path)
        
        # Disable all buttons and combo box
        self.set_buttons_enabled(False)
        
        # Show cancel and pause buttons
        if self.cancel_button:
            self.cancel_button.setVisible(True)
        if self.pause_button:
            self.pause_button.setVisible(True)
            self.pause_button.setText(translator.t('buttons.pause'))
        
        # Set progress bar
        bar = self.stack.currentWidget().findChild(QtWidgets.QProgressBar)
        if bar:
            if total_size > 0:
                bar.setRange(0, 1000)
                bar.setValue(0)
                bar.setFormat("0.0% - 0.0 MB/s - --:--")
            else:
                bar.setRange(0, 0)  # indeterminate
        self.download_thread = DownloadThread(url, file_path, self.last_version_diagbox, total_size)
        self.download_thread.progress.connect(self.update_progress)
        self.download_thread.finished.connect(self.on_download_finished)
        self.download_thread.start()

    def setup_ui(self):
        # Main layout
        main_layout = QtWidgets.QHBoxLayout(self)
        main_layout.setContentsMargins(10,10,10,10)

        # Sidebar
        sidebar = QtWidgets.QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(110)
        vbox = QtWidgets.QVBoxLayout(sidebar)
        vbox.setContentsMargins(12,12,12,12)
        vbox.setSpacing(14)

        # Title in sidebar
        title_sidebar = QtWidgets.QLabel(translator.t('app.title'))
        title_sidebar.setObjectName("titleLabel")
        title_sidebar.setWordWrap(True)
        title_sidebar.setStyleSheet("font-size: 13px;")
        title_sidebar.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        vbox.addWidget(title_sidebar)

        icon_folder = BASE / "icons"
        btn_diag = SidebarButton("", icon_folder / "diag.svg")
        btn_setup = SidebarButton("", icon_folder / "setup.svg")
        btn_update = SidebarButton("", icon_folder / "update.svg")
        btn_info = SidebarButton("", icon_folder / "info.svg")

        # Make first checked
        btn_diag.setChecked(True)

        vbox.addWidget(btn_diag)
        vbox.addWidget(btn_setup)
        vbox.addWidget(btn_update)
        vbox.addStretch()
        vbox.addWidget(btn_info)

        # Display app version under Info button in the sidebar
        self.sidebar_version_label = QtWidgets.QLabel(translator.t('labels.version', version=APP_VERSION))
        self.sidebar_version_label.setObjectName("sidebarVersion")
        self.sidebar_version_label.setStyleSheet("font-size:11px; color: #b0b0b0;")
        self.sidebar_version_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        vbox.addWidget(self.sidebar_version_label)

        # Title bar and content
        content = QtWidgets.QFrame()
        content.setObjectName("content")
        content_layout = QtWidgets.QVBoxLayout(content)
        content_layout.setContentsMargins(14,8,14,14)
        content_layout.setSpacing(10)

        # Custom title row (min/close only)
        title_row = QtWidgets.QHBoxLayout()
        title_row.addStretch()
        btn_min = QtWidgets.QPushButton(translator.t('buttons.minimize'))
        btn_close = QtWidgets.QPushButton(translator.t('buttons.close'))
        btn_min.setFixedSize(34,28)
        btn_close.setFixedSize(34,28)
        btn_min.setObjectName("titleButton")
        btn_close.setObjectName("titleButton")
        title_row.addWidget(btn_min)
        title_row.addWidget(btn_close)

        content_layout.addLayout(title_row)

        # Stacked widget with pages
        self.stack = QtWidgets.QStackedWidget()
        self.stack.addWidget(self.page_config())
        self.stack.addWidget(self.page_install())
        self.stack.addWidget(self.page_update())
        self.stack.addWidget(self.page_about())

        content_layout.addWidget(self.stack)
        
        # Footer with progress bar
        self.footer = QtWidgets.QFrame()
        self.footer.setObjectName("footer")
        self.footer.setFixedHeight(60)
        footer_layout = QtWidgets.QVBoxLayout(self.footer)
        footer_layout.setContentsMargins(10, 8, 10, 8)
        footer_layout.setSpacing(4)
        
        # Progress label
        self.footer_label = QtWidgets.QLabel(translator.t('labels.ready'))
        self.footer_label.setStyleSheet("color: #b0b0b0; font-size: 11px;")
        footer_layout.addWidget(self.footer_label)
        
        # Progress bar
        self.footer_progress = QtWidgets.QProgressBar()
        self.footer_progress.setRange(0, 1000)
        self.footer_progress.setValue(0)
        self.footer_progress.setTextVisible(True)
        self.footer_progress.setFormat("")
        self.footer_progress.setFixedHeight(20)
        footer_layout.addWidget(self.footer_progress)
        
        content_layout.addWidget(self.footer)

        main_layout.addWidget(sidebar)
        main_layout.addWidget(content, 1)

        # Connections
        btn_diag.clicked.connect(lambda: self.switch_page(0, btn_diag))
        btn_setup.clicked.connect(lambda: self.switch_page(1, btn_setup))
        btn_update.clicked.connect(lambda: self.switch_page(2, btn_update))
        btn_info.clicked.connect(lambda: self.switch_page(3, btn_info))
        btn_close.clicked.connect(self.close)
        btn_min.clicked.connect(self.showMinimized)

        # Apply style
        self.setStyleSheet(load_qss())

        

        # Initial system check
        self.check_system()

    def check_system(self):
        # Check OS
        system = platform.system()
        release = platform.release()
        if system == "Windows":
            version_info = sys.getwindowsversion()
            if version_info.major == 10:
                if version_info.build >= 22000:
                    os_text = "Windows 11"
                else:
                    os_text = "Windows 10"
                os_text += " 64 Bits" if platform.machine().endswith('64') else " 32 Bits"
            else:
                os_text = f"{system} {release}"
        else:
            os_text = f"{system} {release}"

        # Check RAM
        ram_gb = psutil.virtual_memory().total / (1024 ** 3)
        ram_ok = ram_gb >= 3
        ram_text = f"{ram_gb:.1f} GB"
        if not ram_ok:
            ram_text += " (min 3 GB)"

        # Check free storage (C: drive)
        try:
            storage = psutil.disk_usage('C:\\')
            free_gb = storage.free / (1024 ** 3)
            storage_ok = free_gb >= 15
            storage_text = f"{free_gb:.1f} GB"
            if not storage_ok:
                storage_text += " (min 15 GB)"
        except:
            storage_text = "N/A"
            storage_ok = True  # Assume ok if can't check

        # Store as instance attributes for later checks
        self.ram_ok = ram_ok
        self.ram_gb = ram_gb
        self.storage_ok = storage_ok
        self.free_gb = free_gb if 'free_gb' in locals() else 0

        # Update labels if they exist
        if hasattr(self, 'os_label'):
            self.os_label.setText(os_text)
            self.os_label.setStyleSheet("")  # Default
        if hasattr(self, 'ram_label'):
            self.ram_label.setText(ram_text)
            self.ram_label.setStyleSheet("color: red;" if not ram_ok else "")
        if hasattr(self, 'storage_label'):
            self.storage_label.setText(storage_text)
            self.storage_label.setStyleSheet("color: red;" if not storage_ok else "")

    def check_system_requirements(self):
        """Check if system meets minimum requirements and show warning if not.
        Returns True if requirements are met, False otherwise."""
        problems = []
        solutions = []
        
        # Check RAM requirement
        if hasattr(self, 'ram_ok') and not self.ram_ok:
            ram_gb = getattr(self, 'ram_gb', 0)
            problems.append(translator.t('messages.requirements.low_ram', current=f"{ram_gb:.1f} GB", minimum="3 GB"))
            solutions.append(translator.t('messages.requirements.solution_ram'))
        
        # Check storage requirement
        if hasattr(self, 'storage_ok') and not self.storage_ok:
            free_gb = getattr(self, 'free_gb', 0)
            problems.append(translator.t('messages.requirements.low_storage', current=f"{free_gb:.1f} GB", minimum="15 GB"))
            solutions.append(translator.t('messages.requirements.solution_storage'))
        
        # If there are problems, show warning dialog
        if problems:
            problem_text = "\n• " + "\n• ".join(problems)
            solution_text = "\n\n" + translator.t('messages.requirements.solutions') + "\n• " + "\n• ".join(solutions)
            
            reply = QtWidgets.QMessageBox.warning(
                self,
                translator.t('messages.requirements.title'),
                translator.t('messages.requirements.warning') + problem_text + solution_text + "\n\n" + translator.t('messages.requirements.continue'),
                QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                QtWidgets.QMessageBox.StandardButton.No
            )
            
            return reply == QtWidgets.QMessageBox.StandardButton.Yes
        
        return True

    def install_runtimes(self):
        """Start the bundled runtimes installer in a background thread."""
        try:
            # Update UI to show runtimes are starting and lock button
            QtCore.QTimer.singleShot(0, lambda: self._set_runtimes_ui_running(True, translator.t('messages.install.runtimes.started')))
            threading.Thread(target=self._run_runtimes_installer, daemon=True).start()
        except Exception as e:
            logger.error(f"Failed to start runtimes installer thread: {e}")

    def _run_runtimes_installer(self):
        runtimes_path = Path(r"C:\AWRoot\Extra\runtimes\runtimes.exe")
        if not runtimes_path.exists():
            msg = translator.t('messages.install.runtimes.not_found', path=str(runtimes_path))
            logger.warning(msg)
            self.manual_runtimes_finished.emit(False, msg)
            return

        logger.info(translator.t('messages.install.runtimes.started'))
        try:
            proc = subprocess.run(
                [str(runtimes_path), '/y'],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
            )
            
            # Log outputs
            try:
                out = (proc.stdout or '').strip()
                err = (proc.stderr or '').strip()
                if out:
                    logger.info(f"Manual runtimes stdout: {out[:2000]}")
                if err:
                    logger.warning(f"Manual runtimes stderr: {err[:2000]}")
            except Exception:
                pass
            
            # Build result message and emit signal
            if proc.returncode == 0:
                result_msg = translator.t('messages.install.runtimes.success')
                logger.info(result_msg)
                self.manual_runtimes_finished.emit(True, result_msg)
            else:
                result_msg = translator.t('messages.install.runtimes.failed', code=proc.returncode)
                logger.warning(f"{result_msg}: {proc.stderr[:400]}")
                result_msg += "\n\n" + (proc.stderr or "")
                self.manual_runtimes_finished.emit(False, result_msg)
            
        except Exception as e:
            msg = translator.t('messages.install.runtimes.error', error=str(e))
            logger.error(msg, exc_info=True)
            self.manual_runtimes_finished.emit(False, msg)
    def switch_page(self, index, button):
        # uncheck all sibling buttons in sidebar
        sidebar = self.findChild(QtWidgets.QFrame, "sidebar")
        if sidebar:
            for child in sidebar.findChildren(SidebarButton):
                child.setChecked(False)
        button.setChecked(True)
        self.stack.setCurrentIndex(index)
        # If navigating to install page, check for a newer Diagbox version
        if index == 1:
            # small delay to ensure UI has switched
            QtCore.QTimer.singleShot(100, self.on_enter_install_page)

    def parse_version_to_list(self, version_str):
        """Extract first numeric version (like '09.186' or '09.85') and
        convert to list of ints for lexicographic comparison.
        If no numeric version found, returns [0]."""
        try:
            m = re.search(r"(\d+(?:\.\d+)*)", str(version_str))
            if not m:
                return [0]
            parts = [int(x) for x in m.group(1).split('.')]
            return parts
        except Exception:
            return [0]

    def compare_versions(self, a, b):
        """Compare two version strings (return 1 if a>b, 0 if equal, -1 if a<b).
        Only numeric parts are compared (extract first numeric sequence).
        """
        a_list = self.parse_version_to_list(a)
        b_list = self.parse_version_to_list(b)
        # Pad with zeros
        max_len = max(len(a_list), len(b_list))
        a_list += [0] * (max_len - len(a_list))
        b_list += [0] * (max_len - len(b_list))
        for ai, bi in zip(a_list, b_list):
            if ai > bi:
                return 1
            if ai < bi:
                return -1
        return 0

    def get_latest_available_version(self):
        """Return the version string (raw) from `self.version_options` that
        is numerically the latest. If none, returns None."""
        try:
            best = None
            for _, version, _url in getattr(self, 'version_options', []):
                if not best:
                    best = version
                    continue
                if self.compare_versions(version, best) == 1:
                    best = version
            return best
        except Exception as e:
            logger.debug(f"Error determining latest available version: {e}")
            return None

    def on_enter_install_page(self):
        """Called when user navigates to the install page: check installed
        version vs the latest available and notify if newer exists."""
        try:
            installed = self.check_installed_version()
            latest = self.get_latest_available_version()
            # logger.info(f"on_enter_install_page: installed={installed!r}, latest={latest!r}")
            if not latest:
                return

            # Prepare banner text depending on whether Diagbox is installed
            if not installed:
                text = translator.t('messages.install.no_installed', latest=latest)
            else:
                cmp = self.compare_versions(latest, installed)
                if cmp == 1:
                    # latest is newer than installed
                    text = translator.t('messages.install.available', latest=latest, installed=installed)
                else:
                    # Installed is up-to-date or newer; show an up-to-date banner
                    text = translator.t('messages.install.up_to_date')

            # Populate and show banner (non-blocking) and hide the duplicate
            # installed/latest labels to avoid repeating the same info.
            if hasattr(self, 'install_banner_label'):
                self.install_banner_label.setText(text)
            if hasattr(self, 'install_banner'):
                before = self.install_banner.isVisible()
                # logger.info(f"install_banner before visible={before}")
                self.install_banner.setVisible(True)
                after = self.install_banner.isVisible()
                # logger.info(f"install_banner after visible={after}")
            # Show or hide the Download button depending on state. If we are
            # already up-to-date, hide the download action to avoid confusion.
            try:
                if hasattr(self, 'install_banner_download'):
                    up_to_date_msg = translator.t('messages.install.up_to_date')
                    if text == up_to_date_msg:
                        self.install_banner_download.setVisible(False)
                    else:
                        self.install_banner_download.setVisible(True)
            except Exception:
                pass
            if hasattr(self, 'header_installed'):
                try:
                    self.header_installed.setVisible(False)
                except Exception:
                    pass
            if hasattr(self, 'header_online'):
                try:
                    self.header_online.setVisible(False)
                except Exception:
                    pass
        except Exception as e:
            logger.debug(f"Error while checking install page updates: {e}")

    def page_config(self):
        w = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(w)
        layout.setSpacing(12)

        header = QtWidgets.QLabel(translator.t('labels.system_config'))
        header.setObjectName("sectionHeader")
        layout.addWidget(header)

        reqs = QtWidgets.QFrame()
        form = QtWidgets.QFormLayout(reqs)
        form.setLabelAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
        form.setFormAlignment(QtCore.Qt.AlignmentFlag.AlignTop)
        form.setHorizontalSpacing(30)
        self.os_label = QtWidgets.QLabel()
        self.storage_label = QtWidgets.QLabel()
        self.ram_label = QtWidgets.QLabel()
        form.addRow(translator.t('labels.windows_version'), self.os_label)
        form.addRow(translator.t('labels.free_storage'), self.storage_label)
        form.addRow(translator.t('labels.ram'), self.ram_label)
        layout.addWidget(reqs)

        # Application language selection
        lang_section = QtWidgets.QFrame()
        lang_layout = QtWidgets.QHBoxLayout(lang_section)
        lang_layout.setContentsMargins(0, 10, 0, 0)
        lang_label = QtWidgets.QLabel("Application Language :")
        self.app_language_combo = QtWidgets.QComboBox()
        self.app_language_combo.addItem("Français", userData="fr")
        self.app_language_combo.addItem("English", userData="en")
        
        # Set current language
        current_index = 0 if translator.language == 'fr' else 1
        self.app_language_combo.setCurrentIndex(current_index)
        
        self.app_language_combo.setMinimumWidth(150)
        self.app_language_combo.currentIndexChanged.connect(self.on_app_language_changed)
        lang_layout.addWidget(lang_label)
        lang_layout.addWidget(self.app_language_combo)
        lang_layout.addStretch()
        layout.addWidget(lang_section)

        layout.addStretch()
        # Row with Re-check system and Install runtimes buttons
        recheck = QtWidgets.QPushButton(translator.t('buttons.recheck'))
        recheck.setFixedWidth(160)
        recheck.clicked.connect(self.check_system)

        runtimes_btn = QtWidgets.QPushButton(translator.t('buttons.install_runtimes'))
        runtimes_btn.setFixedWidth(160)
        runtimes_btn.clicked.connect(self.install_runtimes)
        # keep a reference so we can enable/disable it during install
        self.runtimes_btn = runtimes_btn

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.setSpacing(8)
        btn_row.addWidget(recheck)
        btn_row.addWidget(runtimes_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        return w

    def page_install(self):
        w = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(w)
        layout.setSpacing(10)

        # Fetch online version if not already
        if not self.last_version_diagbox:
            self.fetch_last_version_diagbox()

        installed_version = self.check_installed_version()
        version_text = installed_version if installed_version else translator.t('labels.not_installed')
        self.header_installed = QtWidgets.QLabel(translator.t('labels.installed_version', version=version_text))
        self.header_installed.setObjectName("sectionHeader")
        layout.addWidget(self.header_installed)

        # Show the latest numeric version available (prefer version_options data)
        latest_available = self.get_latest_available_version() or (self.last_version_diagbox if self.last_version_diagbox else 'Unknown')
        self.header_online = QtWidgets.QLabel(translator.t('labels.last_version', version=latest_available))
        self.header_online.setObjectName("sectionHeader")

        # Inline non-blocking update banner (hidden by default). This replaces
        # the previous modal MessageBox notifications when the user enters the
        # install page. `on_enter_install_page()` will populate and show it.
        self.install_banner = QtWidgets.QFrame()
        self.install_banner.setObjectName('installBanner')
        self.install_banner.setVisible(False)
        banner_layout = QtWidgets.QHBoxLayout(self.install_banner)
        banner_layout.setContentsMargins(10,10,10,10)
        banner_layout.setSpacing(10)

        # Left icon/accent to make banner more visually distinct
        self.install_banner_icon = QtWidgets.QLabel("\u2139")
        self.install_banner_icon.setObjectName('installBannerIcon')
        self.install_banner_icon.setFixedWidth(28)
        self.install_banner_icon.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        banner_layout.addWidget(self.install_banner_icon)

        self.install_banner_label = QtWidgets.QLabel("")
        self.install_banner_label.setObjectName('installBannerLabel')
        self.install_banner_label.setWordWrap(True)
        banner_layout.addWidget(self.install_banner_label, 1)

        self.install_banner_download = QtWidgets.QPushButton(translator.t('buttons.download'))
        self.install_banner_download.setObjectName('bannerDownload')
        self.install_banner_download.clicked.connect(self._banner_download_clicked)
        banner_layout.addWidget(self.install_banner_download)

        # Dismiss button
        self.install_banner_dismiss = QtWidgets.QPushButton(translator.t('buttons.dismiss') if hasattr(translator, 't') else "Dismiss")
        self.install_banner_dismiss.setObjectName('bannerDismiss')
        self.install_banner_dismiss.clicked.connect(self._dismiss_install_banner)
        banner_layout.addWidget(self.install_banner_dismiss)

        layout.addWidget(self.install_banner)
        layout.addWidget(self.header_online)

        # Add a soft drop shadow to make the banner appear raised
        try:
            shadow = QtWidgets.QGraphicsDropShadowEffect(self.install_banner)
            shadow.setBlurRadius(18)
            shadow.setOffset(0, 4)
            shadow.setColor(QtGui.QColor(0, 0, 0, 160))
            self.install_banner.setGraphicsEffect(shadow)
        except Exception:
            pass

        # Check downloaded versions
        downloaded_versions = self.check_downloaded_versions()
        if downloaded_versions:
            downloaded_text = ", ".join([f"{v['version']} ({v['size_mb']:.1f} MB)" for v in downloaded_versions])
            self.header_downloaded = QtWidgets.QLabel(translator.t('labels.downloaded_versions', versions=downloaded_text))
            self.header_downloaded.setObjectName("sectionHeader")
            self.header_downloaded.setStyleSheet("color: #5cb85c;")
            layout.addWidget(self.header_downloaded)
        else:
            self.header_downloaded = None

        sub = QtWidgets.QHBoxLayout()
        left = QtWidgets.QVBoxLayout()
        right = QtWidgets.QVBoxLayout()

        # Version selection dropdown
        version_layout = QtWidgets.QHBoxLayout()
        version_label = QtWidgets.QLabel(translator.t('labels.select_version'))
        self.version_combo = QtWidgets.QComboBox()
        for display_name, version, url in self.version_options:
            self.version_combo.addItem(display_name, userData=(version, url))
        self.version_combo.setMinimumWidth(200)
        version_layout.addWidget(version_label)
        version_layout.addWidget(self.version_combo)
        version_layout.addStretch()
        right.addLayout(version_layout)

        # Toggle auto install
        h = QtWidgets.QHBoxLayout()
        lbl = QtWidgets.QLabel(translator.t('labels.auto_install'))
        toggle = QtWidgets.QCheckBox()
        self.auto_install = toggle
        h.addWidget(lbl)
        h.addWidget(toggle)
        h.addStretch()
        right.addLayout(h)

        # Language selection
        lang_layout = QtWidgets.QHBoxLayout()
        lang_label = QtWidgets.QLabel(translator.t('labels.diagbox_language'))
        self.language_combo = QtWidgets.QComboBox()

        # Read current Diagbox language from file (if present) so we can
        # ensure it's included and selected in the combo.
        current_lang = self.get_diagbox_language()

        # Default language codes in preferred order
        default_codes = [
            "en_GB", "fr_FR", "it_IT", "nl_NL", "pl_PL", "pt_PT",
            "ru_RU", "tr_TR", "sv_SE", "da_DK", "cs_CZ", "de_DE",
            "el_GR", "hr_HR", "zh_CN", "ja_JP", "es_ES", "sl_SI",
            "hu_HU", "fi_FI",
        ]

        # Build languages list using translations; if a translation is missing
        # fallback to the code string itself.
        languages = []
        for code in default_codes:
            name = translator.t(f'languages.{code}')
            # If translator returns the key back (missing), fallback to code
            if name == f'languages.{code}':
                name = code
            languages.append((name, code))

        # If current language exists and is not in defaults, insert it first
        if current_lang and current_lang not in [c for (_, c) in languages]:
            name = translator.t(f'languages.{current_lang}')
            if name == f'languages.{current_lang}':
                name = current_lang
            languages.insert(0, (name, current_lang))

        for display_name, lang_code in languages:
            self.language_combo.addItem(display_name, userData=lang_code)

        # Select current language if available
        if current_lang:
            for i in range(self.language_combo.count()):
                if self.language_combo.itemData(i) == current_lang:
                    self.language_combo.setCurrentIndex(i)
                    break
        
        self.language_combo.setMinimumWidth(150)
        self.language_combo.currentIndexChanged.connect(self.on_language_changed)
        lang_layout.addWidget(lang_label)
        lang_layout.addWidget(self.language_combo)
        lang_layout.addStretch()
        right.addLayout(lang_layout)

        # Buttons grid
        grid = QtWidgets.QGridLayout()
        grid.setSpacing(12)
        btns = [
            (translator.t('buttons.download'), "download"),
            (translator.t('buttons.install'), "install"),
            (translator.t('buttons.clean'), "clean"),
            (translator.t('buttons.install_vci'), "vci"),
            (translator.t('buttons.launch'), "launch"),
            (translator.t('buttons.kill_process'), "kill"),
        ]
        for i, (txt, action) in enumerate(btns):
            b = QtWidgets.QPushButton(txt)
            b.setMinimumHeight(44)
            b.setObjectName("actionButton")
            if action == "download":
                self.download_button = b
                b.clicked.connect(self.download_diagbox)
            elif action == "install":
                self.install_button = b
                b.clicked.connect(self.install_diagbox)
            elif action == "clean":
                b.clicked.connect(self.clean_diagbox)
            elif action == "vci":
                b.clicked.connect(self.install_vci_driver)
            elif action == "launch":
                b.clicked.connect(self.launch_diagbox)
            elif action == "kill":
                b.clicked.connect(self.kill_diagbox)
            grid.addWidget(b, i//3, i%3)

        right.addLayout(grid)

        # Pause and Cancel buttons (hidden by default)
        buttons_row = QtWidgets.QHBoxLayout()
        
        self.pause_button = QtWidgets.QPushButton(translator.t('buttons.pause'))
        self.pause_button.setMinimumHeight(44)
        self.pause_button.setObjectName("actionButton")
        self.pause_button.setStyleSheet("background-color: #f0ad4e; color: white;")
        self.pause_button.clicked.connect(self.toggle_pause_download)
        self.pause_button.setVisible(False)
        buttons_row.addWidget(self.pause_button)
        
        self.cancel_button = QtWidgets.QPushButton(translator.t('buttons.cancel'))
        self.cancel_button.setMinimumHeight(44)
        self.cancel_button.setObjectName("actionButton")
        self.cancel_button.setStyleSheet("background-color: #d9534f; color: white;")
        self.cancel_button.clicked.connect(self.cancel_download)
        self.cancel_button.setVisible(False)
        buttons_row.addWidget(self.cancel_button)
        
        right.addLayout(buttons_row)
        
        layout.addLayout(sub)
        layout.addLayout(right)

        # Disable install button if a Diagbox version is already installed
        try:
            installed_version = self.check_installed_version()
            if installed_version and hasattr(self, 'install_button'):
                self.install_button.setEnabled(False)
                self.install_button.setToolTip(translator.t('messages.install.must_clean_tooltip', installed=installed_version))
            elif hasattr(self, 'install_button'):
                self.install_button.setEnabled(True)
                self.install_button.setToolTip("")
        except Exception:
            pass

        return w

    def page_update(self):
        w = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(w)
        layout.addWidget(QtWidgets.QLabel(translator.t('messages.page_update')))
        layout.addStretch()
        return w

    def page_about(self):
        w = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(w)
        layout.setSpacing(10)
        
        # Top section with logo/version (left) and changelog (right)
        top_section = QtWidgets.QWidget()
        top_layout = QtWidgets.QHBoxLayout(top_section)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(10)
        
        # Left: Logo and version
        left_section = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_section)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        logo = QtWidgets.QLabel()
        pix = QtGui.QPixmap(str(BASE / "icons" / "logo.png"))
        if not pix.isNull():
            pix = pix.scaledToWidth(160, QtCore.Qt.TransformationMode.SmoothTransformation)
            logo.setPixmap(pix)
        else:
            logo.setText("Logo non disponible")
        left_layout.addWidget(logo)
        
        version_label = QtWidgets.QLabel(translator.t('labels.version', version=APP_VERSION))
        version_label.setStyleSheet("font-size: 14px;")
        left_layout.addWidget(version_label)
        left_layout.addStretch()
        
        top_layout.addWidget(left_section)
        
        # Right: Changelog frame
        changelog_frame = QtWidgets.QFrame()
        changelog_frame.setObjectName("changelogFrame")
        changelog_frame.setStyleSheet("""
            #changelogFrame {
                background-color: #2a2a2a;
                border: 1px solid #3a3a3a;
                border-radius: 6px;
                padding: 10px;
            }
        """)
        changelog_frame.setMinimumWidth(300)
        changelog_frame.setMaximumWidth(400)
        changelog_layout = QtWidgets.QVBoxLayout(changelog_frame)
        changelog_layout.setContentsMargins(10, 10, 10, 10)
        changelog_layout.setSpacing(8)
        
        changelog_title = QtWidgets.QLabel("📋 Changelog")
        changelog_title.setStyleSheet("font-size: 14px; font-weight: bold; color: #ffffff;")
        changelog_layout.addWidget(changelog_title)
        
        self.changelog_text = QtWidgets.QTextEdit()
        self.changelog_text.setReadOnly(True)
        self.changelog_text.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #d4d4d4;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 11px;
                border: 1px solid #3a3a3a;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        self.changelog_text.setPlainText("Loading changelog...")
        self.changelog_text.setMinimumHeight(150)
        changelog_layout.addWidget(self.changelog_text)
        
        top_layout.addWidget(changelog_frame)
        
        layout.addWidget(top_section)
        
        # Load changelog asynchronously
        QtCore.QTimer.singleShot(500, self.load_changelog)
        
        # Toggle console button
        self.toggle_log_btn = QtWidgets.QPushButton(translator.t('buttons.hide_console'))
        self.toggle_log_btn.setObjectName("actionButton")
        self.toggle_log_btn.setFixedHeight(44)
        self.toggle_log_btn.clicked.connect(self.toggle_console)
        layout.addWidget(self.toggle_log_btn)

        # Open logs button
        self.open_log_btn = QtWidgets.QPushButton(translator.t('buttons.open_log'))
        self.open_log_btn.setObjectName("actionButton")
        self.open_log_btn.setFixedHeight(44)
        self.open_log_btn.clicked.connect(self.open_logs)
        layout.addWidget(self.open_log_btn)
        
        # Console widget (hidden by default)
        self.log_widget = QtWidgets.QTextEdit()
        self.log_widget.setReadOnly(True)
        self.log_widget.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #d4d4d4;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                border: 1px solid #3a3a3a;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        self.log_widget.setVisible(True)
        self.log_widget.setMinimumHeight(200)
        layout.addWidget(self.log_widget)
        
        # Add logging handler for this widget (only INFO+ so debug lines stay in file)
        self.log_handler = QTextEditLogger(self.log_widget)
        self.log_handler.setLevel(logging.INFO)
        self.log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(self.log_handler)
        # Populate the UI console with the existing log file contents so the
        # in-app view matches the file (startup logs are written before the
        # QTextEdit handler is attached). Filter out DEBUG lines to avoid
        # duplicating verbose debug entries in the UI.
        try:
            if log_file.exists():
                raw = log_file.read_text(encoding='utf-8')
                if raw:
                    # Remove lines that are DEBUG-level entries
                    filtered_lines = [l for l in raw.splitlines() if ' - DEBUG - ' not in l]
                    filtered = "\n".join(filtered_lines)
                    self.log_widget.setPlainText(filtered)
        except Exception as e:
            # Use debug here (file may be unreadable) but UI handler ignores DEBUG
            logger.debug(f"Failed to populate UI console from log file: {e}")
        
        layout.addStretch()
        return w
    
    def load_changelog(self):
        """Load changelog from GitHub release for current app version"""
        try:
            # Build GitHub API URL for the release matching current version
            version_tag = f"v{APP_VERSION}"
            api_url = f"https://api.github.com/repos/RetroGameSets/PSA-DIAG/releases/tags/{version_tag}"
            
            logger.info(f"Fetching changelog from: {api_url}")
            response = requests.get(api_url, timeout=5)
            
            if response.status_code == 200:
                data = response.json()
                body = data.get('body', '')
                
                if body:
                    # Extract the changelog from the body
                    # The body format from build.yml is:
                    # ## PSA-DIAG x.x.x.x
                    # ### Installation
                    # ...
                    # ### Changes
                    # - commit message
                    
                    # Try to extract just the Changes section
                    if '### Changes' in body:
                        changes_section = body.split('### Changes', 1)[1].strip()
                        # Remove any trailing markdown sections
                        if '###' in changes_section:
                            changes_section = changes_section.split('###', 1)[0].strip()
                        changelog_content = changes_section
                    else:
                        # Fallback: use entire body
                        changelog_content = body
                    
                    self.changelog_text.setPlainText(changelog_content)
                    logger.info("Changelog loaded successfully")
                else:
                    self.changelog_text.setPlainText("No changelog available for this version.")
            elif response.status_code == 404:
                self.changelog_text.setPlainText(f"Release {version_tag} not found on GitHub.\n\nThis may be a development version.")
                logger.warning(f"Release {version_tag} not found (404)")
            else:
                self.changelog_text.setPlainText(f"Failed to load changelog (HTTP {response.status_code})")
                logger.warning(f"Failed to fetch changelog: HTTP {response.status_code}")
                
        except Exception as e:
            logger.error(f"Error loading changelog: {e}", exc_info=True)
            self.changelog_text.setPlainText(f"Error loading changelog:\n{str(e)}")
    
    def toggle_console(self):
        """Toggle console visibility"""
        if self.log_widget:
            is_visible = self.log_widget.isVisible()
            self.log_widget.setVisible(not is_visible)
            
            if self.toggle_log_btn:
                self.toggle_log_btn.setText(translator.t('buttons.hide_console') if not is_visible else translator.t('buttons.show_console'))
            
            # Adjust window height
            if not is_visible:
                self.resize(900, 500)  # Expanded height
            else:
                self.resize(900, 420)  # Original height
    def open_logs(self):
        """Open the logs folder and select the most recent log file if present."""
        try:
            logs_dir = log_folder
            if not logs_dir.exists():
                QtWidgets.QMessageBox.information(self, translator.t('app.title'), translator.t('messages.log.open_failed'))
                return

            log_files = sorted(list(logs_dir.glob('psa_diag_*.log')), key=lambda p: p.stat().st_mtime, reverse=True)
            if log_files:
                latest_log = log_files[0]
                try:
                    if sys.platform == 'win32':
                        os.startfile(str(latest_log))
                    elif sys.platform == 'darwin':
                        subprocess.run(['open', str(latest_log)])
                    else:
                        subprocess.run(['xdg-open', str(latest_log)])
                except Exception:
                    # Fallback: open folder
                    if sys.platform == 'win32':
                        os.startfile(str(logs_dir))
                    elif sys.platform == 'darwin':
                        subprocess.run(['open', str(logs_dir)])
                    else:
                        subprocess.run(['xdg-open', str(logs_dir)])
            else:
                QtWidgets.QMessageBox.information(self, translator.t('app.title'), translator.t('messages.log.open_failed'))
        except Exception as e:
            logger.error(f"Failed to open logs: {e}", exc_info=True)
            QtWidgets.QMessageBox.warning(self, translator.t('app.title'), translator.t('messages.log.open_failed'))

    def _banner_download_clicked(self):
        """Handler for banner Download button: select the latest available
        version in the combo (if present) and invoke the same download flow.
        """
        try:
            latest = self.get_latest_available_version()
            if not latest:
                return

            # Try to find the version in the combo and select it
            if hasattr(self, 'version_combo'):
                found_index = -1
                for i in range(self.version_combo.count()):
                    data = self.version_combo.itemData(i)
                    if data and data[0] == latest:
                        found_index = i
                        break
                if found_index >= 0:
                    self.version_combo.setCurrentIndex(found_index)
                else:
                    # If not present, add it as a temporary entry (use archive.org fallback URL)
                    url = None
                    for display_name, version, vurl in getattr(self, 'version_options', []):
                        if version == latest:
                            url = vurl
                            break
                    if not url:
                        url = f"https://archive.org/download/psa-diag.fr/Diagbox_Install_{latest}.7z"
                    self.version_combo.addItem(latest, userData=(latest, url))
                    self.version_combo.setCurrentIndex(self.version_combo.count() - 1)

            # Ensure last_version_diagbox is set to the latest
            self.last_version_diagbox = latest

            # Hide the banner now that user requested the download
            if hasattr(self, 'install_banner'):
                self.install_banner.setVisible(False)

            # Restore headers so user still sees installed/latest info
            if hasattr(self, 'header_installed'):
                try:
                    self.header_installed.setVisible(True)
                except Exception:
                    pass
            if hasattr(self, 'header_online'):
                try:
                    self.header_online.setVisible(True)
                except Exception:
                    pass

            # Trigger the normal download flow
            self.download_diagbox()
        except Exception as e:
            logger.error(f"Error handling banner download click: {e}", exc_info=True)

    def _dismiss_install_banner(self):
        """Hide the install banner and restore the installed/latest headers."""
        try:
            if hasattr(self, 'install_banner'):
                self.install_banner.setVisible(False)
            if hasattr(self, 'header_installed'):
                try:
                    self.header_installed.setVisible(True)
                except Exception:
                    pass
            if hasattr(self, 'header_online'):
                try:
                    self.header_online.setVisible(True)
                except Exception:
                    pass
        except Exception as e:
            logger.debug(f"Error dismissing install banner: {e}")
    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.MouseButton.LeftButton:
            # Check if click is on a widget that should not trigger window drag
            widget = self.childAt(event.pos())
            if widget is not None:
                # Check the widget and all its parents
                current = widget
                while current is not None and current != self:
                    if isinstance(current, (QtWidgets.QComboBox, QtWidgets.QPushButton, 
                                          QtWidgets.QCheckBox, QtWidgets.QProgressBar,
                                          QtWidgets.QLineEdit, QtWidgets.QAbstractItemView)):
                        self.dragPos = None  # Disable dragging for interactive widgets
                        event.ignore()
                        return
                    current = current.parent()
            
            self.dragPos = event.globalPosition().toPoint()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == QtCore.Qt.MouseButton.LeftButton:
            # Don't move if dragPos is None (clicked on interactive widget)
            if not hasattr(self, 'dragPos') or self.dragPos is None:
                return
                
            # Don't move window if a combo box popup is open
            for combo in self.findChildren(QtWidgets.QComboBox):
                if combo.view().isVisible():
                    return
            
            delta = event.globalPosition().toPoint() - self.dragPos
            self.move(self.pos() + delta)
            self.dragPos = event.globalPosition().toPoint()
            event.accept()
    
    def mouseReleaseEvent(self, event):
        """Reset drag position on mouse release"""
        if event.button() == QtCore.Qt.MouseButton.LeftButton:
            self.dragPos = None
            event.accept()
    
    def closeEvent(self, event):
        """Handle application close - stop any running processes"""
        # Stop download thread if running
        if self.download_thread and self.download_thread.isRunning():
            self.download_thread.cancel()
            self.download_thread.wait(1000)  # Wait max 1 second
        
        # Stop installation thread if running
        if self.install_thread and self.install_thread.isRunning():
            self.install_thread.stop()
            self.install_thread.wait(2000)  # Wait max 2 seconds
        
        event.accept()


if __name__ == "__main__":
    # Hide console window first
    hide_console()
    # Terminate any leftover updater helpers that might still be running
    try:
        kill_updater_processes()
    except Exception as e:
        logger.debug(f"kill_updater_processes failed: {e}")
    
    # Check if running as admin, if not relaunch with admin privileges
    if not is_admin():
        logger.warning("Not running as admin, requesting elevation...")
        if run_as_admin():
            logger.info("Admin elevation requested, exiting current instance")
            sys.exit(0)  # Exit current instance
        else:
            # If elevation failed, continue anyway (user might have cancelled)
            logger.warning("Continuing without admin privileges (some features may not work)")
    
    logger.info("Creating QApplication")
    app = QtWidgets.QApplication([])
    logger.info("Creating MainWindow")
    win = MainWindow()
    win.show()
    logger.info("Application started successfully")
    sys.exit(app.exec())
