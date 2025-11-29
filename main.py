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
import subprocess
import logging
from datetime import datetime
import json

# Determine base path for resources.
if getattr(sys, 'frozen', False):
    BASE = Path(sys._MEIPASS)
else:
    BASE = Path(__file__).resolve().parent

# Persistent config directory (where we save preferences). Use APPDATA on Windows 
CONFIG_DIR = Path(os.getenv('APPDATA', Path.home())) / 'PSA_DIAG'
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
APP_VERSION = "2.1.0.2"
URL_LAST_VERSION_PSADIAG = "https://psa-diag.fr/diagbox/install/last_version_psadiag.json"
URL_LAST_VERSION_DIAGBOX = "https://psa-diag.fr/diagbox/install/last_version_diagbox.json"

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
            icon = QtGui.QIcon(str(icon_path))
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
            
            # Use 7z.exe for much faster extraction
            seven_zip_exe = BASE / "tools" / "7z.exe"
            
            if not seven_zip_exe.exists():
                logger.error(f"7z.exe not found at {seven_zip_exe}")
                self.finished.emit(False, f"7z.exe not found at {seven_zip_exe}")
                return
            
            logger.info("Starting 7z extraction...")
            self.progress.emit(0)
            
            try:
                # Run 7z.exe with real-time output
                # -bsp1 = show progress in stdout
                self.process = subprocess.Popen(
                    [str(seven_zip_exe), "x", self.path, "-oC:\\", "-y", "-bsp1"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    bufsize=1,
                    universal_newlines=True,
                    creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
                )
                
                # Read output in real-time to get progress
                while True:
                    output = self.process.stdout.readline()
                    if output == '' and self.process.poll() is not None:
                        break
                    if output:
                        # Log the output for debugging
                        logger.debug(f"7z output: {output.strip()}")
                        
                        # 7z outputs progress like "1% 2909 - filename"
                        output = output.strip()
                        if '%' in output and ' - ' in output:
                            try:
                                # Extract percentage from beginning of line
                                percent_str = output.split('%')[0].strip()
                                if percent_str.isdigit():
                                    percent = int(percent_str)
                                    logger.debug(f"Progress extracted: {percent}%")
                                    self.progress.emit(percent)
                                
                                # Extract filename after " - "
                                filename = output.split(' - ', 1)[1] if ' - ' in output else ''
                                if filename:
                                    self.file_progress.emit(filename)
                            except Exception as e:
                                logger.warning(f"Error parsing progress: {e}")
                                pass
                
                # Get return code
                return_code = self.process.poll()
                stderr = self.process.stderr.read()
                
                if return_code != 0:
                    if "permission" in stderr.lower() or "access" in stderr.lower():
                        extraction_errors.append("Some files skipped due to permission errors")
                    elif stderr:
                        extraction_errors.append(f"7z error: {stderr[:200]}")
                
            except Exception as e:
                extraction_errors.append(f"Extraction error: {str(e)}")
            
            self.progress.emit(100)
            
            # Build result message
            if extraction_errors:
                error_summary = "\n".join(extraction_errors)
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
    """Thread for cleaning Diagbox folders and shortcuts"""
    finished = QtCore.Signal(bool, str, int)  # success, message, success_count
    progress = QtCore.Signal(int, int)  # current, total
    item_progress = QtCore.Signal(str)  # current item being deleted

    def __init__(self, folders, shortcuts):
        super().__init__()
        self.folders = folders
        self.shortcuts = shortcuts
        self.failed_items = []

    def run(self):
        total_items = len(self.folders) + len(self.shortcuts)
        current_item = 0
        success_count = 0
        
        # Delete folders
        for folder in self.folders:
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
        
        # Version options: (display_name, version, url)
        self.version_options = [
            (f"Diagbox {self.last_version_diagbox} (Latest)", self.last_version_diagbox, f"https://archive.org/download/psa-diag.fr/Diagbox_Install_{self.last_version_diagbox}.7z"),
            ("Diagbox 9.85", "9.85", "https://archive.org/download/psa-diag.fr/Diagbox_Install_9.85.7z")
        ]

        # Connect signals
        self.download_finished.connect(self.on_download_finished)

        self.setup_ui()
        
        # Check for app updates after UI is ready
        QtCore.QTimer.singleShot(1000, self.check_app_update)

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
                    content = f.read()
                    for line in content.splitlines():
                        if '=' in line:
                            return line.split('=', 1)[1].strip()
            except Exception as e:
                logger.error(f"Error reading language file: {e}")
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
                if file.startswith("Diagbox_Install_") and file.endswith(".7z"):
                    # Extract version from filename
                    version = file.replace("Diagbox_Install_", "").replace(".7z", "")
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
        # Re-enable all buttons and combo box FIRST
        self.set_buttons_enabled(True)
        
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
        
        # Show message box AFTER re-enabling buttons
        QtWidgets.QMessageBox.information(self, translator.t('messages.install.title'), message)
        
        # Reset footer after a delay
        QtCore.QTimer.singleShot(3000, self.reset_footer)
    
    def reset_footer(self):
        """Reset footer to ready state"""
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(translator.t('labels.ready'))
        if hasattr(self, 'footer_progress'):
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

    def install_diagbox(self):
        logger.info("Install Diagbox initiated")
        # Get selected version from combo box
        if hasattr(self, 'version_combo'):
            selected_data = self.version_combo.currentData()
            if selected_data:
                version, url = selected_data
                self.last_version_diagbox = version
                self.diagbox_path = os.path.join(self.download_folder, f"Diagbox_Install_{version}.7z")
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
        shortcut_names = [
            "Diagbox Language Changer.lnk",
            "Diagbox.lnk",
            "PSA Interface Checker.lnk",
            "Terminate Diagbox Processes.lnk"
        ]
        
        for shortcut in shortcut_names:
            shortcut_path = os.path.join(public_desktop, shortcut)
            if os.path.exists(shortcut_path):
                shortcuts_to_delete.append(shortcut_path)
        
        if not folders_to_delete and not shortcuts_to_delete:
            QtWidgets.QMessageBox.information(
                self,
                translator.t('messages.clean.title'),
                translator.t('messages.clean.nothing_to_clean')
            )
            return
        
        # Confirm deletion
        items_list = []
        if folders_to_delete:
            items_list.append("Folders:")
            items_list.extend([f"- {folder}" for folder in folders_to_delete])
        if shortcuts_to_delete:
            if items_list:
                items_list.append("")
            items_list.append("Shortcuts:")
            items_list.extend([f"- {os.path.basename(s)}" for s in shortcuts_to_delete])
        
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
            self.footer_progress.setRange(0, total_items)
            self.footer_progress.setValue(0)
        if hasattr(self, 'footer_label'):
            self.footer_label.setText("Cleaning Diagbox...")
        
        # Start cleaning in thread
        self.clean_thread = CleanThread(folders_to_delete, shortcuts_to_delete)
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
            total_items = self.footer_progress.maximum()
            self.footer_progress.setValue(total_items)
        
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
        """Install VCI Driver using ACTIAPnPInstaller.exe"""
        logger.info("Install VCI Driver initiated")
        driver_path = r"C:\AWRoot\extra\Drivers\xsevo\ACTIAPnPInstaller.exe"
        
        # Check if the installer exists
        if not os.path.exists(driver_path):
            logger.error(f"VCI Driver installer not found: {driver_path}")
            QtWidgets.QMessageBox.warning(
                self,
                translator.t('messages.vci_driver.title'),
                translator.t('messages.vci_driver.not_found', path=driver_path)
            )
            return
        
        try:
            # Run the installer with /nodisplay flag
            result = subprocess.run(
                [driver_path, "/nodisplay"],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
            )
            
            # Return code 0 = success, 6 = already installed or success with warning
            if result.returncode == 0 or result.returncode == 6:
                QtWidgets.QMessageBox.information(
                    self,
                    translator.t('messages.vci_driver.title'),
                    translator.t('messages.vci_driver.success')
                )
            else:
                QtWidgets.QMessageBox.warning(
                    self,
                    translator.t('messages.vci_driver.title'),
                    translator.t('messages.vci_driver.warning', code=result.returncode)
                )
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                translator.t('messages.vci_driver.title'),
                translator.t('messages.vci_driver.error', error=str(e))
            )

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
            self.diagbox_path = os.path.join(self.download_folder, f"Diagbox_Install_{self.last_version_diagbox}.7z")
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
                        logger.info("User accepted update, opening download page")
                        import webbrowser
                        webbrowser.open("https://github.com/RetroGameSets/PSA-DIAG/releases/latest")
                else:
                    logger.info("App is up to date")
        except Exception as e:
            # Silently fail if update check fails (no internet, server down, etc.)
            logger.warning(f"Update check failed: {e}")

    def download_diagbox(self):
        logger.info("Download Diagbox button clicked")
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
        
        self.diagbox_path = os.path.join(self.download_folder, f"Diagbox_Install_{self.last_version_diagbox}.7z")
        
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

    def switch_page(self, index, button):
        # uncheck all sibling buttons in sidebar
        sidebar = self.findChild(QtWidgets.QFrame, "sidebar")
        if sidebar:
            for child in sidebar.findChildren(SidebarButton):
                child.setChecked(False)
        button.setChecked(True)
        self.stack.setCurrentIndex(index)

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
        recheck = QtWidgets.QPushButton(translator.t('buttons.recheck'))
        recheck.setFixedWidth(160)
        recheck.clicked.connect(self.check_system)
        layout.addWidget(recheck, 0, QtCore.Qt.AlignmentFlag.AlignLeft)

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

        header_online = QtWidgets.QLabel(translator.t('labels.last_version', version=self.last_version_diagbox if self.last_version_diagbox else 'Unknown'))
        header_online.setObjectName("sectionHeader")
        layout.addWidget(header_online)

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
        
        # Language options: (display_name, lang_code)
        languages = [
            (translator.t('languages.en_GB'), "en_GB"),
            (translator.t('languages.fr_FR'), "fr_FR"),
            (translator.t('languages.it_IT'), "it_IT"),
            (translator.t('languages.nl_NL'), "nl_NL"),
            (translator.t('languages.pl_PL'), "pl_PL"),
            (translator.t('languages.pt_PT'), "pt_PT"),
            (translator.t('languages.ru_RU'), "ru_RU"),
            (translator.t('languages.tr_TR'), "tr_TR"),
            (translator.t('languages.sv_SE'), "sv_SE"),
            (translator.t('languages.da_DK'), "da_DK"),
            (translator.t('languages.cs_CZ'), "cs_CZ"),
            (translator.t('languages.de_DE'), "de_DE"),
            (translator.t('languages.el_GR'), "el_GR"),
            (translator.t('languages.hr_HR'), "hr_HR"),
            (translator.t('languages.zh_CN'), "zh_CN"),
            (translator.t('languages.ja_JP'), "ja_JP"),
            (translator.t('languages.es_ES'), "es_ES"),
            (translator.t('languages.sl_SI'), "sl_SI"),
            (translator.t('languages.hu_HU'), "hu_HU"),
            (translator.t('languages.fi_FI'), "fi_FI"),
        ]
        
        for display_name, lang_code in languages:
            self.language_combo.addItem(display_name, userData=lang_code)
        
        # Set current language if available
        current_lang = self.get_diagbox_language()
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
        
        # Top section with logo and version
        top_section = QtWidgets.QWidget()
        top_layout = QtWidgets.QHBoxLayout(top_section)
        top_layout.setContentsMargins(0, 0, 0, 0)
        
        logo = QtWidgets.QLabel()
        pix = QtGui.QPixmap(str(BASE / "icons" / "logo.png"))
        if not pix.isNull():
            pix = pix.scaledToWidth(160, QtCore.Qt.TransformationMode.SmoothTransformation)
            logo.setPixmap(pix)
        else:
            logo.setText("Logo non disponible")
        top_layout.addWidget(logo)
        
        version_label = QtWidgets.QLabel(translator.t('labels.version', version=APP_VERSION))
        version_label.setStyleSheet("font-size: 14px;")
        top_layout.addWidget(version_label)
        top_layout.addStretch()
        
        layout.addWidget(top_section)
        
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
        
        # Add logging handler for this widget
        self.log_handler = QTextEditLogger(self.log_widget)
        self.log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(self.log_handler)
        
        layout.addStretch()
        return w
    
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
            # Determine latest log file
            logs_dir = log_folder
            log_files = sorted(logs_dir.glob('psa_diag_*.log'), key=lambda p: p.stat().st_mtime, reverse=True)
            if log_files:
                latest = log_files[0]
                # On Windows, open Explorer and select the file
                if sys.platform == 'win32':
                    try:
                        subprocess.run(['explorer', f'/select,{str(latest)}'])
                        return
                    except Exception:
                        pass
                # Fallback: open folder or file using start/open
                try:
                    if sys.platform == 'win32':
                        os.startfile(str(latest))
                    elif sys.platform == 'darwin':
                        subprocess.run(['open', str(latest)])
                    else:
                        subprocess.run(['xdg-open', str(logs_dir)])
                    return
                except Exception:
                    pass

            # No logs found - try to open the folder instead
            if sys.platform == 'win32':
                os.startfile(str(logs_dir))
            elif sys.platform == 'darwin':
                subprocess.run(['open', str(logs_dir)])
            else:
                subprocess.run(['xdg-open', str(logs_dir)])
        except Exception as e:
            logger.error(f"Failed to open logs: {e}")
            QtWidgets.QMessageBox.warning(self, translator.t('app.title'), translator.t('messages.log.open_failed'))
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
