"""
PSA-DIAG FREE
"""
import sys
from pathlib import Path
from PySide6 import QtCore, QtGui, QtWidgets #type:ignore
import psutil #type:ignore
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
from config import CONFIG_DIR, APP_VERSION, URL_LAST_VERSION_PSADIAG, URL_LAST_VERSION_DIAGBOX, URL_VERSION_OPTIONS, URL_REMOTE_MESSAGES, ARCHIVE_PASSWORD, URL_VHD_DOWNLOAD, URL_VHD_TORRENT

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
logger.info("[STEP 0] -- Initializing....")
logger.info(f"PSA-DIAG v{APP_VERSION}")
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
    """Terminate any leftover updater.exe and aria2c.exe processes from previous runs."""
    try:
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                name = (proc.info.get('name') or '').lower()
                exe = proc.info.get('exe') or ''
                # Kill both updater.exe and aria2c.exe
                if name in ['updater.exe', 'aria2c.exe'] or (exe and os.path.basename(exe).lower() in ['updater.exe', 'aria2c.exe']):
                    logger.info(f"Terminating leftover {name} PID={proc.pid}")
                    try:
                        proc.terminate()
                        proc.wait(timeout=2)
                    except Exception:
                        try:
                            proc.kill()
                        except Exception:
                            logger.debug(f"Failed to kill {name} PID={proc.pid}")
            except Exception as e:
                logger.debug(f"Error while checking process: {e}")
    except Exception as e:
        logger.debug(f"Failed to enumerate processes to kill updater.exe/aria2c.exe: {e}")

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
            # Define icon size
            icon_size = QtCore.QSize(48, 48)
            
            # Render SVG at target size for crisp display
            try:
                from PySide6.QtSvg import QSvgRenderer #type:ignore
                
                # Render original (colored) SVG
                renderer = QSvgRenderer(str(icon_path))
                orig_pix = QtGui.QPixmap(icon_size)
                orig_pix.fill(QtCore.Qt.transparent)
                painter = QtGui.QPainter(orig_pix)
                painter.setRenderHint(QtGui.QPainter.Antialiasing)
                painter.setRenderHint(QtGui.QPainter.SmoothPixmapTransform)
                renderer.render(painter)
                painter.end()
                
                # Create white-tinted version for unchecked state
                white_pix = QtGui.QPixmap(orig_pix.size())
                white_pix.fill(QtCore.Qt.transparent)
                p = QtGui.QPainter(white_pix)
                p.setRenderHints(QtGui.QPainter.Antialiasing | QtGui.QPainter.SmoothPixmapTransform)
                p.drawPixmap(0, 0, orig_pix)
                p.setCompositionMode(QtGui.QPainter.CompositionMode_SourceIn)
                p.fillRect(white_pix.rect(), QtGui.QColor('white'))
                p.end()
                
            except Exception:
                # Fallback: load as regular pixmap and scale
                orig_pix = QtGui.QPixmap(str(icon_path))
                orig_pix = orig_pix.scaled(icon_size, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                
                white_pix = QtGui.QPixmap(orig_pix.size())
                white_pix.fill(QtCore.Qt.transparent)
                p = QtGui.QPainter(white_pix)
                p.setRenderHints(QtGui.QPainter.Antialiasing | QtGui.QPainter.SmoothPixmapTransform)
                p.drawPixmap(0, 0, orig_pix)
                p.setCompositionMode(QtGui.QPainter.CompositionMode_SourceIn)
                p.fillRect(white_pix.rect(), QtGui.QColor('white'))
                p.end()

            icon = QtGui.QIcon()
            # Off = not checked -> white icon
            icon.addPixmap(white_pix, QtGui.QIcon.Normal, QtGui.QIcon.Off)
            # On = checked -> original (colored) icon
            icon.addPixmap(orig_pix, QtGui.QIcon.Normal, QtGui.QIcon.On)
            self.setIcon(icon)
            self.setIconSize(icon_size)
        # Center the icon by removing text and ensuring proper alignment
        self.setText("")
class SplashScreen(QtWidgets.QWidget):
    """Modern stylized loading splash screen with animation"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(QtCore.Qt.WindowType.FramelessWindowHint | QtCore.Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # Set size
        self.setGeometry(0, 0, 900, 550)
        
        # Center on screen
        screen_geometry = QtGui.QGuiApplication.primaryScreen().availableGeometry()
        x = (screen_geometry.width() - 900) // 2
        y = (screen_geometry.height() - 550) // 2
        self.move(x, y)
        
        # Animation counter for spinner effect
        self.animation_counter = 0
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.update_animation)
        self.timer.start(50)
        
        # Create layout
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Create main widget
        main_widget = QtWidgets.QWidget()
        main_layout = QtWidgets.QVBoxLayout(main_widget)
        main_layout.setSpacing(30)
        main_layout.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        main_layout.setContentsMargins(50, 50, 50, 50)
        
        # App title
        title_label = QtWidgets.QLabel("PSA-DIAG")
        title_font = QtGui.QFont()
        title_font.setPointSize(36)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: white;")
        title_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # Spinner (animated dots)
        self.spinner_label = QtWidgets.QLabel()
        spinner_font = QtGui.QFont()
        spinner_font.setPointSize(28)
        self.spinner_label.setFont(spinner_font)
        self.spinner_label.setStyleSheet("color: #00A8E1; font-weight: bold;")
        self.spinner_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.spinner_label)
        
        # Status text
        status_label = QtWidgets.QLabel("Loading... Please Wait")
        status_font = QtGui.QFont()
        status_font.setPointSize(12)
        status_label.setFont(status_font)
        status_label.setStyleSheet("color: #CCCCCC;")
        status_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(status_label)
        
        # Progress bar (indeterminate style)
        progress_bar = QtWidgets.QProgressBar()
        progress_bar.setMaximum(0)  # Makes it indeterminate
        progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 5px;
                background-color: rgba(255, 255, 255, 30);
                height: 6px;
            }
            QProgressBar::chunk {
                background-color: #00A8E1;
                border-radius: 5px;
            }
        """)
        progress_bar.setMaximumWidth(300)
        progress_layout = QtWidgets.QHBoxLayout()
        progress_layout.addStretch()
        progress_layout.addWidget(progress_bar)
        progress_layout.addStretch()
        main_layout.addLayout(progress_layout)
        
        layout.addWidget(main_widget)
    
    def update_animation(self):
        """Update spinner animation"""
        self.animation_counter = (self.animation_counter + 1) % 4
        spinner_chars = ["⠋", "⠙", "⠹", "⠸"]
        self.spinner_label.setText(spinner_chars[self.animation_counter])
    
    def paintEvent(self, event):
        """Draw gradient background"""
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        
        # Create gradient with reduced transparency (more opaque)
        gradient = QtGui.QLinearGradient(0, 0, self.width(), self.height())
        gradient.setColorAt(0.0, QtGui.QColor(20, 20, 30, 245))
        gradient.setColorAt(0.5, QtGui.QColor(30, 30, 45, 245))
        gradient.setColorAt(1.0, QtGui.QColor(20, 20, 30, 245))
        
        # Draw rounded rectangle with gradient
        path = QtGui.QPainterPath()
        path.addRoundedRect(QtCore.QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5), 20, 20)
        painter.fillPath(path, gradient)
        
        # Draw elegant border with rounded corners
        painter.setPen(QtGui.QPen(QtGui.QColor(0, 168, 225, 180), 2))
        painter.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), 20, 20)
        
        painter.end()
    
    def closeEvent(self, event):
        """Stop animation timer when closing"""
        self.timer.stop()
        super().closeEvent(event)


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
            logger.info(f"Starting download")
            response = requests.get(self.url, stream=True, timeout=30)
            
            # Check for HTTP errors with user-friendly messages
            if response.status_code != 200:
                status_code = response.status_code
                logger.error(f"Download failed with HTTP {status_code}")
                
                # Generate user-friendly error message based on status code
                if status_code == 404:
                    error_msg = translator.t('messages.download.error_404')
                elif status_code == 502:
                    error_msg = translator.t('messages.download.error_502')
                elif status_code == 503:
                    error_msg = translator.t('messages.download.error_503')
                elif status_code == 403:
                    error_msg = translator.t('messages.download.error_403')
                elif status_code == 500:
                    error_msg = translator.t('messages.download.error_500')
                elif 400 <= status_code < 500:
                    error_msg = translator.t('messages.download.error_4xx', code=status_code)
                elif 500 <= status_code < 600:
                    error_msg = translator.t('messages.download.error_5xx', code=status_code)
                else:
                    error_msg = translator.t('messages.download.error_generic', code=status_code)
                
                self.finished.emit(False, error_msg)
                return
            
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
        except requests.exceptions.Timeout:
            logger.error(f"Download timeout")
            error_msg = translator.t('messages.download.error_timeout')
            self.finished.emit(False, error_msg)
        except requests.exceptions.ConnectionError:
            logger.error(f"Download connection error")
            error_msg = translator.t('messages.download.error_connection')
            self.finished.emit(False, error_msg)
        except Exception as e:
            logger.error(f"Download exception: {e}", exc_info=True)
            error_msg = translator.t('messages.download.error_generic', code=str(e))
            self.finished.emit(False, error_msg)

class VHDXDownloadThread(QtCore.QThread):
    """Thread for downloading VHDX files"""
    finished = QtCore.Signal(bool, str)  # success, message
    progress = QtCore.Signal(int, float, str)  # progress (0-1000), speed (MB/s), ETA
    
    def __init__(self, url, destination_folder, destination_drive):
        super().__init__()
        self.url = url
        self.destination_folder = destination_folder
        self.destination_drive = destination_drive
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
            # Create VHD folder at root of selected drive
            vhd_folder = os.path.join(f"{self.destination_drive}:\\", "VHD")
            os.makedirs(vhd_folder, exist_ok=True)
            logger.info(f"VHD folder created/verified: {vhd_folder}")
            
            # Extract filename from URL or use default
            filename = os.path.basename(self.url) if self.url else "PSA-DIAG-Dynamic.vhdx"
            if not filename or '.' not in filename:
                filename = "PSA-DIAG-Dynamic.vhdx"
            
            file_path = os.path.join(vhd_folder, filename)
            logger.info(f"Starting VHDX download to: {file_path}")
            logger.info(f"Download URL: {self.url}")
            
            # Start download with streaming
            response = requests.get(self.url, stream=True, timeout=30)
            
            # Check for HTTP errors
            if response.status_code != 200:
                logger.error(f"Download failed with HTTP {response.status_code}")
                self.finished.emit(False, f"Erreur HTTP {response.status_code}")
                return
            
            response.raise_for_status()
            
            # Get content length
            total_size = 0
            content_length = response.headers.get('content-length')
            if content_length:
                total_size = int(content_length)
                logger.info(f"VHDX size: {total_size / (1024**3):.2f} GB")
            
            downloaded = 0
            chunk_count = 0
            start_time = time.time()
            
            with open(file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    # Check if paused
                    while self._is_paused and not self._is_cancelled:
                        time.sleep(0.1)
                    
                    if self._is_cancelled:
                        f.close()
                        if os.path.exists(file_path):
                            os.remove(file_path)
                        logger.warning("VHDX download cancelled")
                        self.finished.emit(False, "Téléchargement annulé")
                        return
                    
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        chunk_count += 1
                        
                        if chunk_count % 100 == 0:
                            f.flush()
                            elapsed = time.time() - start_time
                            speed = downloaded / elapsed / (1024 * 1024) if elapsed > 0 else 0
                            
                            if total_size > 0:
                                progress = int((downloaded / total_size) * 1000)
                                remaining = total_size - downloaded
                                eta_seconds = remaining / (speed * 1024 * 1024) if speed > 0 else 0
                                eta_str = f"{int(eta_seconds // 60):02d}:{int(eta_seconds % 60):02d}"
                            else:
                                progress = 0
                                eta_str = f"{downloaded / (1024**3):.2f} GB"
                            
                            self.progress.emit(progress, speed, eta_str)
            
            # Download complete
            if total_size == 0 or downloaded >= total_size:
                self.progress.emit(1000, 0, "00:00")
            
            if os.path.exists(file_path):
                logger.info(f"VHDX download completed: {file_path}")
                self.finished.emit(True, f"Téléchargement terminé: {filename}")
            else:
                logger.error("Download failed: File not found after download")
                self.finished.emit(False, "Échec du téléchargement")
                
        except requests.exceptions.Timeout:
            logger.error("VHDX download timeout")
            self.finished.emit(False, "Timeout de téléchargement")
        except requests.exceptions.ConnectionError:
            logger.error("VHDX download connection error")
            self.finished.emit(False, "Erreur de connexion")
        except Exception as e:
            logger.error(f"VHDX download exception: {e}", exc_info=True)
            self.finished.emit(False, f"Erreur: {str(e)}")


class TorrentDownloadThread(QtCore.QThread):
    """Thread for downloading VHDX files via torrent"""
    finished = QtCore.Signal(bool, str)  # success, message
    progress = QtCore.Signal(int, float, str)  # progress (0-1000), speed (MB/s), ETA
    
    def __init__(self, torrent_url, destination_folder, destination_drive, target_file="PSA-DIAG.vhdx"):
        super().__init__()
        self.torrent_url = torrent_url
        self.destination_folder = destination_folder
        self.destination_drive = destination_drive
        self.target_file = target_file
        self._is_cancelled = False
        self._is_paused = False
        self.process = None
        
    def cancel(self):
        """Cancel the download"""
        self._is_cancelled = True
        if self.process:
            try:
                # Close stdout first to release the handle
                if self.process.stdout:
                    try:
                        self.process.stdout.close()
                    except:
                        pass
                
                # First try to terminate gracefully
                self.process.terminate()
                logger.info(f"Sent terminate signal to aria2c PID {self.process.pid}")
                
                # Wait up to 3 seconds for graceful termination
                try:
                    self.process.wait(timeout=3)
                    logger.info("aria2c terminated gracefully")
                except subprocess.TimeoutExpired:
                    # Force kill if still running
                    logger.warning("aria2c didn't terminate, force killing...")
                    self.process.kill()
                    self.process.wait(timeout=2)
                    logger.info("aria2c force killed")
                    
            except Exception as e:
                logger.error(f"Failed to terminate aria2c: {e}")
                # Last resort: use taskkill with process tree
                if sys.platform == 'win32' and hasattr(self, 'process_pid'):
                    try:
                        subprocess.run(['taskkill', '/F', '/T', '/PID', str(self.process_pid)], 
                                     capture_output=True, timeout=5)
                        logger.info(f"Used taskkill /T on PID {self.process_pid}")
                    except Exception as e2:
                        logger.error(f"taskkill also failed: {e2}")
        
    def pause(self):
        """Pause the download (not supported with aria2c)"""
        self._is_paused = True
        
    def resume(self):
        """Resume the download"""
        self._is_paused = False
    
    def run(self):
        import tempfile
        
        try:
            # Create VHD folder at root of selected drive
            vhd_folder = os.path.join(f"{self.destination_drive}:\\", "VHD")
            os.makedirs(vhd_folder, exist_ok=True)
            logger.info(f"VHD folder created/verified: {vhd_folder}")
            
            logger.info(f"Starting torrent download from: {self.torrent_url}")
            logger.info(f"Target file: {self.target_file}")
            logger.info(f"Destination: {vhd_folder}")
            
            # Download torrent file
            logger.info("Downloading .torrent file...")
            torrent_response = requests.get(self.torrent_url, timeout=30)
            torrent_response.raise_for_status()
            
            # Create temp file for torrent
            with tempfile.NamedTemporaryFile(mode='wb', suffix='.torrent', delete=False) as temp_torrent:
                temp_torrent.write(torrent_response.content)
                torrent_file_path = temp_torrent.name
            
            logger.info(f"Torrent file saved to: {torrent_file_path}")
            
            # Check for aria2c executable (bundled or system)
            aria2c_paths = [
                BASE / "tools" / "aria2c.exe",  # Bundled
                "aria2c"  # System PATH
            ]
            
            aria2c_exe = None
            for path in aria2c_paths:
                if isinstance(path, Path) and path.exists():
                    aria2c_exe = str(path)
                    break
                elif isinstance(path, str) and shutil.which(path):
                    aria2c_exe = path
                    break
            
            if not aria2c_exe:
                error_msg = "aria2c non trouvé. Téléchargez-le depuis https://github.com/aria2/aria2/releases"
                logger.error(error_msg)
                self.finished.emit(False, error_msg)
                os.unlink(torrent_file_path)
                return
            
            logger.info(f"Using aria2c: {aria2c_exe}")
            
            # First, list files in torrent to find target file index
            logger.info("Listing files in torrent...")
            list_cmd = [aria2c_exe, "--show-files", torrent_file_path]
            try:
                result = subprocess.run(
                    list_cmd,
                    capture_output=True,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
                )
                
                # Parse output to find file index
                # Format: idx|path/to/file.ext|size
                target_index = None
                for line in result.stdout.split('\n'):
                    if '|' in line:
                        parts = line.strip().split('|')
                        if len(parts) >= 2:
                            idx = parts[0].strip()
                            filepath = parts[1].strip()
                            # Check if filename matches (basename comparison)
                            if filepath.endswith(self.target_file) or os.path.basename(filepath) == self.target_file:
                                target_index = idx
                                logger.info(f"Found target file at index: {idx} - {filepath}")
                                break
                
                if not target_index:
                    error_msg = f"File '{self.target_file}' not found in torrent"
                    logger.error(error_msg)
                    logger.error(f"Available files:\n{result.stdout}")
                    self.finished.emit(False, error_msg)
                    os.unlink(torrent_file_path)
                    return
            except Exception as e:
                logger.error(f"Failed to list torrent files: {e}")
                self.finished.emit(False, f"Erreur lors de l'analyse du torrent: {str(e)}")
                os.unlink(torrent_file_path)
                return
            
            # Prepare aria2c command with file selection
            cmd = [
                aria2c_exe,
                torrent_file_path,
                f"--dir={vhd_folder}",
                f"--select-file={target_index}",  # Only download target file
                f"--index-out={target_index}={self.target_file}",  # Output directly without subdirs
                "--file-allocation=none",  # Disable pre-allocation for faster start
                "--seed-time=0",  # Don't seed after download
                "--bt-max-peers=50",
                "--max-connection-per-server=16",
                "--min-split-size=1M",
                "--split=16",
                "--enable-rpc=false",
                "--summary-interval=0",  # Disable summary
                "--console-log-level=notice",
                "--show-console-readout=true"  # Show progress on console
            ]
            
            logger.info(f"Starting aria2c download (file index: {target_index})...")
            
            # Emit initial progress
            self.progress.emit(0, 0.0, "--:--")
            
            # Use CREATE_NEW_PROCESS_GROUP to allow proper termination
            creation_flags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
            if sys.platform == 'win32':
                creation_flags |= subprocess.CREATE_NEW_PROCESS_GROUP
            
            self.process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True,
                creationflags=creation_flags
            )
            
            # Store PID for forced killing if needed
            self.process_pid = self.process.pid
            logger.info(f"aria2c started with PID: {self.process_pid}")
            
            # Read line by line instead of character by character
            while True:
                if self._is_cancelled:
                    logger.info("Cancellation requested, terminating aria2c...")
                    try:
                        # Close stdout to release handle
                        if self.process.stdout:
                            self.process.stdout.close()
                        
                        self.process.terminate()
                        self.process.wait(timeout=3)
                        logger.info("aria2c terminated on cancel")
                    except subprocess.TimeoutExpired:
                        logger.warning("Force killing aria2c on cancel...")
                        self.process.kill()
                        try:
                            self.process.wait(timeout=2)
                        except:
                            pass
                    except Exception as e:
                        logger.error(f"Error during cancel: {e}")
                    
                    logger.warning("Torrent download cancelled")
                    try:
                        os.unlink(torrent_file_path)
                    except:
                        pass
                    self.finished.emit(False, "Téléchargement annulé")
                    return
                
                # Check if process finished
                ret = self.process.poll()
                if ret is not None:
                    if ret == 0:
                        break
                    else:
                        error_msg = f"aria2c a échoué (code {ret})"
                        logger.error(error_msg)
                        # Read remaining output for error messages
                        remaining = self.process.stdout.read()
                        if remaining:
                            logger.error(f"aria2c output: {remaining}")
                        os.unlink(torrent_file_path)
                        self.finished.emit(False, error_msg)
                        return
                
                # Read line (handles both \n and \r)
                line = self.process.stdout.readline()
                if not line:
                    time.sleep(0.1)
                    continue
                
                line = line.strip()
                if not line:
                    continue
                
                # Skip DHT, BitTorrent notices, and redirection logs
                if any(skip_keyword in line for skip_keyword in [
                    "[ERROR]", "[NOTICE]", "DHT:", "BitTorrent:", 
                    "Redirecting to", "DHTRoutingTable", "listening on"
                ]):
                    continue
                
                # Skip file allocation lines - we only care about actual download progress
                if "FileAlloc:" in line:
                    continue
                
                # Parse progress from aria2c output
                # Formats: [#abc123 100MiB/1GiB(0%) CN:5 DL:1.2MiB ETA:5m30s]
                try:
                    if "DL:" in line and line.startswith("[#"):
                        # Find the first bracket block (main download progress)
                        first_bracket_end = line.find("]")
                        if first_bracket_end > 0:
                            main_line = line[:first_bracket_end]
                            
                            # Extract downloaded/total bytes (e.g., "80KiB/30GiB" or "1.5GiB/30GiB")
                            # Pattern: [#hash downloaded/total(pct%) ...]
                            progress_pct = 0.0
                            progress = 0
                            
                            # First try to extract percentage from parentheses as fallback
                            paren_start = main_line.find("(")
                            if paren_start > 0:
                                paren_end = main_line.find("%", paren_start)
                                if paren_end > paren_start:
                                    pct_str = main_line[paren_start+1:paren_end].strip()
                                    try:
                                        progress_pct = float(pct_str)
                                        progress = int(progress_pct * 10)
                                        logger.debug(f"Using aria2c reported percentage: {progress_pct:.2f}%")
                                    except ValueError:
                                        pass
                            
                            # Try to calculate from bytes if available
                            try:
                                hash_end = main_line.find(" ", 2)  # Skip "[#"
                                if hash_end > 0:
                                    rest = main_line[hash_end:].strip()
                                    # Find the part before "("
                                    paren_pos = rest.find("(")
                                    if paren_pos > 0:
                                        size_part = rest[:paren_pos].strip()  # e.g., "80KiB/30GiB"
                                        if "/" in size_part:
                                            downloaded_str, total_str = size_part.split("/", 1)
                                            
                                            # Parse size strings to bytes
                                            def parse_size(s):
                                                s = s.strip()
                                                multipliers = [
                                                    ('TiB', 1024**4),
                                                    ('GiB', 1024**3),
                                                    ('MiB', 1024**2),
                                                    ('KiB', 1024),
                                                    ('B', 1),
                                                ]
                                                for suffix, mult in multipliers:
                                                    if s.endswith(suffix):
                                                        num_str = s[:-len(suffix)].strip()
                                                        try:
                                                            return float(num_str) * mult
                                                        except ValueError:
                                                            return 0.0
                                                return 0.0
                                            
                                            downloaded_bytes = parse_size(downloaded_str)
                                            total_bytes = parse_size(total_str)
                                            
                                            if total_bytes > 0 and downloaded_bytes > 0:
                                                calc_progress_pct = (downloaded_bytes / total_bytes) * 100
                                                # Use calculated value if it's reasonable
                                                if calc_progress_pct >= progress_pct:
                                                    progress_pct = calc_progress_pct
                                                    progress = int(progress_pct * 10)
                                                    logger.debug(f"Calculated progress from bytes: {downloaded_str}/{total_str} = {progress_pct:.2f}%")
                            except Exception as e:
                                logger.debug(f"Failed to parse bytes progress: {e}")
                            
                            # Extract speed (DL:xxxKiB or DL:xxxMiB)
                            speed_mb = 0.0
                            if "DL:" in main_line:
                                dl_start = main_line.find("DL:") + 3
                                # Find next space or end
                                dl_end = len(main_line)
                                for i in range(dl_start, len(main_line)):
                                    if main_line[i] in [' ', ']', '\t']:
                                        dl_end = i
                                        break
                                speed_str = main_line[dl_start:dl_end].strip()
                                
                                logger.debug(f"Speed string extracted: '{speed_str}'")
                                
                                try:
                                    if "MiB" in speed_str:
                                        speed_mb = float(speed_str.replace("MiB", "").strip())
                                    elif "KiB" in speed_str:
                                        speed_mb = float(speed_str.replace("KiB", "").strip()) / 1024
                                    elif "GiB" in speed_str:
                                        speed_mb = float(speed_str.replace("GiB", "").strip()) * 1024
                                    elif "B" in speed_str and "i" not in speed_str:
                                        val = speed_str.replace("B", "").strip()
                                        if val and val != "0":
                                            speed_mb = float(val) / (1024 * 1024)
                                    
                                    logger.debug(f"Parsed speed: {speed_mb:.2f} MB/s")
                                except ValueError as ve:
                                    logger.warning(f"Cannot parse speed '{speed_str}': {ve}")
                            
                            # Extract ETA
                            eta_str = "--:--"
                            if "ETA:" in main_line:
                                eta_start = main_line.find("ETA:") + 4
                                eta_end = main_line.find("]", eta_start)
                                if eta_end == -1:
                                    eta_end = len(main_line)
                                eta_str = main_line[eta_start:eta_end].strip()
                            
                            # Always emit progress update
                            current_time = time.time()
                            if not hasattr(self, '_last_emit_time'):
                                self._last_emit_time = 0
                            
                            if current_time - self._last_emit_time >= 0.5:
                                logger.debug(f"Emitting progress: {progress_pct:.2f}% | {speed_mb:.2f} MB/s | {eta_str}")
                                self.progress.emit(progress, speed_mb, eta_str)
                                self._last_emit_time = current_time
                                
                except Exception as e:
                    logger.error(f"Failed to parse progress from '{line}': {e}", exc_info=True)
            
            # Download complete
            logger.info("Torrent download completed")
            self.progress.emit(1000, 0, "00:00")
            
            # Clean up torrent file
            os.unlink(torrent_file_path)
            
            # Verify file exists
            downloaded_file = os.path.join(vhd_folder, self.target_file)
            if os.path.exists(downloaded_file):
                logger.info(f"VHDX torrent download completed: {downloaded_file}")
                self.finished.emit(True, f"Téléchargement terminé: {self.target_file}")
            else:
                logger.error("Download failed: File not found after download")
                self.finished.emit(False, "Échec du téléchargement")
                
        except Exception as e:
            logger.error(f"Torrent download exception: {e}", exc_info=True)
            self.finished.emit(False, f"Erreur: {str(e)}")


class VHDXInstallThread(QtCore.QThread):
    """Thread for installing VHDX with BCD configuration"""
    finished = QtCore.Signal(bool, str)  # success, message
    
    def __init__(self, vhdx_path, description="PSA-DIAG"):
        super().__init__()
        self.vhdx_path = vhdx_path
        self.description = description
    
    def run(self):
        try:
            logger.info(f"Starting VHDX installation: {self.vhdx_path}")
            logger.info(f"Boot description: {self.description}")
            
            # Verify VHDX file exists
            if not os.path.exists(self.vhdx_path):
                error_msg = f"VHDX file not found: {self.vhdx_path}"
                logger.error(error_msg)
                self.finished.emit(False, error_msg)
                return
            
            # Extract relative path for BCD format
            # BCD expects: vhd=[locate]path format, e.g., vhd=[locate]\VHD\PSA-DIAG.vhdx
            vhd_path_relative = os.path.splitdrive(self.vhdx_path)[1]  # e.g., "\VHD\PSA-DIAG.vhdx"
            vhd_bcd_format = f"[locate]{vhd_path_relative}"
            
            logger.info(f"BCD VHD format: vhd={vhd_bcd_format}")
            
            # PowerShell script to configure BCD
            ps_script = f"""
$VHD = '{vhd_bcd_format}'
$description = '{self.description}'

try {{
    # Check if an entry with this description already exists
    $existingEntries = bcdedit /enum | Select-String -Pattern "description\\s+$description"
    
    if ($existingEntries) {{
        Write-Host "Boot entry '$description' already exists."
        
        # Extract CLSID from the existing entry
        $beforeDescription = bcdedit /enum | Select-String -Pattern "identificateur" -Context 0,10 | Where-Object {{
            $_.Context.PostContext -match "description\\s+$description"
        }}
        
        if ($beforeDescription -match '{{([a-fA-F0-9-]+)}}') {{
            $existingCLSID = "{{$($matches[1])}}"
            Write-Host "Existing CLSID found: $existingCLSID"
            Write-Host "Deleting old entry..."
            bcdedit /delete $existingCLSID /f
            if ($LASTEXITCODE -ne 0) {{
                throw "Failed to delete existing entry (exit code: $LASTEXITCODE)"
            }}
            Write-Host "Old entry deleted successfully"
        }}
    }}
    
    # Copy current boot entry
    $BootEntryCopy = bcdedit /copy '{{current}}' /d $description
    if ($LASTEXITCODE -ne 0) {{
        throw "Failed to copy boot entry (exit code: $LASTEXITCODE)"
    }}
    
    # Extract CLSID from output using regex
    if ($BootEntryCopy -match '{{([a-fA-F0-9-]+)}}') {{
        $CLSID = "{{$($matches[1])}}"
    }} else {{
        throw "Failed to extract CLSID from bcdedit output"
    }}
    
    Write-Host "Boot entry created: $CLSID"
    
    # Set device to VHD
    $result1 = bcdedit /set $CLSID device vhd=$VHD
    if ($LASTEXITCODE -ne 0) {{
        throw "Failed to set device (exit code: $LASTEXITCODE)"
    }}
    
    # Set osdevice to VHD
    $result2 = bcdedit /set $CLSID osdevice vhd=$VHD
    if ($LASTEXITCODE -ne 0) {{
        throw "Failed to set osdevice (exit code: $LASTEXITCODE)"
    }}
    
    # Set locale to fr-FR
    $result4 = bcdedit /set $CLSID locale fr-FR
    if ($LASTEXITCODE -ne 0) {{
        Write-Host "Warning: Failed to set locale (exit code: $LASTEXITCODE)"
    }}
    
    # Enable HAL detection
    $result5 = bcdedit /set $CLSID detecthal on
    if ($LASTEXITCODE -ne 0) {{
        throw "Failed to enable detecthal (exit code: $LASTEXITCODE)"
    }}
    
    Write-Host "BCD configuration completed successfully"
    Write-Host "CLSID: $CLSID"
    exit 0
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}
"""
            
            logger.info("Executing PowerShell BCD configuration script")
            
            # Execute PowerShell script
            proc = subprocess.run(
                ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
            )
            
            stdout = (proc.stdout or '').strip()
            stderr = (proc.stderr or '').strip()
            
            logger.info(f"PowerShell exit code: {proc.returncode}")
            if stdout:
                logger.info(f"PowerShell stdout: {stdout}")
            if stderr:
                logger.warning(f"PowerShell stderr: {stderr}")
            
            if proc.returncode == 0:
                success_msg = (
                    f"VHDX installation successful!\n\n"
                    f"The system is configured to boot from:\n{self.vhdx_path}\n\n"
                    f"Restart your computer and select '{self.description}' from the boot menu."
                )
                logger.info("VHDX installation completed successfully")
                self.finished.emit(True, success_msg)
            else:
                error_msg = f"BCD configuration failed.\n\n{stderr if stderr else 'Unknown error'}"
                logger.error(f"VHDX installation failed: {error_msg}")
                self.finished.emit(False, error_msg)
                
        except Exception as e:
            error_msg = f"Installation error: {str(e)}"
            logger.error(f"VHDX installation exception: {e}", exc_info=True)
            self.finished.emit(False, error_msg)


class BCDCleanupThread(QtCore.QThread):
    """Thread for removing PSA-DIAG BCD entries with backup"""
    finished = QtCore.Signal(bool, str)  # success, message
    
    def __init__(self):
        super().__init__()
    
    def run(self):
        try:
            logger.info("Starting BCD cleanup for PSA-DIAG entries")
            
            # PowerShell script to backup BCD and remove PSA-DIAG entries
            ps_script = r"""
try {
    # Create backup directory in temp
    $backupDir = Join-Path $env:TEMP "BCD_Backups"
    if (-not (Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    }
    
    # Generate backup filename with timestamp
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupFile = Join-Path $backupDir "BCD_Backup_$timestamp"
    
    # Backup current BCD
    Write-Host "Creating BCD backup: $backupFile"
    bcdedit /export $backupFile
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to backup BCD (exit code: $LASTEXITCODE)"
    }
    Write-Host "BCD backup created successfully"
    
    # Find all PSA-DIAG entries
    $bcdOutput = bcdedit /enum | Out-String
    $entries = @()
    
    # Parse bcdedit output to find PSA-DIAG entries
    $lines = $bcdOutput -split "`r?`n"
    $currentId = $null
    $currentDescription = $null
    
    foreach ($line in $lines) {
        if ($line -match 'identificateur\s+(\{[a-fA-F0-9-]+\})') {
            $currentId = $matches[1]
            $currentDescription = $null
        }
        elseif ($line -match 'description\s+(.+)') {
            $currentDescription = $matches[1].Trim()
            if ($currentDescription -eq "PSA-DIAG" -and $currentId) {
                $entries += $currentId
                Write-Host "Found PSA-DIAG entry: $currentId"
            }
        }
    }
    
    if ($entries.Count -eq 0) {
        Write-Host "No PSA-DIAG entries found in BCD"
        Write-Output "NO_ENTRIES|$backupFile"
        exit 0
    }
    
    # Delete each PSA-DIAG entry
    $deletedCount = 0
    foreach ($entry in $entries) {
        Write-Host "Deleting entry: $entry"
        bcdedit /delete $entry /f
        if ($LASTEXITCODE -eq 0) {
            $deletedCount++
            Write-Host "Successfully deleted: $entry"
        }
        else {
            Write-Host "Warning: Failed to delete $entry (exit code: $LASTEXITCODE)"
        }
    }
    
    Write-Output "SUCCESS|$deletedCount|$backupFile"
    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
"""
            
            logger.info("Executing PowerShell BCD cleanup script")
            
            # Execute PowerShell script
            proc = subprocess.run(
                ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
            )
            
            stdout = (proc.stdout or '').strip()
            stderr = (proc.stderr or '').strip()
            
            logger.info(f"PowerShell exit code: {proc.returncode}")
            if stdout:
                logger.info(f"PowerShell stdout: {stdout}")
            if stderr:
                logger.warning(f"PowerShell stderr: {stderr}")
            
            if proc.returncode == 0:
                # Parse output
                if stdout.startswith("NO_ENTRIES|"):
                    parts = stdout.split("|")
                    backup_file = parts[1] if len(parts) > 1 else "unknown"
                    msg = translator.t('messages.bcd_cleanup.no_entries', backup=backup_file)
                    logger.info("No PSA-DIAG entries found")
                    self.finished.emit(True, msg)
                elif stdout.startswith("SUCCESS|"):
                    parts = stdout.split("|")
                    count = parts[1] if len(parts) > 1 else "0"
                    backup_file = parts[2] if len(parts) > 2 else "unknown"
                    msg = translator.t('messages.bcd_cleanup.success', count=count, backup=backup_file)
                    logger.info(f"Successfully removed {count} PSA-DIAG entries")
                    self.finished.emit(True, msg)
                else:
                    msg = translator.t('messages.bcd_cleanup.success_generic')
                    self.finished.emit(True, msg)
            else:
                error_msg = f"BCD cleanup failed.\n\n{stderr if stderr else 'Unknown error'}"
                logger.error(f"BCD cleanup failed: {error_msg}")
                self.finished.emit(False, error_msg)
                
        except Exception as e:
            error_msg = f"BCD cleanup error: {str(e)}"
            logger.error(f"BCD cleanup exception: {e}", exc_info=True)
            self.finished.emit(False, error_msg)


class InstallThread(QtCore.QThread):
    finished = QtCore.Signal(bool, str)
    progress = QtCore.Signal(int)  # progress percentage
    file_progress = QtCore.Signal(str)  # current file being extracted
    # Signals to inform UI about runtimes installer state
    runtimes_started = QtCore.Signal()
    runtimes_finished = QtCore.Signal(bool, str)  # success, message
    # Signal to report driver installation result when run inside the install thread
    driver_finished = QtCore.Signal(bool, str)  # success, message
    # Signal to report defender rule creation result
    defender_finished = QtCore.Signal(bool, str)

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
            
            # FIRST: Create Windows Defender exclusions BEFORE extraction (may require admin)
            try:
                defender_paths = [
                    r"C:\AWRoot",
                    r"C:\INSTALL",
                    r"C:\Program Files (x86)\PSA VCI",
                    r"C:\Program Files\PSA VCI",
                    r"C:\Windows\VCX.dll",
                ]
                # Build PowerShell script to add any missing exclusions and return JSON
                ps_paths = ",".join(["'{}'".format(p) for p in defender_paths])
                ps_script = (
                    "try { $existing=(Get-MpPreference).ExclusionPath; $added=@(); $failed=@();"
                    + "foreach($p in @(" + ps_paths + ")) { if($existing -notcontains $p) { try { Add-MpPreference -ExclusionPath $p; $added += $p } catch { $failed += $p } } }"
                    + "$res = @{added=$added; failed=$failed}; $res | ConvertTo-Json -Compress } catch { Write-Error $_; exit 1 }"
                )
                cmd = ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script]
                logger.info("Attempting to create Defender exclusions via PowerShell (before extraction)")
                proc = subprocess.run(cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
                if proc.returncode == 0:
                    out = (proc.stdout or '').strip()
                    try:
                        import json
                        j = json.loads(out) if out else {"added":[], "failed":[]}
                        added = j.get('added') or []
                        failed = j.get('failed') or []
                        if failed:
                            msg = f"Failed to add exclusions: {', '.join(failed)}"
                            logger.warning(msg)
                            try:
                                self.defender_finished.emit(False, msg)
                            except Exception:
                                pass
                        else:
                            msg = f"Added exclusions: {', '.join(added)}" if added else "No changes needed"
                            logger.info(msg)
                            try:
                                self.defender_finished.emit(True, msg)
                            except Exception:
                                pass
                    except Exception as e:
                        logger.error(f"Failed to parse Defender PowerShell output: {e}")
                        try:
                            self.defender_finished.emit(False, f"Parse error: {e}")
                        except Exception:
                            pass
                else:
                    err = (proc.stderr or '').strip()
                    logger.error(f"PowerShell exited with code {proc.returncode}: {err}")
                    try:
                        self.defender_finished.emit(False, f"PowerShell error: {err}")
                    except Exception:
                        pass
            except Exception as e:
                logger.error(f"Failed to create Defender exclusions: {e}", exc_info=True)
                try:
                    self.defender_finished.emit(False, str(e))
                except Exception:
                    pass
            
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

            # Post-extraction verification: ensure expected install artifacts exist
            verification_paths = [
                r"C:\AWRoot\bin\launcher\Diagbox.exe",
                r"C:\AWRoot\bin\fi\Version.ini",
            ]
            verification_ok = False
            if extraction_succeeded:
                # Give the extraction a moment to flush files (short poll)
                for _ in range(6):  # up to ~3 seconds
                    for p in verification_paths:
                        try:
                            if os.path.exists(p):
                                verification_ok = True
                                break
                        except Exception:
                            continue
                    if verification_ok:
                        break
                    time.sleep(0.5)

                if not verification_ok:
                    logger.warning("Post-extraction verification failed: expected files not found")
                    extraction_succeeded = False
                    extraction_errors.append("Extraction incomplete or interrupted: expected files missing")

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
            # If there were extraction errors, treat the whole installation as failed.
            if extraction_errors:
                error_summary = "\n".join(extraction_errors + runtimes_warnings)
                logger.error(f"Installation failed due to extraction errors: {error_summary}")
                message = translator.t('messages.install.failed', error=error_summary)
                self.finished.emit(False, message)
            else:
                # No extraction errors; treat runtimes warnings as non-fatal warnings
                if runtimes_warnings:
                    warning_summary = "\n".join(runtimes_warnings)
                    logger.warning(f"Installation completed with warnings: {warning_summary}")
                    message = translator.t('messages.install.warnings', warnings=warning_summary)
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
    manual_defender_finished = QtCore.Signal(bool, str)  # success, message - for manual defender button

    def __init__(self, splash=None):
        super().__init__()
        self.splash = splash  # Keep reference to splash screen
        
        self.setWindowTitle(translator.t('app.title'))
        self.setWindowFlags(QtCore.Qt.WindowType.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground)
        self.resize(1220, 500)  # Increased width to accommodate log sidebar

        # Download variables
        self.download_folder = "C:\\INSTALL\\"
        self.last_version_diagbox = ""
        self.diagbox_path = ""
        self.auto_install = None
        self.download_thread = None
        self.install_thread = None
        self.cancel_button = None
        self.pause_button = None
        self.vhd_pause_button = None
        self.vhd_cancel_button = None
        self.dragPos = QtCore.QPoint()
        self.log_widget = None
        self.log_handler = None
        
        # Version options: load from remote JSON (configured in `config.URL_VERSION_OPTIONS`)
        # Falls back to the built-in defaults if remote fetch fails.
        logger.info("[STEP 1] -- Loading version options...")
        self.version_options = self.load_version_options()
        
        # Fetch last version after version_options is loaded
        logger.info("[STEP 2] -- Fetching last Diagbox version...")
        self.fetch_last_version_diagbox()
        
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
        self.manual_defender_finished.connect(self._on_manual_defender_finished)

        self.setup_ui()
        
        # Keep splash screen visible until all loading is complete
        # It will be closed by _close_splash_screen() after load_changelog finishes
        
        # Check for app updates after UI is ready
        QtCore.QTimer.singleShot(1000, self.check_app_update)
    
    def _close_splash_screen(self):
        """Close the splash screen once all loading is complete"""
        if self.splash:
            logger.debug("Closing splash screen - all data loaded")
            self.splash.close()
            self.splash = None

    def load_version_options(self):
    

        try:
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
                logger.info(f"Loaded {len(options)} version options")
                return options
            else:
                logger.warning("Version options JSON did not contain valid entries")
                # Return empty list to indicate maintenance / no available versions
                return []
        except requests.exceptions.RequestException:
            logger.warning("Unable to reach update server. Please check your network connection.")
            return []
        except Exception as e:
            logger.warning(f"Failed to load version options: {e}")
            return []

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
        except requests.exceptions.RequestException:
            logger.debug("Unable to load remote messages (no network connection)")
        except Exception as e:
            logger.debug(f"Failed to load remote messages: {e}")

    def update_global_banner(self):
        """Create or update a single global banner from `self.remote_messages`.

        This banner is shared across multiple pages and only shows messages
        relevant to the current page based on `display_on` property.
        """
        try:
            # Ensure global banner exists
            if not hasattr(self, 'global_banner'):
                self._create_global_banner()
            
            # Update banner content based on current page
            self._update_banner_for_current_page()
            
        except Exception as e:
            logger.debug(f"update_global_banner error: {e}")
    
    def _create_global_banner(self):
        """Create a single reusable banner widget"""
        self.global_banner = QtWidgets.QFrame()
        self.global_banner.setObjectName('installBanner')
        self.global_banner.setMaximumHeight(150)  # Limit banner height
        main_banner_layout = QtWidgets.QVBoxLayout(self.global_banner)
        main_banner_layout.setContentsMargins(10,10,10,10)
        main_banner_layout.setSpacing(8)
        
        bann_layout = QtWidgets.QHBoxLayout()
        bann_layout.setSpacing(8)
        bann_layout.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        
        # Left arrow (no vertical stretch)
        left_arrow = QtWidgets.QPushButton('\u276E')
        left_arrow.setObjectName('bannerArrow')
        left_arrow.setFixedSize(32, 32)
        left_arrow.clicked.connect(self._prev_banner)
        bann_layout.addWidget(left_arrow, 0, QtCore.Qt.AlignmentFlag.AlignVCenter)
        self.banner_left_arrow = left_arrow
        
        # Center content
        center_layout = QtWidgets.QVBoxLayout()
        center_layout.setSpacing(8)
        
        lbl = QtWidgets.QLabel("")
        lbl.setObjectName('installBannerLabel')
        lbl.setWordWrap(True)
        lbl.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        center_layout.addWidget(lbl)
        
        # Link button row
        link_layout = QtWidgets.QHBoxLayout()
        link_layout.addStretch()
        self.banner_link_layout = link_layout
        link_layout.addStretch()
        center_layout.addLayout(link_layout)
        
        # Pagination dots
        dots_layout = QtWidgets.QHBoxLayout()
        dots_layout.setSpacing(6)
        dots_layout.addStretch()
        self.banner_dots_layout = dots_layout
        self.banner_dots = []
        dots_layout.addStretch()
        center_layout.addLayout(dots_layout)
        
        bann_layout.addLayout(center_layout, 1)
        self.banner_link = None
        
        # Right arrow (no vertical stretch)
        right_arrow = QtWidgets.QPushButton('\u276F')
        right_arrow.setObjectName('bannerArrow')
        right_arrow.setFixedSize(32, 32)
        right_arrow.clicked.connect(self._next_banner)
        bann_layout.addWidget(right_arrow, 0, QtCore.Qt.AlignmentFlag.AlignVCenter)
        self.banner_right_arrow = right_arrow
        
        main_banner_layout.addLayout(bann_layout)
        
        self.banner_label = lbl
        self.banner_messages = []
        self.banner_index = 0
        self.banner_timer = QtCore.QTimer(self)
        self.banner_timer.timeout.connect(self._advance_banner)
    
    def _update_banner_for_current_page(self):
        """Update banner content based on current page"""
        try:
            if not getattr(self, 'remote_messages', None):
                if hasattr(self, 'global_banner'):
                    self.global_banner.setVisible(False)
                return
            
            # Determine current page
            current_index = self.stack.currentIndex()
            page_keys = {0: 'home', 1: 'download'}  # Map stack index to page key
            page_key = page_keys.get(current_index)
            
            if not page_key:
                if hasattr(self, 'global_banner'):
                    self.global_banner.setVisible(False)
                return
            
            # Collect messages for current page
            now = QtCore.QDateTime.currentDateTimeUtc()
            candidates = []
            for msg in self.remote_messages:
                raw = msg.get('raw', {})
                display_on = raw.get('display_on') or []
                if display_on and page_key not in display_on:
                    continue
                # Check time window
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
                self.global_banner.setVisible(False)
                return
            
            # Update banner content
            self.banner_messages = candidates
            self.banner_index = 0
            
            # Ensure banner is in current page layout
            current_widget = self.stack.currentWidget()
            if current_widget:
                layout = current_widget.layout()
                if layout:
                    # Remove banner from old parent if exists
                    if self.global_banner.parent() != current_widget:
                        old_parent = self.global_banner.parent()
                        if old_parent:
                            old_parent.layout().removeWidget(self.global_banner)
                        layout.insertWidget(0, self.global_banner)
            
            # Start/stop timer based on message count
            if len(self.banner_messages) > 1:
                if not self.banner_timer.isActive():
                    self.banner_timer.start(8000)
            else:
                self.banner_timer.stop()
            
            # Update display
            self._update_banner_dots()
            self._show_banner_message(self.banner_index)
            self.global_banner.setVisible(True)
            
        except Exception as e:
            logger.debug(f"_update_banner_for_current_page error: {e}")
    
    def _show_banner_message(self, index):
        """Display message at `index` from `self.banner_messages`."""
        try:
            if not getattr(self, 'banner_messages', None):
                return
            msg = self.banner_messages[index]
            raw = msg.get('raw', {})
            langmap = msg.get('lang', {})
            lang_code = translator.language if hasattr(translator, 'language') else 'en'
            text_entry = langmap.get(lang_code) or langmap.get('en') or {}
            text = text_entry.get('text') if isinstance(text_entry, dict) else text_entry or str(text_entry)
            link = text_entry.get('link') if isinstance(text_entry, dict) else None
            link_text = text_entry.get('link_text') if isinstance(text_entry, dict) else None

            self.banner_label.setText(text or '')
            if link:
                if not getattr(self, 'banner_link', None):
                    btn = QtWidgets.QPushButton(link_text or 'Link')
                    btn.setObjectName('bannerLink')
                    btn._url = link
                    btn.clicked.connect(self._open_button_url)
                    self.banner_link_layout.insertWidget(1, btn)
                    self.banner_link = btn
                else:
                    self.banner_link.setText(link_text or 'Link')
                    self.banner_link._url = link
                    self.banner_link.setVisible(True)
            else:
                if getattr(self, 'banner_link', None):
                    self.banner_link.setVisible(False)
        except Exception as e:
            logger.debug(f"_show_banner_message error: {e}")
    
    def _advance_banner(self):
        try:
            if not getattr(self, 'banner_messages', None):
                return
            self.banner_index = (self.banner_index + 1) % len(self.banner_messages)
            self._show_banner_message(self.banner_index)
            self._update_banner_dots()
        except Exception as e:
            logger.debug(f"_advance_banner error: {e}")
    
    def _next_banner(self):
        try:
            if not getattr(self, 'banner_messages', None):
                return
            if self.banner_timer.isActive():
                self.banner_timer.stop()
            self.banner_index = (self.banner_index + 1) % len(self.banner_messages)
            self._show_banner_message(self.banner_index)
            self._update_banner_dots()
            if len(self.banner_messages) > 1:
                self.banner_timer.start(8000)
        except Exception as e:
            logger.debug(f"_next_banner error: {e}")
    
    def _prev_banner(self):
        try:
            if not getattr(self, 'banner_messages', None):
                return
            if self.banner_timer.isActive():
                self.banner_timer.stop()
            self.banner_index = (self.banner_index - 1) % len(self.banner_messages)
            self._show_banner_message(self.banner_index)
            self._update_banner_dots()
            if len(self.banner_messages) > 1:
                self.banner_timer.start(8000)
        except Exception as e:
            logger.debug(f"_prev_banner error: {e}")
    
    def _update_banner_dots(self):
        """Update pagination dots to reflect current message index."""
        try:
            if not hasattr(self, 'banner_dots_layout'):
                return
            num_messages = len(getattr(self, 'banner_messages', []))
            if num_messages <= 1:
                for dot in self.banner_dots:
                    dot.setVisible(False)
                if hasattr(self, 'banner_left_arrow'):
                    self.banner_left_arrow.setVisible(False)
                if hasattr(self, 'banner_right_arrow'):
                    self.banner_right_arrow.setVisible(False)
                return
            
            if hasattr(self, 'banner_left_arrow'):
                self.banner_left_arrow.setVisible(True)
            if hasattr(self, 'banner_right_arrow'):
                self.banner_right_arrow.setVisible(True)
            
            current_index = getattr(self, 'banner_index', 0)
            
            while len(self.banner_dots) > num_messages:
                dot = self.banner_dots.pop()
                self.banner_dots_layout.removeWidget(dot)
                dot.deleteLater()
            
            while len(self.banner_dots) < num_messages:
                dot = QtWidgets.QLabel('\u25CF')
                dot.setObjectName('bannerDot')
                dot.setFixedSize(10, 10)
                dot.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                dot.setStyleSheet('font-size: 10px; color: #888;')
                self.banner_dots_layout.insertWidget(len(self.banner_dots) + 1, dot)
                self.banner_dots.append(dot)
            
            for i, dot in enumerate(self.banner_dots):
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
            logger.debug(f"Diagbox language file not found: {lang_file}")

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
        try:
            current_widget = self.stack.currentWidget()
            if current_widget:
                for child in current_widget.findChildren(QtWidgets.QPushButton):
                    if child != self.cancel_button and child != self.pause_button and child != self.vhd_cancel_button and child != self.vhd_pause_button:
                        child.setEnabled(enabled)
        except Exception as e:
            logger.warning(f"Failed to toggle buttons: {e}")

        # Disable/enable combo box
        if hasattr(self, 'version_combo') and self.version_combo:
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

    def cancel_vhd_download(self):
        """Cancel the current VHD download"""
        if hasattr(self, 'vhdx_download_thread') and self.vhdx_download_thread and self.vhdx_download_thread.isRunning():
            self.vhdx_download_thread.cancel()

    def toggle_pause_vhd_download(self):
        """Pause or resume the VHD download"""
        if hasattr(self, 'vhdx_download_thread') and self.vhdx_download_thread and self.vhdx_download_thread.isRunning():
            if self.vhdx_download_thread._is_paused:
                self.vhdx_download_thread.resume()
                if self.vhd_pause_button:
                    self.vhd_pause_button.setText(translator.t('buttons.pause'))
            else:
                self.vhdx_download_thread.pause()
                if self.vhd_pause_button:
                    self.vhd_pause_button.setText(translator.t('buttons.resume'))

    def toggle_pause_vhd_download(self):
        """Pause or resume the VHD download"""
        if hasattr(self, 'vhdx_download_thread') and self.vhdx_download_thread and self.vhdx_download_thread.isRunning():
            if self.vhdx_download_thread._is_paused:
                self.vhdx_download_thread.resume()
                if self.vhd_pause_button:
                    self.vhd_pause_button.setText(translator.t('buttons.pause'))
            else:
                self.vhdx_download_thread.pause()
                if self.vhd_pause_button:
                    self.vhd_pause_button.setText(translator.t('buttons.resume'))

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

            # Defender exclusions line
            defender_success = None
            defender_msg = None
            if hasattr(self, '_defender_install_result') and self._defender_install_result is not None:
                defender_success, defender_msg = self._defender_install_result

            def_label = translator.t('messages.install.summary.defender')
            if defender_success is None:
                def_status = translator.t('messages.install.summary.not_run')
            else:
                if defender_success:
                    def_status = translator.t('messages.install.summary.ok')
                else:
                    def_status = translator.t('messages.install.summary.error', msg=(defender_msg or ''))
            lines.append(f"-- {def_label} : {def_status}")

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
            
            # Update Diagbox language selector visibility
            self._update_diagbox_language_visibility(installed_version is not None)
        
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
            self.footer_progress.setFormat(f"Extraction... {value}%")
        if hasattr(self, 'footer_label'):
            self.footer_label.setText("InstallDiagbox...")
        
        QtWidgets.QApplication.processEvents()
    
    def update_install_file(self, filename):
        """Update current file being extracted"""
        # Update footer label with truncated filename
        if hasattr(self, 'footer_label'):
            display_name = filename if len(filename) <= 60 else "..." + filename[-57:]
            self.footer_label.setText(f"Install: {display_name}")
        
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

    def _on_manual_defender_finished(self, success: bool, message: str):
        """Called when manual defender button finishes - runs in main thread"""
        try:
            logger.info(f"[MANUAL DEFENDER] Finished: success={success}")
            if success:
                QtWidgets.QMessageBox.information(self, translator.t('messages.defender.title'), message)
            else:
                QtWidgets.QMessageBox.warning(self, translator.t('messages.defender.title'), message)
        except Exception as e:
            logger.error(f"_on_manual_defender_finished error: {e}", exc_info=True)

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

    def _on_defender_finished_from_installthread(self, success: bool, message: str):
        """Handler for Defender exclusions creation result emitted by InstallThread."""
        try:
            logger.info(f"[DEFENDER] Finished callback invoked: success={success}")
            try:
                self._defender_install_result = (success, message)
            except Exception:
                self._defender_install_result = (success, message)
        except Exception as e:
            logger.error(f"_on_defender_finished_from_installthread error: {e}", exc_info=True)

    def install_diagbox(self):
        logger.info("Install Diagbox initiated")
    
        
        # If a Diagbox version is already installed, require cleaning first
        installed_version = self.check_installed_version()
        if installed_version:
            
            logger.warning(f"Diagbox is already installed, need to clean before continue")
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
        # Get selected version from combo box, or use first available local file if in maintenance mode
        if hasattr(self, 'version_combo') and self.version_combo is not None:
            selected_data = self.version_combo.currentData()
            if selected_data:
                version, url = selected_data
                self.last_version_diagbox = version
                # Try three possible formats:
                # 1. Normalized format (what download_diagbox creates): 09.186.7z
                # 2. Full version format: 09.186_PSA_DIAG.7z
                # 3. Old format: Diagbox_Install_09.186_PSA_DIAG.7z
                normalized = self._sanitize_version_for_filename(version)
                normalized_format = os.path.join(self.download_folder, f"{normalized}.7z")
                full_format = os.path.join(self.download_folder, f"{version}.7z")
                old_format = os.path.join(self.download_folder, f"Diagbox_Install_{version}.7z")
                
                if os.path.exists(normalized_format):
                    self.diagbox_path = normalized_format
                elif os.path.exists(full_format):
                    self.diagbox_path = full_format
                else:
                    self.diagbox_path = old_format
                logger.info(f"Installing version: {version}, path: {self.diagbox_path}")
        else:
            # Maintenance mode: no combo box, use first available local file
            downloaded_versions = self.check_downloaded_versions()
            if downloaded_versions:
                first_file = downloaded_versions[0]
                self.last_version_diagbox = first_file['version']
                self.diagbox_path = first_file['path']
                logger.info(f"Installing local file (maintenance mode): {self.last_version_diagbox}, path: {self.diagbox_path}")
            else:
                QtWidgets.QMessageBox.warning(
                    self,
                    translator.t('messages.install.title'),
                    translator.t('messages.install.no_local_files')
                )
                return
        
        if not os.path.exists(self.diagbox_path):
            logger.error(f"Diagbox file not found: {self.diagbox_path}")
            # Get the version being attempted
            version = self.last_version_diagbox if self.last_version_diagbox else "Unknown"
            # Build all three possible paths for error message
            normalized = self._sanitize_version_for_filename(version)
            normalized_path = os.path.join(self.download_folder, f"{normalized}.7z")
            full_path = os.path.join(self.download_folder, f"{version}.7z")
            old_path = os.path.join(self.download_folder, f"Diagbox_Install_{version}.7z")
            all_paths = f"{normalized_path}\n  ou\n{full_path}\n  ou\n{old_path}"
            QtWidgets.QMessageBox.warning(
                self, 
                translator.t('messages.install.title'), 
                translator.t('messages.install.file_not_found', version=version, path=all_paths)
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
            # Connect defender finished signal so we can include it in the final summary
            try:
                self.install_thread.defender_finished.connect(self._on_defender_finished_from_installthread)
            except Exception:
                logger.debug("Failed to connect defender_finished signal")
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
        
        # Initialize footer progress to indeterminate mode
        if hasattr(self, 'footer_progress'):
            # Use indeterminate progress (0, 0) during cleaning
            self.footer_progress.setRange(0, 0)
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
        
        # Refresh install page and update all info (version info, language selector, etc.)
        self.refresh_install_page()
        
        # Also call on_enter_install_page to update version info banner and language selector
        try:
            QtCore.QTimer.singleShot(100, self.on_enter_install_page)
        except Exception:
            try:
                self.on_enter_install_page()
            except Exception:
                pass
        
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
        """Get the latest Diagbox version from version_options.
        
        Note: This uses get_latest_available_version() which reads from
        self.version_options (loaded from remote JSON). This ensures
        consistency with what is displayed in the UI.
        """
        try:
            latest = self.get_latest_available_version()
            if latest:
                self.last_version_diagbox = latest
                logger.info(f"Last Diagbox version: {self.last_version_diagbox}")
                # Try three possible file formats
                normalized = self._sanitize_version_for_filename(self.last_version_diagbox)
                normalized_format = os.path.join(self.download_folder, f"{normalized}.7z")
                full_format = os.path.join(self.download_folder, f"{self.last_version_diagbox}.7z")
                old_format = os.path.join(self.download_folder, f"Diagbox_Install_{self.last_version_diagbox}.7z")
                
                if os.path.exists(normalized_format):
                    self.diagbox_path = normalized_format
                elif os.path.exists(full_format):
                    self.diagbox_path = full_format
                else:
                    self.diagbox_path = old_format
            else:
                logger.warning("No version options available")
        except Exception as e:
            logger.error(f"Failed to fetch last version: {e}", exc_info=True)

    def check_app_update(self):
        """Check if a newer version of PSA-DIAG is available"""
        try:
            logger.info("[STEP 4] -- Checking for app updates...")
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
        except requests.exceptions.RequestException:
            logger.info("Unable to check for updates. Please verify your network connection.")
        except Exception as e:
            logger.warning(f"Update check failed: {e}")

    def perform_self_update(self, latest_version):
        """Download the latest release .exe from GitHub and invoke the updater helper.

        Note: This implementation uses the GitHub Releases API to locate a release asset
        that looks like an executable (.exe). It downloads to `CONFIG_DIR/updates/`.
        It then launches the `updater.py` helper (bundled in the project) which will
        wait for this process to exit, replace the running exe, and optionally restart it.
        """
        try:
            # Create a session with retry strategy for robust downloads
            from requests.adapters import HTTPAdapter
            from urllib3.util.retry import Retry
            
            # Use proper headers to avoid GitHub blocking
            headers = {
                'User-Agent': 'PSA-DIAG/2.3.1.0'
            }
            
            session = requests.Session()
            session.headers.update(headers)
            
            # More aggressive retry strategy for GitHub downloads
            retry_strategy = Retry(
                total=5,  # More retries for large files
                backoff_factor=2,  # Wait 2, 4, 8, 16 seconds between retries
                status_forcelist=[429, 500, 502, 503, 504],
                allowed_methods=["GET"],
                raise_on_status=False  # Don't raise on status code, we'll handle it
            )
            adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=1, pool_maxsize=1)
            session.mount("https://", adapter)
            session.mount("http://", adapter)
            
            api_url = "https://api.github.com/repos/RetroGameSets/PSA-DIAG/releases/latest"
            logger.debug(f"Querying GitHub API for latest release: {api_url}")
            r = session.get(api_url, timeout=15)
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

            # Download with progress and robust streaming
            logger.debug(f"Downloading update asset: {download_url} -> {download_path}")
            logger.info(f"Download size: {exe_asset.get('size', 'unknown')} bytes")
            
            # Use increased timeout and smaller chunks for stability
            downloaded_bytes = 0
            # Headers specific to the file download (not for API calls)
            download_headers = {
                'User-Agent': 'PSA-DIAG/2.3.1.0',
                'Accept': 'application/octet-stream'
            }
            
            try:
                with session.get(download_url, stream=True, timeout=120, headers=download_headers, verify=True) as resp:
                    resp.raise_for_status()
                    total_size = int(resp.headers.get('content-length', 0))
                    
                    with open(download_path, 'wb') as f:
                        # Download in smaller chunks with reasonable buffer size
                        for chunk in resp.iter_content(chunk_size=32768):  # 32KB chunks
                            if chunk:
                                f.write(chunk)
                                downloaded_bytes += len(chunk)
                                if total_size > 0:
                                    percent = (downloaded_bytes / total_size) * 100
                                    logger.debug(f"Download progress: {percent:.1f}% ({downloaded_bytes}/{total_size} bytes)")
            except requests.exceptions.ChunkedEncodingError as e:
                logger.error(f"Chunked encoding error during download: {e}. Downloaded {downloaded_bytes} bytes so far.")
                # Check if we have at least some data
                if os.path.exists(download_path) and os.path.getsize(download_path) > 1024 * 1024:  # At least 1MB
                    logger.warning("Partial download detected but size is reasonable, attempting to continue")
                else:
                    raise
            
            if os.path.exists(download_path):
                file_size = os.path.getsize(download_path)
                logger.info(f"Download complete: {download_path} ({file_size} bytes)")
            else:
                raise FileNotFoundError(f"Downloaded file not found: {download_path}")
                
            session.close()

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
        
        # Get selected version from combo box (and associated URL)
        url = None
        if hasattr(self, 'version_combo'):
            selected_data = self.version_combo.currentData()
            if selected_data:
                version, url = selected_data
                self.last_version_diagbox = version
                logger.info(f"Selected version: {version}")
        # If we still don't have an URL, try to resolve it from loaded version options
        if not url:
            # Ensure we have a last version value to look up
            if not self.last_version_diagbox:
                self.fetch_last_version_diagbox()
            if not self.last_version_diagbox:
                return

            # Try to find an entry in self.version_options matching the version
            for display_name, version, vurl in getattr(self, 'version_options', []):
                if version == self.last_version_diagbox or display_name == self.last_version_diagbox:
                    url = vurl
                    break

        # If no URL was found, warn the user and abort (no hard-coded archive.org fallback)
        if not url:
            QtWidgets.QMessageBox.warning(self, translator.t('messages.download.title'),
                                          translator.t('messages.download.no_url', version=self.last_version_diagbox))
            return

        # Set path to new format (use a normalized numeric version for filename)
        normalized = self._sanitize_version_for_filename(self.last_version_diagbox)
        self.diagbox_path = os.path.join(self.download_folder, f"{normalized}.7z")
        
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
        btn_vhd = SidebarButton("", icon_folder / "vhd.svg")
        btn_info = SidebarButton("", icon_folder / "info.svg")

        # Make first checked
        btn_diag.setChecked(True)

        vbox.addWidget(btn_diag)
        vbox.addWidget(btn_setup)
        vbox.addWidget(btn_vhd)
        vbox.addStretch()
        vbox.addWidget(btn_info)

        # Logo in sidebar bottom (clickable link to website)
        logo = QtWidgets.QLabel()
        logo.setOpenExternalLinks(True)
        logo.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        pix = QtGui.QPixmap(str(BASE / "icons" / "logo.png"))
        if not pix.isNull():
            pix = pix.scaledToWidth(80, QtCore.Qt.TransformationMode.SmoothTransformation)
            logo.setPixmap(pix)
            logo.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        # Make logo clickable
        logo.mousePressEvent = lambda event: QtGui.QDesktopServices.openUrl(QtCore.QUrl("https://www.psa-diag.fr/"))
        vbox.addWidget(logo)

        # Donate button under logo
        donate_btn = QtWidgets.QPushButton("♥ Donate")
        donate_btn.setFixedHeight(32)
        donate_btn.setStyleSheet("""
            QPushButton {
                background-color: #2a2a2a;
                color: #4fc3f7;
                border: 1px solid #4fc3f7;
                border-radius: 4px;
                font-size: 11px;
                font-weight: bold;
                padding: 4px;
            }
            QPushButton:hover {
                background-color: #4fc3f7;
                color: #1e1e1e;
            }
            QPushButton:pressed {
                background-color: #0288d1;
            }
        """)
        donate_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        donate_btn.clicked.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl("https://buymeacoffee.com/PsaDiagOfficial")))
        vbox.addWidget(donate_btn)

        # Display app version under logo in the sidebar
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
        self.stack.addWidget(self.page_vhd())
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

        # Right sidebar for logs
        log_sidebar = QtWidgets.QFrame()
        log_sidebar.setObjectName("sidebar")
        log_sidebar.setFixedWidth(280)
        log_sidebar_layout = QtWidgets.QVBoxLayout(log_sidebar)
        log_sidebar_layout.setContentsMargins(12, 12, 12, 12)
        log_sidebar_layout.setSpacing(10)
        
        # Log title
        log_title = QtWidgets.QLabel("LOGS")
        log_title.setObjectName("titleLabel")
        log_title.setStyleSheet("font-size: 13px; font-weight: bold;")
        log_title.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        log_sidebar_layout.addWidget(log_title)
        
        # Log text widget
        self.log_widget = QtWidgets.QTextEdit()
        self.log_widget.setObjectName("logWidget")
        self.log_widget.setReadOnly(True)
        log_sidebar_layout.addWidget(self.log_widget)
        
        # Open logs button at bottom of log sidebar
        self.open_log_btn = QtWidgets.QPushButton(translator.t('buttons.open_log'))
        self.open_log_btn.setFixedHeight(32)
        self.open_log_btn.setStyleSheet("font-size: 10px;")
        self.open_log_btn.clicked.connect(self.open_logs)
        log_sidebar_layout.addWidget(self.open_log_btn)
        
        # Add logging handler
        self.log_handler = QTextEditLogger(self.log_widget)
        self.log_handler.setLevel(logging.INFO)
        self.log_handler.setFormatter(logging.Formatter('%(levelname)s - %(message)s'))
        logger.addHandler(self.log_handler)
        
        # Populate with existing log file contents
        try:
            if log_file.exists():
                raw = log_file.read_text(encoding='utf-8')
                if raw:
                    filtered_lines = []
                    for line in raw.splitlines():
                        # Skip DEBUG lines
                        if ' - DEBUG - ' in line:
                            continue
                        # Remove timestamp (format: "2025-12-07 14:37:25,655 - ")
                        # Match timestamp pattern and remove it
                        import re
                        line_without_timestamp = re.sub(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},\d{3} - ', '', line)
                        filtered_lines.append(line_without_timestamp)
                    
                    filtered = "\n".join(filtered_lines)
                    self.log_widget.setPlainText(filtered)
                    # Auto-scroll to bottom
                    self.log_widget.moveCursor(QtGui.QTextCursor.MoveOperation.End)
        except Exception as e:
            logger.debug(f"Failed to populate log sidebar: {e}")

        main_layout.addWidget(sidebar)
        main_layout.addWidget(content, 1)
        main_layout.addWidget(log_sidebar)

        # Connections
        btn_diag.clicked.connect(lambda: self.switch_page(0, btn_diag))
        btn_setup.clicked.connect(lambda: self.switch_page(1, btn_setup))
        btn_vhd.clicked.connect(lambda: self.switch_page(2, btn_vhd))
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

    def create_defender_rules(self):
        """Create Windows Defender exclusion rules in a background thread."""
        try:
            logger.info("Manual Defender rules creation initiated")
            threading.Thread(target=self._run_defender_rules_creation, daemon=True).start()
        except Exception as e:
            logger.error(f"Failed to start Defender rules creation thread: {e}")
            self.manual_defender_finished.emit(False, f"Failed to start: {e}")

    def _run_defender_rules_creation(self):
        """Run PowerShell script to add Defender exclusions (background thread)."""
        try:
            defender_paths = [
                r"C:\AWRoot",
                r"C:\INSTALL",
                r"C:\Program Files (x86)\PSA VCI",
                r"C:\Program Files\PSA VCI",
                r"C:\Windows\VCX.dll",
            ]
            # Build PowerShell script to add any missing exclusions and return JSON
            ps_paths = ",".join(["'{}'".format(p) for p in defender_paths])
            ps_script = (
                "try { $existing=(Get-MpPreference).ExclusionPath; $added=@(); $failed=@();"
                + "foreach($p in @(" + ps_paths + ")) { if($existing -notcontains $p) { try { Add-MpPreference -ExclusionPath $p; $added += $p } catch { $failed += $p } } }"
                + "$res = @{added=$added; failed=$failed}; $res | ConvertTo-Json -Compress } catch { Write-Error $_; exit 1 }"
            )
            cmd = ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script]
            logger.info("Creating Defender exclusions via PowerShell (manual)")
            proc = subprocess.run(cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
            if proc.returncode == 0:
                out = (proc.stdout or '').strip()
                try:
                    import json
                    j = json.loads(out) if out else {"added":[], "failed":[]}
                    added = j.get('added') or []
                    failed = j.get('failed') or []
                    if failed:
                        msg = translator.t('messages.defender.failed', paths=", ".join(failed))
                        logger.warning(msg)
                        self.manual_defender_finished.emit(False, msg)
                    else:
                        if added:
                            msg = translator.t('messages.defender.success', count=len(added))
                        else:
                            msg = translator.t('messages.defender.no_changes')
                        logger.info(msg)
                        self.manual_defender_finished.emit(True, msg)
                except Exception as e:
                    logger.error(f"Failed to parse Defender PowerShell output: {e}")
                    self.manual_defender_finished.emit(False, f"Parse error: {e}")
            else:
                err = (proc.stderr or '').strip()
                logger.error(f"PowerShell exited with code {proc.returncode}: {err}")
                msg = translator.t('messages.defender.error', error=err[:200])
                self.manual_defender_finished.emit(False, msg)
        except Exception as e:
            logger.error(f"Failed to create Defender exclusions: {e}", exc_info=True)
            self.manual_defender_finished.emit(False, str(e))

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
        
        # Update global banner for new page
        QtCore.QTimer.singleShot(50, self.update_global_banner)
        
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

    def _sanitize_version_for_filename(self, version_str: str) -> str:
        """Return a safe filename base for a given version string.
        Prefer the first numeric sequence (e.g. '09.186') if present; otherwise
        fall back to a sanitized version of the original string.
        """
        try:
            m = re.search(r"(\d+(?:\.\d+)*)", str(version_str))
            if m:
                return m.group(1)
            # Fallback: replace any unsafe filename chars with underscore
            safe = re.sub(r"[^A-Za-z0-9._-]", "_", str(version_str))
            return safe
        except Exception:
            return str(version_str)

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
        version vs the latest available and show simple info message."""
        try:
            installed = self.check_installed_version()
            latest = self.get_latest_available_version()
            
            # Update Diagbox language selector visibility based on installed version
            self._update_diagbox_language_visibility(installed is not None)
            
            if not latest:
                if hasattr(self, 'install_version_info'):
                    self.install_version_info.setVisible(False)
                return

            # Prepare simple info text depending on whether Diagbox is installed
            if not installed:
                text = translator.t('messages.install.no_installed', latest=latest)
                self.install_version_info.setStyleSheet("color: #f0ad4e;")  # Orange for warning
            else:
                cmp = self.compare_versions(latest, installed)
                if cmp == 1:
                    # latest is newer than installed
                    text = translator.t('messages.install.available', latest=latest, installed=installed)
                    self.install_version_info.setStyleSheet("color: #5cb85c;")  # Green for update available
                else:
                    # Installed is up-to-date or newer
                    text = translator.t('messages.install.up_to_date', latest=latest)
                    self.install_version_info.setStyleSheet("color: #5cb85c;")  # Green for up-to-date

            # Show simple info message above combo
            if hasattr(self, 'install_version_info'):
                self.install_version_info.setText(text)
                self.install_version_info.setVisible(True)
                
        except Exception as e:
            logger.debug(f"Error while checking install page updates: {e}")

    def _update_diagbox_language_visibility(self, show):
        """Show or hide the Diagbox language selector based on whether Diagbox is installed"""
        try:
            if hasattr(self, 'diagbox_lang_label'):
                self.diagbox_lang_label.setVisible(show)
            if hasattr(self, 'language_combo'):
                self.language_combo.setVisible(show)
        except Exception as e:
            logger.debug(f"Error updating Diagbox language visibility: {e}")
    
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
        # Row with Re-check system, Install runtimes and Defender Rules buttons
        recheck = QtWidgets.QPushButton(translator.t('buttons.recheck'))
        recheck.setFixedWidth(160)
        recheck.clicked.connect(self.check_system)

        runtimes_btn = QtWidgets.QPushButton(translator.t('buttons.install_runtimes'))
        runtimes_btn.setFixedWidth(160)
        runtimes_btn.clicked.connect(self.install_runtimes)
        # keep a reference so we can enable/disable it during install
        self.runtimes_btn = runtimes_btn

        defender_btn = QtWidgets.QPushButton(translator.t('buttons.defender_rules'))
        defender_btn.setFixedWidth(160)
        defender_btn.clicked.connect(self.create_defender_rules)
        self.defender_btn = defender_btn

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.setSpacing(8)
        btn_row.addWidget(recheck)
        btn_row.addWidget(runtimes_btn)
        btn_row.addWidget(defender_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        return w

    def page_install(self):
        w = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(w)
        layout.setSpacing(10)

        # Page title
        page_title = QtWidgets.QLabel(translator.t('titles.diagbox_native_install'))
        page_title.setObjectName("pageTitle")
        page_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #5cb85c; padding: 10px;")
        page_title.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(page_title)

        # Fetch online version if not already
        if not self.last_version_diagbox:
            self.fetch_last_version_diagbox()

        # Simple version info label (will be populated by on_enter_install_page)
        self.install_version_info = QtWidgets.QLabel("")
        self.install_version_info.setObjectName("installVersionInfo")
        self.install_version_info.setWordWrap(True)
        self.install_version_info.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.install_version_info.setVisible(False)
        layout.addWidget(self.install_version_info)

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

        # Version selection dropdown (or maintenance message if versions unavailable)
        version_layout = QtWidgets.QHBoxLayout()
        version_label = QtWidgets.QLabel(translator.t('labels.select_version'))

        if not getattr(self, 'version_options', None):
            # Show maintenance message instead of combo when no versions available
            maint_msg = translator.t('messages.download.maintenance')
            self.version_combo = None
            self.version_maintenance = QtWidgets.QLabel(maint_msg)
            self.version_maintenance.setWordWrap(True)
            self.version_maintenance.setStyleSheet('color: #b94a48;')
            self.version_maintenance.setMinimumWidth(200)
            version_layout.addWidget(version_label)
            version_layout.addWidget(self.version_maintenance)
            version_layout.addStretch()
            right.addLayout(version_layout)
            # Disable download/install actions to prevent attempts
            # Buttons are created later; mark flags to disable after creation
            self._disable_actions_due_to_maintenance = True
        else:
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

        # Language selection (will be hidden if no Diagbox installed)
        self.diagbox_lang_layout = QtWidgets.QHBoxLayout()
        lang_label = QtWidgets.QLabel(translator.t('labels.diagbox_language'))
        self.language_combo = QtWidgets.QComboBox()
        self.diagbox_lang_label = lang_label

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
        self.diagbox_lang_layout.addWidget(lang_label)
        self.diagbox_lang_layout.addWidget(self.language_combo)
        self.diagbox_lang_layout.addStretch()
        right.addLayout(self.diagbox_lang_layout)
        
        # Hide language selector initially if no Diagbox installed
        installed = self.check_installed_version()
        self._update_diagbox_language_visibility(installed is not None)

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

        # If maintenance mode due to missing versions, disable download button only
        # Allow install if local files exist
        try:
            if getattr(self, '_disable_actions_due_to_maintenance', False):
                if hasattr(self, 'download_button') and self.download_button is not None:
                    self.download_button.setEnabled(False)
                    self.download_button.setToolTip(translator.t('messages.download.maintenance_tooltip'))
                # Check if local files exist - if yes, allow install
                downloaded_versions = self.check_downloaded_versions()
                if hasattr(self, 'install_button') and self.install_button is not None:
                    if downloaded_versions:
                        self.install_button.setEnabled(True)
                        self.install_button.setToolTip(translator.t('messages.install.local_file_available'))
                    else:
                        self.install_button.setEnabled(False)
                        self.install_button.setToolTip(translator.t('messages.download.maintenance_tooltip'))
                # Also disable banner download action
                try:
                    self.install_banner_download.setEnabled(False)
                except Exception:
                    pass
        except Exception:
            pass

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

    def page_vhd(self):
        w = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(w)
        layout.setSpacing(12)

        # Page title
        page_title = QtWidgets.QLabel(translator.t('titles.diagbox_vhdx_install'))
        page_title.setObjectName("pageTitle")
        page_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #5cb85c; padding: 10px;")
        page_title.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(page_title)

        # Header - Download Size (will be updated dynamically)
        self.vhdx_size_header = QtWidgets.QLabel(translator.t('vhd.labels.download_size'))
        self.vhdx_size_header.setObjectName("sectionHeader")
        self.vhdx_size_header.setWordWrap(True)
        layout.addWidget(self.vhdx_size_header)
        
        # VHD Download URL from config
        self.vhd_download_link = URL_VHD_DOWNLOAD
        
        # Fetch download size on page load
        QtCore.QTimer.singleShot(500, self.fetch_vhdx_download_size)

        # Download server selection
        server_layout = QtWidgets.QHBoxLayout()
        server_label = QtWidgets.QLabel(translator.t('vhd.labels.download_server'))
        server_label.setStyleSheet("font-size: 12px;")
        self.vhdx_server_combo = QtWidgets.QComboBox()
        self.vhdx_server_combo.setMinimumWidth(200)
        self.vhdx_server_combo.addItem(translator.t('vhd.servers.server2'), userData="torrent")
        self.vhdx_server_combo.addItem(translator.t('vhd.servers.server1'), userData="direct")
        server_layout.addWidget(server_label)
        server_layout.addWidget(self.vhdx_server_combo)
        server_layout.addStretch()
        layout.addLayout(server_layout)

        # Destination disk selection
        disk_layout = QtWidgets.QHBoxLayout()
        disk_label = QtWidgets.QLabel(translator.t('vhd.labels.destination_disk'))
        disk_label.setStyleSheet("font-size: 12px;")
        self.vhdx_disk_combo = QtWidgets.QComboBox()
        self.vhdx_disk_combo.setMinimumWidth(150)
        # Populate with available drives
        self.populate_vhdx_drives()
        # Connect combo box change to update config display
        self.vhdx_disk_combo.currentIndexChanged.connect(self.update_vhdx_config)
        disk_layout.addWidget(disk_label)
        disk_layout.addWidget(self.vhdx_disk_combo)
        disk_layout.addStretch()
        layout.addLayout(disk_layout)

        # Auto installation toggle
        auto_install_layout = QtWidgets.QHBoxLayout()
        auto_install_label = QtWidgets.QLabel(translator.t('labels.auto_install'))
        auto_install_label.setStyleSheet("font-size: 12px;")
        self.vhdx_auto_install_toggle = QtWidgets.QCheckBox()
        self.vhdx_auto_install_toggle.setChecked(False)
        auto_install_layout.addStretch()
        auto_install_layout.addWidget(self.vhdx_auto_install_toggle)
        auto_install_layout.addWidget(auto_install_label)
        layout.addLayout(auto_install_layout)

        # Action buttons - Télécharger and Installer
        btn_layout = QtWidgets.QHBoxLayout()
        
        self.vhdx_download_btn = QtWidgets.QPushButton(translator.t('buttons.download'))
        self.vhdx_download_btn.setMinimumHeight(44)
        self.vhdx_download_btn.setObjectName("actionButton")
        self.vhdx_download_btn.setMinimumWidth(200)
        self.vhdx_download_btn.clicked.connect(self.download_vhdx)
        
        self.vhdx_install_btn = QtWidgets.QPushButton(translator.t('buttons.install'))
        self.vhdx_install_btn.setMinimumHeight(44)
        self.vhdx_install_btn.setObjectName("actionButton")
        self.vhdx_install_btn.setMinimumWidth(200)
        self.vhdx_install_btn.clicked.connect(self.install_vhdx)
        
        btn_layout.addWidget(self.vhdx_download_btn)
        btn_layout.addWidget(self.vhdx_install_btn)
        layout.addLayout(btn_layout)

        # Pause and Cancel buttons for VHD (hidden by default)
        vhd_buttons_row = QtWidgets.QHBoxLayout()
        
        self.vhd_pause_button = QtWidgets.QPushButton(translator.t('buttons.pause'))
        self.vhd_pause_button.setMinimumHeight(44)
        self.vhd_pause_button.setObjectName("actionButton")
        self.vhd_pause_button.setStyleSheet("background-color: #f0ad4e; color: white;")
        self.vhd_pause_button.clicked.connect(self.toggle_pause_vhd_download)
        self.vhd_pause_button.setVisible(False)
        vhd_buttons_row.addWidget(self.vhd_pause_button)
        
        self.vhd_cancel_button = QtWidgets.QPushButton(translator.t('buttons.cancel'))
        self.vhd_cancel_button.setMinimumHeight(44)
        self.vhd_cancel_button.setObjectName("actionButton")
        self.vhd_cancel_button.setStyleSheet("background-color: #d9534f; color: white;")
        self.vhd_cancel_button.clicked.connect(self.cancel_vhd_download)
        self.vhd_cancel_button.setVisible(False)
        vhd_buttons_row.addWidget(self.vhd_cancel_button)
        
        layout.addLayout(vhd_buttons_row)

        # BCD Cleanup button
        bcd_cleanup_layout = QtWidgets.QHBoxLayout()
        bcd_cleanup_layout.addStretch()
        
        self.bcd_cleanup_btn = QtWidgets.QPushButton(translator.t('buttons.remove_bcd_entries'))
        self.bcd_cleanup_btn.setMinimumHeight(40)
        self.bcd_cleanup_btn.setObjectName("warningButton")
        self.bcd_cleanup_btn.setStyleSheet("background-color: #d9534f; color: white; font-size: 11px;")
        self.bcd_cleanup_btn.setMinimumWidth(250)
        self.bcd_cleanup_btn.clicked.connect(self.remove_bcd_entries)
        self.bcd_cleanup_btn.setToolTip(translator.t('tooltips.remove_bcd_entries'))
        
        bcd_cleanup_layout.addWidget(self.bcd_cleanup_btn)
        bcd_cleanup_layout.addStretch()
        layout.addLayout(bcd_cleanup_layout)

        # Separator line
        separator = QtWidgets.QFrame()
        separator.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        separator.setStyleSheet("background-color: #3a3a3a;")
        layout.addWidget(separator)

        # PC Configuration section
        config_label = QtWidgets.QLabel(translator.t('labels.system_config'))
        config_label.setObjectName("sectionHeader")
        config_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(config_label)

        # Configuration details
        config_frame = QtWidgets.QFrame()
        config_layout = QtWidgets.QFormLayout(config_frame)
        config_layout.setLabelAlignment(QtCore.Qt.AlignmentFlag.AlignRight)
        config_layout.setFormAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
        config_layout.setHorizontalSpacing(20)
        
        self.vhdx_windows_label = QtWidgets.QLabel()
        self.vhdx_storage_label = QtWidgets.QLabel()
        self.vhdx_ram_label = QtWidgets.QLabel()
        
        config_layout.addRow(translator.t('labels.windows_version'), self.vhdx_windows_label)
        config_layout.addRow(translator.t('labels.free_storage'), self.vhdx_storage_label)
        config_layout.addRow(translator.t('labels.ram'), self.vhdx_ram_label)
        
        layout.addWidget(config_frame)

        # Update configuration display
        self.update_vhdx_config()

        layout.addStretch()
        return w

    def populate_vhdx_drives(self):
        """Populate the drive combo box with available drives"""
        try:
            import string
            from ctypes import windll
            
            self.vhdx_disk_combo.clear()
            drives = []
            bitmask = windll.kernel32.GetLogicalDrives()
            for letter in string.ascii_uppercase:
                if bitmask & 1:
                    drives.append(letter)
                bitmask >>= 1
            
            for drive in drives:
                try:
                    usage = psutil.disk_usage(f"{drive}:\\")
                    free_gb = usage.free / (1024**3)
                    self.vhdx_disk_combo.addItem(f"{drive}: ({free_gb:.1f} GB free)", userData=drive)
                except:
                    self.vhdx_disk_combo.addItem(f"{drive}:", userData=drive)
        except Exception as e:
            logger.error(f"Failed to populate drives: {e}")
            self.vhdx_disk_combo.addItem("C:", userData="C")

    def update_vhdx_config(self):
        """Update VHD configuration display"""
        # Windows version
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
        self.vhdx_windows_label.setText(os_text)

        # Storage - check for minimum 50 GB requirement
        try:
            selected_drive = self.vhdx_disk_combo.currentData()
            if selected_drive:
                storage = psutil.disk_usage(f"{selected_drive}:\\")
                free_gb = storage.free / (1024 ** 3)
                self.vhdx_storage_label.setText(f"{free_gb:.1f} GB")
                
                # Mark in red if less than 50 GB
                if free_gb < 50:
                    self.vhdx_storage_label.setStyleSheet("color: red; font-weight: bold;")
                else:
                    self.vhdx_storage_label.setStyleSheet("")
        except:
            self.vhdx_storage_label.setText("N/A")
            self.vhdx_storage_label.setStyleSheet("")

        # RAM
        ram_gb = psutil.virtual_memory().total / (1024 ** 3)
        self.vhdx_ram_label.setText(f"{ram_gb:.1f} GB")

    def fetch_vhdx_download_size(self):
        """Fetch VHDX download size from URL"""
        if not self.vhd_download_link:
            # Default URL or fetch from config
            # TODO: Set actual VHD download link
            self.vhd_download_link = ""  # Placeholder
            self.vhdx_size_header.setText(translator.t('vhd.labels.download_size_default'))
            return
        
        try:
            logger.debug(f"Fetching VHDX file size from: {self.vhd_download_link}")
            response = requests.head(self.vhd_download_link, allow_redirects=True, timeout=10)
            content_length = response.headers.get('Content-Length')
            
            if content_length:
                file_size_bytes = int(content_length)
                file_size_gb = round(file_size_bytes / 1024 / 1024 / 1024, 2)
                
                self.vhdx_size_header.setText(
                    translator.t('vhd.labels.download_size_with_space', size=file_size_gb)
                )
                logger.info(f"VHDX file size: {file_size_gb} GB")
            else:
                self.vhdx_size_header.setText(translator.t('vhd.labels.download_size_default'))
        except Exception as e:
            logger.error(f"Failed to fetch VHDX file size: {e}")
            self.vhdx_size_header.setText(translator.t('vhd.labels.download_size_default'))
    
    def check_vhdx_disk_space(self):
        """Check if selected drive has at least 50 GB free space"""
        try:
            selected_drive = self.vhdx_disk_combo.currentData()
            if selected_drive:
                storage = psutil.disk_usage(f"{selected_drive}:\\")
                free_gb = storage.free / (1024 ** 3)
                return free_gb >= 50
        except Exception as e:
            logger.error(f"Failed to check disk space: {e}")
        return False

    def download_vhdx(self):
        """Download VHDX file"""
        # Check disk space before downloading
        if not self.check_vhdx_disk_space():
            selected_drive = self.vhdx_disk_combo.currentData()
            try:
                storage = psutil.disk_usage(f"{selected_drive}:\\")
                free_gb = storage.free / (1024 ** 3)
                QtWidgets.QMessageBox.warning(
                    self,
                    "Espace insuffisant",
                    f"Espace disponible sur {selected_drive}: : {free_gb:.1f} GO\n\n"
                    f"Espace minimum requis : 50 GO\n\n"
                    f"Veuillez libérer de l'espace ou sélectionner un autre disque."
                )
            except:
                QtWidgets.QMessageBox.warning(
                    self,
                    "Espace insuffisant",
                    "Espace disque insuffisant pour télécharger et installer le VHDX.\n\n"
                    "Espace minimum requis : 50 GO"
                )
            return
        
        # Get selected server mode
        selected_server = self.vhdx_server_combo.currentData()
        logger.info(f"Selected download server: {selected_server}")
        
        # Check if URL is configured
        if selected_server == "direct" and not self.vhd_download_link:
            QtWidgets.QMessageBox.warning(
                self,
                "Configuration manquante",
                "L'URL de téléchargement VHDX n'est pas configurée.\n\n"
                "Veuillez contacter l'administrateur."
            )
            return
        
        # Disable buttons
        self.set_buttons_enabled(False)
        
        # Update footer
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(translator.t('vhd.download.in_progress'))
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setRange(0, 1000)
            self.footer_progress.setValue(0)
            self.footer_progress.setFormat("%p%")
        
        # Get selected drive
        selected_drive = self.vhdx_disk_combo.currentData()
        
        # Show pause and cancel buttons
        if self.vhd_pause_button:
            self.vhd_pause_button.setVisible(True)
            self.vhd_pause_button.setText(translator.t('buttons.pause'))
        if self.vhd_cancel_button:
            self.vhd_cancel_button.setVisible(True)
        
        # Start download thread based on selected server
        if selected_server == "torrent":
            logger.info(f"Starting VHDX torrent download to drive {selected_drive}:")
            self.vhdx_download_thread = TorrentDownloadThread(
                URL_VHD_TORRENT,
                f"{selected_drive}:\\VHD",
                selected_drive,
                target_file="PSA-DIAG.vhdx"
            )
        else:  # direct
            logger.info(f"Starting VHDX direct download to drive {selected_drive}:")
            self.vhdx_download_thread = VHDXDownloadThread(
                self.vhd_download_link,
                f"{selected_drive}:\\VHD",
                selected_drive
            )
        
        self.vhdx_download_thread.progress.connect(self.update_vhdx_progress)
        self.vhdx_download_thread.finished.connect(self.on_vhdx_download_finished)
        self.vhdx_download_thread.start()

    def update_vhdx_progress(self, value, speed, eta):
        """Update VHDX download progress"""
        logger.debug(f"update_vhdx_progress called: value={value}, speed={speed:.2f}, eta={eta}")
        
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setValue(value)
            if speed > 0:
                self.footer_progress.setFormat(f"{value/10:.1f}% - {speed:.2f} MB/s - ETA: {eta}")
            else:
                self.footer_progress.setFormat(f"{value/10:.1f}%")
            logger.debug(f"Progress bar updated: {value/10:.1f}%")
        
        if hasattr(self, 'footer_label'):
            speed_text = f"{speed:.2f} MB/s" if speed > 0 else "Connexion..."
            self.footer_label.setText(f"Vitesse: {speed_text}")
            logger.debug(f"Footer label updated: {speed_text}")

    def on_vhdx_download_finished(self, success, message):
        """Called when VHDX download finishes"""
        # Hide pause and cancel buttons
        if self.vhd_pause_button:
            self.vhd_pause_button.setVisible(False)
            self.vhd_pause_button.setText(translator.t('buttons.pause'))
        if self.vhd_cancel_button:
            self.vhd_cancel_button.setVisible(False)
        
        # Re-enable buttons
        self.set_buttons_enabled(True)
        
        # Update footer
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(message)
        if hasattr(self, 'footer_progress'):
            if success:
                self.footer_progress.setValue(1000)
            else:
                self.footer_progress.setValue(0)
        
        if success:
            QtWidgets.QMessageBox.information(
                self,
                translator.t('vhd.download.success'),
                message
            )
            
            # If auto-install is enabled, start installation
            if hasattr(self, 'vhdx_auto_install_toggle') and self.vhdx_auto_install_toggle.isChecked():
                logger.info("Auto-install enabled, starting VHDX installation")
                QtCore.QTimer.singleShot(1000, self.install_vhdx)
        else:
            QtWidgets.QMessageBox.warning(
                self,
                translator.t('vhd.download.failed'),
                message
            )
        
        # Reset footer after delay
        QtCore.QTimer.singleShot(3000, self.reset_footer)

    def install_vhdx(self):
        """Install VHDX file"""
        # Search for PSA-DIAG.vhdx on all available drives
        vhdx_file = None
        found_drive = None
        
        try:
            import string
            # Get all available drive letters
            available_drives = []
            for letter in string.ascii_uppercase:
                drive = f"{letter}:"
                if os.path.exists(drive):
                    available_drives.append(letter)
            
            logger.info(f"Searching for PSA-DIAG.vhdx on drives: {', '.join(available_drives)}")
            
            # Search for PSA-DIAG.vhdx in X:\VHD\ on all drives
            for letter in available_drives:
                vhd_folder = os.path.join(f"{letter}:\\", "VHD")
                target_file = os.path.join(vhd_folder, "PSA-DIAG.vhdx")
                
                if os.path.exists(target_file):
                    vhdx_file = target_file
                    found_drive = letter
                    logger.info(f"Found VHDX file: {vhdx_file}")
                    break
        except Exception as e:
            logger.error(f"Error searching for VHDX file: {e}")
        
        if not vhdx_file:
            QtWidgets.QMessageBox.warning(
                self,
                "Fichier introuvable",
                "Aucun fichier PSA-DIAG.vhdx trouvé dans X:\\VHD\\ sur aucun lecteur.\n\n"
                "Veuillez d'abord télécharger le fichier VHDX."
            )
            return
        
        # Confirm installation
        reply = QtWidgets.QMessageBox.question(
            self,
            "Confirmation",
            f"Êtes-vous sûr de vouloir installer ce VHDX ?\n\n"
            f"Fichier: {os.path.basename(vhdx_file)}\n"
            f"Chemin: {vhdx_file}\n\n"
            f"Cette opération va modifier votre configuration de démarrage (BCD).\n"
            f"Vous devrez redémarrer votre ordinateur après l'installation.",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No
        )
        
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        
        # Disable buttons
        self.set_buttons_enabled(False)
        
        # Update footer
        if hasattr(self, 'footer_label'):
            self.footer_label.setText("Installation VHDX en cours...")
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setRange(0, 0)  # Indeterminate
        
        # Start installation thread
        logger.info(f"Starting VHDX installation: {vhdx_file}")
        self.vhdx_install_thread = VHDXInstallThread(vhdx_file, "PSA-DIAG")
        self.vhdx_install_thread.finished.connect(self.on_vhdx_install_finished)
        self.vhdx_install_thread.start()

    def on_vhdx_install_finished(self, success, message):
        """Called when VHDX installation finishes"""
        # Re-enable buttons
        self.set_buttons_enabled(True)
        
        # Update footer
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(translator.t('vhd.install.complete') if success else translator.t('vhd.install.failed'))
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setRange(0, 1000)
            self.footer_progress.setValue(1000 if success else 0)
        
        if success:
            QtWidgets.QMessageBox.information(
                self,
                translator.t('vhd.install.success'),
                message
            )
        else:
            QtWidgets.QMessageBox.critical(
                self,
                translator.t('vhd.install.failed'),
                message
            )
        
        # Reset footer after delay
        QtCore.QTimer.singleShot(3000, self.reset_footer)

    def remove_bcd_entries(self):
        """Remove PSA-DIAG entries from BCD with backup"""
        # Confirm with user
        reply = QtWidgets.QMessageBox.question(
            self,
            translator.t('messages.bcd_cleanup.confirm_title'),
            translator.t('messages.bcd_cleanup.confirm_message'),
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No
        )
        
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        
        # Disable buttons
        self.set_buttons_enabled(False)
        
        # Update footer
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(translator.t('messages.bcd_cleanup.in_progress'))
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setRange(0, 0)  # Indeterminate
        
        # Start cleanup thread
        logger.info("Starting BCD cleanup")
        self.bcd_cleanup_thread = BCDCleanupThread()
        self.bcd_cleanup_thread.finished.connect(self.on_bcd_cleanup_finished)
        self.bcd_cleanup_thread.start()
    
    def on_bcd_cleanup_finished(self, success, message):
        """Called when BCD cleanup finishes"""
        # Re-enable buttons
        self.set_buttons_enabled(True)
        
        # Update footer
        if hasattr(self, 'footer_label'):
            self.footer_label.setText(translator.t('messages.bcd_cleanup.complete') if success else translator.t('messages.bcd_cleanup.failed'))
        if hasattr(self, 'footer_progress'):
            self.footer_progress.setRange(0, 1000)
            self.footer_progress.setValue(1000 if success else 0)
        
        if success:
            QtWidgets.QMessageBox.information(
                self,
                translator.t('messages.bcd_cleanup.title'),
                message
            )
        else:
            QtWidgets.QMessageBox.critical(
                self,
                translator.t('messages.bcd_cleanup.title'),
                message
            )
        
        # Reset footer after delay
        QtCore.QTimer.singleShot(3000, self.reset_footer)

    def page_about(self):
        w = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(w)
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # Changelog frame (full width)
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
        changelog_layout.addWidget(self.changelog_text)
        
        layout.addWidget(changelog_frame)
        
        # Load changelog asynchronously
        QtCore.QTimer.singleShot(500, self.load_changelog)
        
        return w
    
    def load_changelog(self):
        """Load changelog from the last 10 GitHub releases"""
        try:
            # Get the latest 10 releases
            api_url = "https://api.github.com/repos/RetroGameSets/PSA-DIAG/releases?per_page=10"
            
            logger.info(f"[STEP 3] -- Fetching changelog for last 10 releases")
            response = requests.get(api_url, timeout=15)
            
            if response.status_code == 200:
                releases = response.json()
                
                if not releases:
                    self.changelog_text.setPlainText("No releases found.")
                    return
                
                changelog_parts = []
                
                for release in releases:
                    tag = release.get('tag_name', 'Unknown')
                    name = release.get('name', tag)
                    body = release.get('body', '')
                    published = release.get('published_at', '')
                    
                    # Format published date
                    date_str = ''
                    if published:
                        try:
                            from datetime import datetime
                            dt = datetime.fromisoformat(published.replace('Z', '+00:00'))
                            date_str = dt.strftime('%Y-%m-%d')
                        except:
                            date_str = published.split('T')[0]
                    
                    # Add version header
                    header = f"{'='*50}\n"
                    if date_str:
                        header += f"📦 {name} ({date_str})\n"
                    else:
                        header += f"📦 {name}\n"
                    header += f"{'='*50}\n"
                    
                    changelog_parts.append(header)
                    
                    if body:
                        # Extract the Changes section if present
                        if '### Changes' in body:
                            changes_section = body.split('### Changes', 1)[1].strip()
                            # Remove any trailing markdown sections
                            if '###' in changes_section:
                                changes_section = changes_section.split('###', 1)[0].strip()
                            changelog_parts.append(changes_section + "\n")
                        else:
                            # Use entire body, but skip Installation section if present
                            if '### Installation' in body:
                                # Try to get everything after Installation section
                                parts = body.split('### Installation', 1)
                                if len(parts) > 1 and '###' in parts[1]:
                                    remaining = parts[1].split('###', 1)[1]
                                    changelog_parts.append(remaining.strip() + "\n")
                                else:
                                    changelog_parts.append(body.strip() + "\n")
                            else:
                                changelog_parts.append(body.strip() + "\n")
                    else:
                        changelog_parts.append("No changelog available.\n")
                    
                    changelog_parts.append("\n")
                
                full_changelog = "\n".join(changelog_parts)
                self.changelog_text.setPlainText(full_changelog)
                logger.info(f"Changelog loaded successfully ({len(releases)} releases)")
                
            elif response.status_code == 404:
                self.changelog_text.setPlainText("Releases not found on GitHub.\n\nThis repository may not have any releases yet.")
                logger.warning("Releases not found (404)")
            else:
                self.changelog_text.setPlainText(f"Failed to load changelog (HTTP {response.status_code})")
                logger.warning(f"Failed to fetch changelog: HTTP {response.status_code}")
                
        except requests.exceptions.RequestException as e:
            logger.warning(f"Unable to load changelog: {type(e).__name__}: {str(e)}")
            self.changelog_text.setPlainText("Unable to load changelog.\n\nPlease check your network connection and try again.")
        except Exception as e:
            logger.error(f"Error loading changelog: {type(e).__name__}: {e}")
            self.changelog_text.setPlainText(f"Error loading changelog:\n{str(e)}")
        finally:
            # Close splash screen once all loading is done
            QtCore.QTimer.singleShot(100, self._close_splash_screen)
    
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
        """Clean up resources before closing"""
        # Kill any running aria2c processes
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.info['name'] and proc.info['name'].lower() == 'aria2c.exe':
                        logger.info(f"Terminating aria2c.exe (PID: {proc.pid})")
                        proc.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        except Exception as e:
            logger.error(f"Error killing aria2c processes: {e}")
        
        # Cancel any running download threads
        if hasattr(self, 'vhdx_download_thread') and self.vhdx_download_thread:
            if self.vhdx_download_thread.isRunning():
                logger.info("Cancelling VHDX download thread")
                self.vhdx_download_thread.cancel()
                self.vhdx_download_thread.wait(2000)
        
        if hasattr(self, 'download_thread') and self.download_thread:
            if self.download_thread.isRunning():
                logger.info("Cancelling download thread")
                self.download_thread.cancel()
                self.download_thread.wait(2000)
        
        event.accept()
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
    
    # Set AppUserModelID for Windows taskbar icon
    if sys.platform == 'win32':
        try:
            myappid = 'PSA_DIAG'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            logger.info(f"AppUserModelID set to: {myappid}")
        except Exception as e:
            logger.warning(f"Failed to set AppUserModelID: {e}")
    
    app = QtWidgets.QApplication([])
    
    # Set application icon
    icon_path = BASE / "icons" / "icon.ico"
    if icon_path.exists():
        app.setWindowIcon(QtGui.QIcon(str(icon_path)))
    
    # Show splash screen during initialization
    splash = SplashScreen()
    splash.show()
    app.processEvents()  # Process events to display the splash screen
    
    win = MainWindow(splash=splash)  # Pass splash screen reference to MainWindow
    win.show()
    sys.exit(app.exec())
