"""
PSA-DIAG FREE - PySide6 skeleton
Run:
    pip install -r requirements.txt
    python main.py
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

BASE = Path(__file__).resolve().parent

def is_admin():
    """Check if the script is running with admin privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False

def run_as_admin():
    """Relaunch the script with admin privileges"""
    try:
        if sys.platform == 'win32':
            # Get the path to the Python executable and the script
            script = os.path.abspath(sys.argv[0])
            params = ' '.join([script] + sys.argv[1:])
            
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
        print(f"Failed to elevate privileges: {e}")
        return False
    return False

# Load style
def load_qss():
    qss_path = BASE / "ui" / "style.qss"
    try:
        if qss_path.exists():
            return qss_path.read_text()
    except Exception as e:
        print(f"Erreur lors du chargement du style QSS : {e}")
    return ""

class SidebarButton(QtWidgets.QPushButton):
    def __init__(self, text, icon_path=None, parent=None):
        super().__init__(text, parent)
        self.setCheckable(True)
        self.setMinimumHeight(48)
        self.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        if icon_path and icon_path.exists():
            icon = QtGui.QIcon(str(icon_path))
            self.setIcon(icon)
            self.setIconSize(QtCore.QSize(22,22))


import os

class DownloadThread(QtCore.QThread):
    progress = QtCore.Signal(int, float, str)  # value, speed_mbs, eta_str
    finished = QtCore.Signal(bool, str)

    def __init__(self, url, path, last_version, total_size=0):
        super().__init__()
        self.url = url
        self.path = path
        self.last_version = last_version
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
            response = requests.get(self.url, stream=True)
            response.raise_for_status()
            downloaded = 0
            chunk_count = 0
            start_time = time.time()
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
                        self.finished.emit(False, f"Download Diagbox {self.last_version} cancelled")
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
                                progress = 0
                                eta_str = "--:--"
                            self.progress.emit(progress, speed, eta_str)
            if self.total_size == 0 or downloaded >= self.total_size:
                self.progress.emit(1000, 0, "00:00")
            if os.path.exists(self.path):
                self.finished.emit(True, f"Download Diagbox {self.last_version} ok")
            else:
                self.finished.emit(False, f"Download Diagbox {self.last_version} failed")
        except Exception as e:
            self.finished.emit(False, f"Download Diagbox {self.last_version} failed: {e}")

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
            extraction_errors = []
            
            # Use 7z.exe for much faster extraction
            seven_zip_exe = BASE / "tools" / "7z.exe"
            
            if not seven_zip_exe.exists():
                self.finished.emit(False, f"7z.exe not found at {seven_zip_exe}")
                return
            
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
                        print(f"7z output: {output.strip()}")
                        
                        # 7z outputs progress like "1% 2909 - filename"
                        output = output.strip()
                        if '%' in output and ' - ' in output:
                            try:
                                # Extract percentage from beginning of line
                                percent_str = output.split('%')[0].strip()
                                if percent_str.isdigit():
                                    percent = int(percent_str)
                                    print(f"Progress extracted: {percent}%")
                                    self.progress.emit(percent)
                                
                                # Extract filename after " - "
                                filename = output.split(' - ', 1)[1] if ' - ' in output else ''
                                if filename:
                                    self.file_progress.emit(filename)
                            except Exception as e:
                                print(f"Error parsing progress: {e}")
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
                message = f"Installation completed with warnings:\n\n{error_summary}\n\nDiagbox has been installed to C:."
                self.finished.emit(True, message)
            else:
                self.finished.emit(True, "Diagbox installed successfully to C:.")
                
        except Exception as e:
            self.finished.emit(False, f"Installation failed: {e}")

class MainWindow(QtWidgets.QWidget):
    download_finished = QtCore.Signal(bool, str)  # success, message

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PSA-DIAG FREE")
        self.setWindowFlags(QtCore.Qt.WindowType.FramelessWindowHint)
        self.resize(900, 420)

        # Download variables
        self.download_folder = "C:\\INSTALL\\"
        self.last_version = ""
        self.diagbox_path = ""
        self.auto_install = None
        self.download_thread = None
        self.install_thread = None
        self.cancel_button = None
        self.pause_button = None
        self.dragPos = QtCore.QPoint()
        
        # Fetch last version first
        self.fetch_last_version()
        
        # Version options: (display_name, version, url)
        self.version_options = [
            (f"Diagbox {self.last_version} (Latest)", self.last_version, f"https://archive.batocera.org/recalbox.remix_old_website/web/composer/Diagbox_Install_{self.last_version}.7z"),
            ("Diagbox 9.85", "9.85", "https://archive.batocera.org/recalbox.remix_old_website/web/composer/Diagbox_Install_9.85.7z")
        ]

        # Connect signals
        self.download_finished.connect(self.on_download_finished)

        self.setup_ui()

    def update_progress(self, value, speed, eta):
        bar = self.stack.currentWidget().findChild(QtWidgets.QProgressBar)
        if bar:
            bar.setValue(value)
            bar.setFormat(f"{value / 10:.1f}% - {speed:.1f} MB/s - {eta}")
            bar.repaint()
            QtWidgets.QApplication.processEvents()

    def on_download_finished(self, success, message):
        QtWidgets.QMessageBox.information(self, "Download", message)
        
        # Hide cancel and pause buttons
        if self.cancel_button:
            self.cancel_button.setVisible(False)
        if self.pause_button:
            self.pause_button.setVisible(False)
            self.pause_button.setText("Pause")
        
        # Re-enable all buttons and combo box
        self.set_buttons_enabled(True)
        
        bar = self.stack.currentWidget().findChild(QtWidgets.QProgressBar)
        if bar:
            bar.setRange(0, 1000)
            bar.setValue(1000 if success else 0)
            if success:
                bar.setFormat("100.0% - Download Complete")
            else:
                bar.setFormat("0.0% - Download Failed")
        if success and self.auto_install and self.auto_install.isChecked():
            self.install_diagbox()

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
                    self.pause_button.setText("Pause")
            else:
                self.download_thread.pause()
                if self.pause_button:
                    self.pause_button.setText("Resume")

    def on_install_finished(self, success, message, install_button, bar):
        QtWidgets.QMessageBox.information(self, "Install", message)
        
        # Re-enable all buttons and combo box
        self.set_buttons_enabled(True)
        
        if bar:
            bar.setRange(0, 100)
            bar.setValue(100 if success else 0)
            bar.setFormat("Installation complete" if success else "Installation failed")

    def update_install_progress(self, value):
        """Update installation progress bar"""
        bar = self.stack.currentWidget().findChild(QtWidgets.QProgressBar)
        if bar:
            bar.setValue(value)
            bar.setFormat(f"Extracting... {value}%")
            bar.repaint()
            QtWidgets.QApplication.processEvents()
    
    def update_install_file(self, filename):
        """Update current file being extracted"""
        if hasattr(self, 'file_label'):
            # Truncate filename if too long
            if len(filename) > 80:
                filename = "..." + filename[-77:]
            self.file_label.setText(filename)
            self.file_label.repaint()
            QtWidgets.QApplication.processEvents()

    def install_diagbox(self):
        # Get selected version from combo box
        if hasattr(self, 'version_combo'):
            selected_data = self.version_combo.currentData()
            if selected_data:
                version, url = selected_data
                self.last_version = version
                self.diagbox_path = os.path.join(self.download_folder, f"Diagbox_Install_{version}.7z")
        
        if not os.path.exists(self.diagbox_path):
            # Get the version being attempted
            version = self.last_version if self.last_version else "Unknown"
            QtWidgets.QMessageBox.warning(
                self, 
                "Install", 
                f"Diagbox file not found.\n\n"
                f"Version: {version}\n"
                f"Expected path: {self.diagbox_path}\n\n"
                f"Please download this version first."
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
        
        # Set progress bar to extracting
        bar = self.stack.currentWidget().findChild(QtWidgets.QProgressBar)
        if bar:
            bar.setRange(0, 100)
            bar.setValue(0)
            bar.setFormat("Extracting... 0%")
        # Start installation in thread
        self.install_thread = InstallThread(self.diagbox_path)
        self.install_thread.progress.connect(self.update_install_progress)
        self.install_thread.file_progress.connect(self.update_install_file)
        self.install_thread.finished.connect(lambda success, message: self.on_install_finished(success, message, install_button, bar))
        self.install_thread.start()

    def clean_diagbox(self):
        """Clean Diagbox installation by removing C:\\APP, C:\\AWRoot, and C:\\APPLIC folders"""
        # Check which folders exist
        folders_to_delete = []
        folders = [r"C:\APP", r"C:\AWRoot", r"C:\APPLIC"]
        
        for folder in folders:
            if os.path.exists(folder):
                folders_to_delete.append(folder)
        
        if not folders_to_delete:
            QtWidgets.QMessageBox.information(
                self,
                "Clean Diagbox",
                "No Diagbox folders found to clean.\n\nFolders checked:\n- C:\\APP\n- C:\\AWRoot\n- C:\\APPLIC"
            )
            return
        
        # Confirm deletion
        folder_list = "\n".join([f"- {folder}" for folder in folders_to_delete])
        reply = QtWidgets.QMessageBox.question(
            self,
            "Clean Diagbox",
            f"This will permanently delete the following folders:\n\n{folder_list}\n\nAre you sure?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No
        )
        
        if reply == QtWidgets.QMessageBox.StandardButton.No:
            return
        
        # Disable all buttons
        self.set_buttons_enabled(False)
        
        # Kill all Diagbox processes before cleaning
        self.kill_diagbox_processes_silent()
        
        # Delete folders
        success_count = 0
        failed_folders = []
        
        for folder in folders_to_delete:
            try:
                shutil.rmtree(folder)
                success_count += 1
            except Exception as e:
                failed_folders.append(f"{folder}: {str(e)}")
        
        # Re-enable buttons
        self.set_buttons_enabled(True)
        
        # Show result
        if failed_folders:
            error_list = "\n".join(failed_folders)
            QtWidgets.QMessageBox.warning(
                self,
                "Clean Diagbox",
                f"Cleaned {success_count} folder(s) successfully.\n\nFailed to delete:\n{error_list}"
            )
        else:
            QtWidgets.QMessageBox.information(
                self,
                "Clean Diagbox",
                f"Successfully cleaned {success_count} Diagbox folder(s)."
            )

    def install_vci_driver(self):
        """Install VCI Driver using ACTIAPnPInstaller.exe"""
        driver_path = r"C:\AWRoot\extra\Drivers\xsevo\ACTIAPnPInstaller.exe"
        
        # Check if the installer exists
        if not os.path.exists(driver_path):
            QtWidgets.QMessageBox.warning(
                self,
                "Install VCI Driver",
                f"VCI Driver installer not found.\n\nExpected location:\n{driver_path}\n\nPlease install Diagbox first."
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
                    "Install VCI Driver",
                    "VCI Driver installed successfully."
                )
            else:
                QtWidgets.QMessageBox.warning(
                    self,
                    "Install VCI Driver",
                    f"VCI Driver installation returned code: {result.returncode}\n\nThe driver may already be installed or require a reboot."
                )
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                "Install VCI Driver",
                f"Failed to install VCI Driver:\n\n{str(e)}"
            )

    def launch_diagbox(self):
        """Launch Diagbox application"""
        diagbox_exe = r"C:\AWRoot\bin\launcher\Diagbox.exe"
        
        # Check if Diagbox.exe exists
        if not os.path.exists(diagbox_exe):
            QtWidgets.QMessageBox.warning(
                self,
                "Launch Diagbox",
                f"Diagbox executable not found.\n\nExpected location:\n{diagbox_exe}\n\nPlease install Diagbox first."
            )
            return
        
        try:
            # Launch Diagbox.exe
            subprocess.Popen([diagbox_exe], cwd=r"C:\AWRoot\bin\launcher")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                "Launch Diagbox",
                f"Failed to launch Diagbox:\n\n{str(e)}"
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
                    "Kill/Close Process",
                    f"Successfully killed {killed_count} process(es)."
                )
            else:
                QtWidgets.QMessageBox.information(
                    self,
                    "Kill/Close Process",
                    "No Diagbox processes found running."
                )
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                "Kill/Close Process",
                f"Error while trying to kill processes:\n\n{str(e)}"
            )

    def fetch_last_version(self):
        try:
            response = requests.get("https://archive.batocera.org/recalbox.remix_old_website/web/composer/last_version.txt")
            response.raise_for_status()
            self.last_version = response.text.strip()
            self.diagbox_path = os.path.join(self.download_folder, f"Diagbox_Install_{self.last_version}.7z")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", f"Failed to fetch last version: {e}")

    def download_diagbox(self):
        # Get selected version from combo box
        if hasattr(self, 'version_combo'):
            selected_data = self.version_combo.currentData()
            if selected_data:
                version, url = selected_data
                self.last_version = version
            else:
                if not self.last_version:
                    self.fetch_last_version()
                if not self.last_version:
                    return
                url = f"https://archive.batocera.org/recalbox.remix_old_website/web/composer/Diagbox_Install_{self.last_version}.7z"
        else:
            if not self.last_version:
                self.fetch_last_version()
            if not self.last_version:
                return
            url = f"https://archive.batocera.org/recalbox.remix_old_website/web/composer/Diagbox_Install_{self.last_version}.7z"
        
        self.diagbox_path = os.path.join(self.download_folder, f"Diagbox_Install_{self.last_version}.7z")
        
        if not os.path.exists(self.download_folder):
            os.makedirs(self.download_folder)
        
        file_path = self.diagbox_path
        
        # Get total size for progress
        try:
            head = requests.head(url)
            total_size = int(head.headers.get('content-length', 0))
        except:
            total_size = 0
        
        # Check if file already exists and size matches
        if os.path.exists(file_path):
            if total_size > 0 and os.path.getsize(file_path) == total_size:
                QtWidgets.QMessageBox.information(self, "Download", f"Diagbox {self.last_version} already downloaded.")
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
            self.pause_button.setText("Pause")
        
        # Set progress bar
        bar = self.stack.currentWidget().findChild(QtWidgets.QProgressBar)
        if bar:
            if total_size > 0:
                bar.setRange(0, 1000)
                bar.setValue(0)
                bar.setFormat("0.0% - 0.0 MB/s - --:--")
            else:
                bar.setRange(0, 0)  # indeterminate
        self.download_thread = DownloadThread(url, file_path, self.last_version, total_size)
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
        sidebar.setFixedWidth(84)
        vbox = QtWidgets.QVBoxLayout(sidebar)
        vbox.setContentsMargins(10,10,10,10)
        vbox.setSpacing(12)

        # Title in sidebar
        title_sidebar = QtWidgets.QLabel("PSA-DIAG FREE")
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
        btn_min = QtWidgets.QPushButton("-")
        btn_close = QtWidgets.QPushButton("X")
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

        header = QtWidgets.QLabel("System Configuration")
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
        form.addRow("Windows Version :", self.os_label)
        form.addRow("Free storage :", self.storage_label)
        form.addRow("RAM :", self.ram_label)
        layout.addWidget(reqs)

        layout.addStretch()
        recheck = QtWidgets.QPushButton("Re-check system")
        recheck.setFixedWidth(160)
        recheck.clicked.connect(self.check_system)
        layout.addWidget(recheck, 0, QtCore.Qt.AlignmentFlag.AlignLeft)

        return w

    def page_install(self):
        w = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(w)
        layout.setSpacing(10)

        # Fetch online version if not already
        if not self.last_version:
            self.fetch_last_version()

        installed_version = self.check_installed_version()
        version_text = installed_version if installed_version else "Not installed"
        header_installed = QtWidgets.QLabel(f"Installed Version : {version_text}")
        header_installed.setObjectName("sectionHeader")
        layout.addWidget(header_installed)

        header_online = QtWidgets.QLabel(f"Online Version : {self.last_version if self.last_version else 'Unknown'}")
        header_online.setObjectName("sectionHeader")
        layout.addWidget(header_online)

        # Check downloaded versions
        downloaded_versions = self.check_downloaded_versions()
        if downloaded_versions:
            downloaded_text = ", ".join([f"{v['version']} ({v['size_mb']:.1f} MB)" for v in downloaded_versions])
            header_downloaded = QtWidgets.QLabel(f"Downloaded : {downloaded_text}")
            header_downloaded.setObjectName("sectionHeader")
            header_downloaded.setStyleSheet("color: #5cb85c;")
            layout.addWidget(header_downloaded)

        sub = QtWidgets.QHBoxLayout()
        left = QtWidgets.QVBoxLayout()
        right = QtWidgets.QVBoxLayout()

        # Version selection dropdown
        version_layout = QtWidgets.QHBoxLayout()
        version_label = QtWidgets.QLabel("Select Version :")
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
        lbl = QtWidgets.QLabel("Auto Install after download :")
        toggle = QtWidgets.QCheckBox()
        self.auto_install = toggle
        h.addWidget(lbl)
        h.addWidget(toggle)
        h.addStretch()
        right.addLayout(h)

        # Buttons grid
        grid = QtWidgets.QGridLayout()
        grid.setSpacing(12)
        btns = [
            ("Download", "download"),
            ("Install", "install"),
            ("Clean Diagbox", "clean"),
            ("Install VCI Driver", "vci"),
            ("Launch Diagbox", "launch"),
            ("Kill/Close Process", "kill"),
        ]
        for i, (txt, _) in enumerate(btns):
            b = QtWidgets.QPushButton(txt)
            b.setMinimumHeight(44)
            b.setObjectName("actionButton")
            if txt == "Download":
                self.download_button = b
                b.clicked.connect(self.download_diagbox)
            elif txt == "Install":
                b.clicked.connect(self.install_diagbox)
            elif txt == "Clean Diagbox":
                b.clicked.connect(self.clean_diagbox)
            elif txt == "Install VCI Driver":
                b.clicked.connect(self.install_vci_driver)
            elif txt == "Launch Diagbox":
                b.clicked.connect(self.launch_diagbox)
            elif txt == "Kill/Close Process":
                b.clicked.connect(self.kill_diagbox)
            grid.addWidget(b, i//3, i%3)

        right.addLayout(grid)

        # Pause and Cancel buttons (hidden by default)
        buttons_row = QtWidgets.QHBoxLayout()
        
        self.pause_button = QtWidgets.QPushButton("Pause")
        self.pause_button.setMinimumHeight(44)
        self.pause_button.setObjectName("actionButton")
        self.pause_button.setStyleSheet("background-color: #f0ad4e; color: white;")
        self.pause_button.clicked.connect(self.toggle_pause_download)
        self.pause_button.setVisible(False)
        buttons_row.addWidget(self.pause_button)
        
        self.cancel_button = QtWidgets.QPushButton("Cancel Download")
        self.cancel_button.setMinimumHeight(44)
        self.cancel_button.setObjectName("actionButton")
        self.cancel_button.setStyleSheet("background-color: #d9534f; color: white;")
        self.cancel_button.clicked.connect(self.cancel_download)
        self.cancel_button.setVisible(False)
        buttons_row.addWidget(self.cancel_button)
        
        right.addLayout(buttons_row)

        # Progress bar
        pb = QtWidgets.QProgressBar()
        pb.setValue(0)
        pb.setTextVisible(True)
        
        # Label for current file being extracted
        self.file_label = QtWidgets.QLabel("")
        self.file_label.setStyleSheet("color: #888; font-size: 11px;")
        self.file_label.setWordWrap(False)
        
        layout.addLayout(sub)
        layout.addLayout(right)
        layout.addWidget(pb)
        layout.addWidget(self.file_label)

        return w

    def page_update(self):
        w = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(w)
        layout.addWidget(QtWidgets.QLabel("Update / Refresh (placeholder)"))
        layout.addStretch()
        return w

    def page_about(self):
        w = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(w)
        logo = QtWidgets.QLabel()
        pix = QtGui.QPixmap(str(BASE / "icons" / "logo.png"))
        if not pix.isNull():
            pix = pix.scaledToWidth(160, QtCore.Qt.TransformationMode.SmoothTransformation)
            logo.setPixmap(pix)
        else:
            logo.setText("Logo non disponible")
        layout.addWidget(logo, 0, QtCore.Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(QtWidgets.QLabel("PSA-DIAG FREE"))
        layout.addWidget(QtWidgets.QLabel("Version: 1.0.0"))
        layout.addWidget(QtWidgets.QLabel("Developed by Mike"))
        layout.addStretch()
        return w
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
                        return
                    current = current.parent()
            
            self.dragPos = event.globalPosition().toPoint()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == QtCore.Qt.MouseButton.LeftButton:
            if hasattr(self, 'dragPos'):
                delta = event.globalPosition().toPoint() - self.dragPos
                self.move(self.pos() + delta)
                self.dragPos = event.globalPosition().toPoint()
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
    # Check if running as admin, if not relaunch with admin privileges
    if not is_admin():
        print("Not running as admin, requesting elevation...")
        if run_as_admin():
            sys.exit(0)  # Exit current instance
        else:
            # If elevation failed, continue anyway (user might have cancelled)
            print("Continuing without admin privileges (some features may not work)")
    
    app = QtWidgets.QApplication([])
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
