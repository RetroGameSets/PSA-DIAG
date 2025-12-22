from pathlib import Path
import os
# Central configuration used by main and build-time metadata
CONFIG_DIR = Path(os.getenv('APPDATA', Path.home())) / 'PSA_DIAG'

# Application version (keep in sync with release tags)
APP_VERSION = "2.3.1.1"

# Archive extraction password used by main.py (kept empty by default)
ARCHIVE_PASSWORD = ""

# Remote endpoints
URL_LAST_VERSION_PSADIAG = "https://psa-diag.fr/diagbox/install/last_version_psadiag.json"
URL_LAST_VERSION_DIAGBOX = "https://psa-diag.fr/diagbox/install/last_version_diagbox.json"
URL_VERSION_OPTIONS = "https://psa-diag.fr/diagbox/install/available_versions.json"
URL_REMOTE_MESSAGES = "https://psa-diag.fr/diagbox/install/banner.json"

# VHD/VHDX download URL
URL_VHD_DOWNLOAD = "https://archive.org/download/psadiag/PSA-DIAG.vhdx"
URL_VHD_TORRENT = "https://archive.org/download/psadiag/psadiag_archive.torrent"
