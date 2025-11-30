from pathlib import Path
import os

# Central configuration used by main and build-time metadata
CONFIG_DIR = Path(os.getenv('APPDATA', Path.home())) / 'PSA_DIAG'
# Do NOT create the directory at import time; building with PyInstaller
# imports this module during the spec execution and creating folders or
# touching the filesystem can fail under some build contexts. Create the
# directory lazily at runtime where needed (main.py already creates logs).

# Application version (keep in sync with release tags)
APP_VERSION = "2.1.0.7"

# Remote endpoints
URL_LAST_VERSION_PSADIAG = "https://psa-diag.fr/diagbox/install/last_version_psadiag.json"
URL_LAST_VERSION_DIAGBOX = "https://psa-diag.fr/diagbox/install/last_version_diagbox.json"
