from pathlib import Path
import os

# Central configuration used by main and build-time metadata
CONFIG_DIR = Path(os.getenv('APPDATA', Path.home())) / 'PSA_DIAG'

# Application version (keep in sync with release tags)
APP_VERSION = "2.1.0.9"

# Remote endpoints
URL_LAST_VERSION_PSADIAG = "https://psa-diag.fr/diagbox/install/last_version_psadiag.json"
URL_LAST_VERSION_DIAGBOX = "https://psa-diag.fr/diagbox/install/last_version_diagbox.json"
URL_VERSION_OPTIONS = "https://psa-diag.fr/diagbox/install/available_versions.json"
# Remote banner/messages JSON (multilingual). Expected format: array of
# objects with `id`, `lang` (map of language codes to text/link), optional
# `link`, `start`/`end` timestamps, etc.
URL_REMOTE_MESSAGES = "https://psa-diag.fr/diagbox/install/banner.json"
