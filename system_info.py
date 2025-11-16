"""
system_info.py - utilities to detect system information.
Fill functions with actual implementation (psutil, platform) when needed.
"""
import platform
def get_windows_version():
    return platform.system() + " " + platform.release()

def get_ram_total_gb():
    try:
        import psutil
        return round(psutil.virtual_memory().total / (1024**3), 1)
    except Exception:
        return None

def get_free_storage_gb(path="C:\\"):
    try:
        import shutil
        total, used, free = shutil.disk_usage(path)
        return round(free / (1024**3), 1)
    except Exception:
        return None
