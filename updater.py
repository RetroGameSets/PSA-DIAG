import argparse
import os
import shutil
import sys
import time
import subprocess


def wait_for_pid(pid, timeout=30):
    start = time.time()
    while time.time() - start < timeout:
        try:
            # On Windows, os.kill with 0 doesn't exist; use psutil if available
            import psutil
            if not psutil.pid_exists(pid):
                return True
        except Exception:
            # Fallback: try opening process via tasklist (best-effort)
            try:
                proc = subprocess.run(["tasklist", "/FI", f"PID eq {pid}"], capture_output=True, text=True)
                if str(pid) not in proc.stdout:
                    return True
            except Exception:
                return True

        time.sleep(0.5)
    return False


def replace_file(src, dst, timeout=30):
    start = time.time()
    last_exc = None
    while time.time() - start < timeout:
        try:
            # Try atomic replace
            if os.path.exists(dst):
                os.replace(src, dst)
            else:
                shutil.move(src, dst)
            return True
        except Exception as e:
            last_exc = e
            time.sleep(0.5)
    raise last_exc


def main():
    parser = argparse.ArgumentParser(description="PSA-DIAG updater helper")
    parser.add_argument("--target", required=True, help="Path to target exe to replace")
    parser.add_argument("--new", required=True, help="Path to new exe file downloaded")
    parser.add_argument("--wait-pid", type=int, default=0, help="PID to wait for before replacing")
    parser.add_argument("--restart", action="store_true", help="Relaunch target after replace")
    parser.add_argument("--timeout", type=int, default=60, help="Timeout in seconds for waiting/replace")

    args = parser.parse_args()

    try:
        if args.wait_pid:
            waited = wait_for_pid(args.wait_pid, timeout=args.timeout)
            if not waited:
                # attempt to continue anyway
                pass

        # Attempt to replace
        replace_file(args.new, args.target, timeout=args.timeout)

        # Optionally restart
        if args.restart:
            try:
                subprocess.Popen([args.target], creationflags=(0x08000000 if os.name == 'nt' else 0))
            except Exception:
                pass

        sys.exit(0)
    except Exception as e:
        print(f"Updater failed: {e}")
        sys.exit(2)


if __name__ == "__main__":
    main()
