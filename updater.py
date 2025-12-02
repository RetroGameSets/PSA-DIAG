import argparse
import os
import shutil
import sys
import time
import subprocess
import threading

try:
    import tkinter as tk
    from tkinter import scrolledtext
    from tkinter import ttk
except Exception:
    tk = None


def wait_for_pid(pid, timeout=30, logger=None):
    start = time.time()
    # Give the process a moment to start closing
    time.sleep(0.5)
    
    # Check if process exists at all first
    process_exists = False
    try:
        import psutil
        process_exists = psutil.pid_exists(pid)
    except Exception:
        try:
            proc = subprocess.run(["tasklist", "/FI", f"PID eq {pid}"], capture_output=True, text=True)
            process_exists = str(pid) in proc.stdout
        except Exception:
            pass
    
    # If process doesn't exist at all, return immediately
    if not process_exists:
        if logger:
            logger(f"Target process {pid} already closed.")
        return True
    
    # Wait for process to exit
    while time.time() - start < timeout:
        try:
            import psutil
            if not psutil.pid_exists(pid):
                if logger:
                    logger(f"Process {pid} has exited.")
                return True
        except Exception:
            try:
                proc = subprocess.run(["tasklist", "/FI", f"PID eq {pid}"], capture_output=True, text=True)
                if str(pid) not in proc.stdout:
                    if logger:
                        logger(f"Process {pid} has exited.")
                    return True
            except Exception:
                # If we can't check, assume it's gone
                return True

        if logger:
            logger(f"Please wait, updating...")
        time.sleep(0.5)
    
    if logger:
        logger(f"Timeout waiting for process {pid} to exit.")
    return False


def replace_file(src, dst, timeout=30, logger=None):
    start = time.time()
    last_exc = None
    while time.time() - start < timeout:
        try:
            # Try atomic replace
            if os.path.exists(dst):
                os.replace(src, dst)
            else:
                shutil.move(src, dst)
            if logger:
                logger(f"Replaced {dst} with {src}")
            return True
        except Exception as e:
            last_exc = e
            # On permission errors, attempt to remove the destination first and retry once
            if logger:
                logger(f"Replace attempt failed: {e}")
            try:
                if isinstance(e, PermissionError) or ('Access is denied' in str(e)):
                    if os.path.exists(dst):
                        if logger:
                            logger(f"Attempting to remove destination before retry: {dst}")
                        try:
                            os.remove(dst)
                            if logger:
                                logger(f"Removed destination {dst}; will retry move")
                            # Try move/replace immediately
                            if os.path.exists(dst):
                                # still exists, give up this retry
                                pass
                            else:
                                if os.path.exists(dst):
                                    pass
                                # If dst no longer exists, try moving
                                if os.path.exists(dst):
                                    pass
                        except Exception as rm_exc:
                            if logger:
                                logger(f"Failed to remove destination {dst}: {rm_exc}")
                # small backoff before next attempt
            except Exception:
                pass
            time.sleep(0.5)
    raise last_exc


def processes_holding_path(path, logger=None):
    """Return list of (pid, name) for processes that have `path` open or whose executable matches path."""
    holders = []
    try:
        import psutil
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                # check exe path
                try:
                    exe = proc.exe()
                    if exe and os.path.normcase(exe) == os.path.normcase(path):
                        holders.append((proc.pid, proc.name()))
                        continue
                except Exception:
                    pass

                # check open files
                try:
                    for f in proc.open_files():
                        if os.path.normcase(f.path) == os.path.normcase(path):
                            holders.append((proc.pid, proc.name()))
                            break
                except Exception:
                    pass
            except Exception:
                pass
    except Exception:
        # psutil not available or error; cannot detect holders
        if logger:
            logger('psutil not available; skipping holder detection')
    return holders




class UpdaterGUI:
    def __init__(self, args):
        self.args = args
        self.root = None
        self.thread = None

        if tk:
            # Minimal centered window
            self.root = tk.Tk()
            self.root.title("PSA-DIAG Updater")
            self.root.resizable(False, False)

            # Build simple layout: label + indeterminate progressbar + close button (disabled until finish)
            pad = 16
            frame = tk.Frame(self.root, padx=pad, pady=pad)
            frame.pack(fill='both', expand=True)

            self.message_var = tk.StringVar(value="PSA-DIAG update in progress, please wait...")
            lbl = tk.Label(frame, textvariable=self.message_var, font=(None, 11))
            lbl.pack(pady=(0, 12))

            self.progress = ttk.Progressbar(frame, mode='indeterminate', length=320)
            self.progress.pack(pady=(0, 12))
            self.progress.start(10)

            btn_frame = tk.Frame(frame)
            btn_frame.pack(fill='x')
            self.close_btn = tk.Button(btn_frame, text='Close', command=self.on_close, state='disabled')
            self.close_btn.pack(side='right')

            # Center the window on screen
            self.root.update_idletasks()
            w = self.root.winfo_width()
            h = self.root.winfo_height()
            ws = self.root.winfo_screenwidth()
            hs = self.root.winfo_screenheight()
            x = (ws // 2) - (w // 2)
            y = (hs // 2) - (h // 2)
            self.root.geometry(f'+{x}+{y}')
            self.root.attributes('-topmost', True)

            # Start work in background thread
            self.thread = threading.Thread(target=self.run)
            self.thread.daemon = True
            self.thread.start()

            # Start Tk mainloop
            self.root.protocol("WM_DELETE_WINDOW", lambda: None)  # disable direct close
            self.root.mainloop()
        else:
            # No tkinter available: fallback to console prints
            self.log = print
            self.status = lambda s: print(s)
            self.run()

    def log(self, msg):
        if self.root:
            def append():
                self.log_widget.configure(state='normal')
                self.log_widget.insert('end', msg + '\n')
                self.log_widget.see('end')
                self.log_widget.configure(state='disabled')
            self.root.after(0, append)
        else:
            print(msg)

    def status(self, msg):
        if self.root:
            self.root.after(0, lambda: self.status_var.set(msg))
        else:
            print(msg)

    def on_retry(self):
        # retry behavior kept for compatibility (not exposed in minimal UI)
        if self.thread and not self.thread.is_alive():
            self.thread = threading.Thread(target=self.run)
            self.thread.daemon = True
            self.thread.start()

    def on_close(self):
        try:
            if self.root:
                self.root.destroy()
        except Exception:
            pass
        sys.exit(0)

    def run(self):
        args = self.args
        # Update UI status
        if tk:
            self.message_var.set('Waiting for target to exit...' if args.wait_pid else 'Preparing to replace')
        else:
            self.status('Waiting for target to exit...' if args.wait_pid else 'Preparing to replace')

        if args.wait_pid:
            if tk:
                self.log(f"Waiting for PID {args.wait_pid} to exit (timeout {args.timeout}s)")
            ok = wait_for_pid(args.wait_pid, timeout=args.timeout, logger=self.log if tk else None)
            if not ok:
                if tk:
                    self.log(f"Timeout waiting for PID {args.wait_pid}; continuing anyway")
                else:
                    self.status(f"Timeout waiting for PID {args.wait_pid}; continuing anyway")

        # After the PID has exited, give the OS a short moment

        # Additional check: wait until no process holds the destination file (handles released)
        dst = args.target
        max_handle_wait = 10
        waited = 0
        try:
            holders = processes_holding_path(dst, logger=self.log)
            while holders and waited < max_handle_wait:
                self.log(f"Detected processes holding target: {holders}; attempting to terminate if they match target exe")
                try:
                    import psutil
                    for pid, name in holders:
                        try:
                            p = psutil.Process(pid)
                            # Only attempt to terminate if the process executable equals the target path
                            try:
                                if os.path.normcase(p.exe()) == os.path.normcase(dst):
                                    self.log(f"Terminating lingering process PID={pid} name={name}")
                                    p.terminate()
                                    try:
                                        p.wait(timeout=2)
                                    except Exception:
                                        p.kill()
                            except Exception:
                                # couldn't read exe; attempt gentle terminate
                                p.terminate()
                        except Exception as e:
                            self.log(f"Could not terminate PID {pid}: {e}")
                except Exception:
                    pass

                time.sleep(1)
                waited += 1
                holders = processes_holding_path(dst, logger=self.log)
            if holders:
                self.log(f"Handles still held after waiting {max_handle_wait}s: {holders}")
        except Exception as e:
            self.log(f"Error while checking/terminating holders: {e}")

        try:
            if tk:
                self.message_var.set('Replacing executable...')
            else:
                self.status('Replacing executable...')

            if tk:
                self.log(f"Replacing {args.target} with {args.new}")
            replace_file(args.new, args.target, timeout=args.timeout, logger=self.log if tk else None)

            if tk:
                self.message_var.set('Update OK, PSA-DIAG will restart...')
                try:
                    self.progress.stop()
                except Exception:
                    pass
                try:
                    self.close_btn.configure(state='normal')
                except Exception:
                    pass
            else:
                self.status('Replacement successful')

            if args.restart:
                if tk:
                    self.log('Relaunching target...')
                try:
                    subprocess.Popen([args.target], creationflags=(0x08000000 if os.name == 'nt' else 0))
                    if tk:
                        self.log('Relaunch requested')
                except Exception as e:
                    if tk:
                        self.log(f'Failed to relaunch: {e}')
                    else:
                        self.status(f'Failed to relaunch: {e}')

            if tk:
                self.log('Updater finished successfully')
            else:
                self.status('Done')
        except Exception as e:
            self.status('Error')
            self.log(f'Updater failed: {e}')
            self.log('You can fix the problem and click Retry')
            if self.root:
                self.root.after(0, lambda: self.retry_btn.configure(state='normal'))
            else:
                print('Updater failed; exiting with code 2')
                sys.exit(2)


def main():
    parser = argparse.ArgumentParser(description="PSA-DIAG updater helper")
    parser.add_argument("--target", required=True, help="Path to target exe to replace")
    parser.add_argument("--new", required=True, help="Path to new exe file downloaded")
    parser.add_argument("--wait-pid", type=int, default=0, help="PID to wait for before replacing")
    parser.add_argument("--restart", action="store_true", help="Relaunch target after replace")
    parser.add_argument("--timeout", type=int, default=60, help="Timeout in seconds for waiting/replace")

    args = parser.parse_args()

    gui = UpdaterGUI(args)


if __name__ == "__main__":
    main()
