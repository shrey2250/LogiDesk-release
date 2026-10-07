"""
updater.py — Auto-update module for LogiDesk
===========================================
Handles checking GitHub Releases for new versions, displaying a styled CustomTkinter
update dialog with changelogs, downloading updates with progress tracking, and performing
safe Windows executable replacement and restart.
"""

import os
import sys
import json
import tempfile
import threading
import subprocess
import urllib.request
import urllib.error
import customtkinter as ctk
from tkinter import messagebox

# Current local application version
APP_VERSION = "1.0.0"

# Target Public GitHub repository for releases
GITHUB_REPO = "shrey2250/LogiDesk-release"


def _parse_version(v_str: str):
    """Parses a version string like 'v1.2.3' or '1.2.0' into a tuple of ints."""
    cleaned = v_str.strip().lstrip("vV")
    parts = []
    for part in cleaned.split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts)


def check_for_updates_async(parent_window, repo: str = GITHUB_REPO, silent: bool = True):
    """
    Checks for updates in a background thread.
    - If silent=True: only displays UI if an update is actually available.
    - If silent=False: alerts user even if the app is already up to date.
    """
    def _worker():
        if not repo or "your-username" in repo:
            # Repo placeholder not configured yet
            if not silent:
                parent_window.after(0, lambda: messagebox.showinfo(
                    "Check for Updates",
                    f"LogiDesk is currently v{APP_VERSION}.\n\nPlease configure your GitHub Releases repository in updater.py."
                ))
            return

        try:
            api_url = f"https://api.github.com/repos/{repo}/releases/latest"
            req = urllib.request.Request(
                api_url,
                headers={"User-Agent": "LogiDesk-AppUpdater", "Accept": "application/vnd.github.v3+json"}
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                if resp.status != 200:
                    return
                data = json.loads(resp.read().decode("utf-8"))

            tag_name = data.get("tag_name", "")
            remote_ver = _parse_version(tag_name)
            local_ver = _parse_version(APP_VERSION)

            if remote_ver > local_ver:
                # Find the Windows .exe asset in the release
                exe_asset = None
                for asset in data.get("assets", []):
                    name = asset.get("name", "").lower()
                    if name.endswith(".exe"):
                        exe_asset = asset
                        break

                download_url = exe_asset["browser_download_url"] if exe_asset else data.get("html_url")
                changelog = data.get("body", "No release notes provided.")

                # Prompt user on main UI thread
                parent_window.after(
                    0,
                    lambda: UpdateDialog(
                        parent=parent_window,
                        current_version=APP_VERSION,
                        latest_version=tag_name,
                        changelog=changelog,
                        download_url=download_url,
                        asset_name=exe_asset["name"] if exe_asset else "LogiDesk.exe"
                    )
                )
            else:
                if not silent:
                    parent_window.after(
                        0,
                        lambda: messagebox.showinfo(
                            "Up to Date",
                            f"You are using the latest version of LogiDesk (v{APP_VERSION})."
                        )
                    )
        except Exception as err:
            if not silent:
                parent_window.after(
                    0,
                    lambda: messagebox.showerror(
                        "Update Check Failed",
                        f"Could not check for updates.\n\nError: {err}"
                    )
                )

    thread = threading.Thread(target=_worker, daemon=True)
    thread.start()


class UpdateDialog(ctk.CTkToplevel):
    """A modal dialog styled with Apple HIG aesthetics to download & install updates."""

    def __init__(self, parent, current_version, latest_version, changelog, download_url, asset_name):
        super().__init__(parent)
        self.parent = parent
        self.download_url = download_url
        self.asset_name = asset_name
        self.is_downloading = False

        self.title("Software Update Available")
        self.geometry("480x420")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self.transient(parent)

        # Center on parent window
        self.update_idletasks()
        try:
            px = parent.winfo_x() + (parent.winfo_width() // 2) - 240
            py = parent.winfo_y() + (parent.winfo_height() // 2) - 210
            self.geometry(f"+{max(0, px)}+{max(0, py)}")
        except Exception:
            pass

        self._build_ui(current_version, latest_version, changelog)

    def _build_ui(self, current_ver, new_ver, changelog):
        # Header Badge
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(20, 10))

        title_lbl = ctk.CTkLabel(
            header,
            text=f"A new version of LogiDesk is available!",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            anchor="w"
        )
        title_lbl.pack(anchor="w")

        sub_lbl = ctk.CTkLabel(
            header,
            text=f"LogiDesk {new_ver} is now available (you have v{current_ver}).",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#3C3C43", "#8E8E93"),
            anchor="w"
        )
        sub_lbl.pack(anchor="w", pady=(2, 0))

        # Changelog preview
        cl_box = ctk.CTkTextbox(
            self,
            height=180,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            wrap="word",
            border_width=1,
            border_color=("#D1D1D6", "#38383A")
        )
        cl_box.pack(fill="both", expand=True, padx=24, pady=10)
        cl_box.insert("1.0", f"Release Notes for {new_ver}:\n\n{changelog.strip()}")
        cl_box.configure(state="disabled")

        # Progress bar (hidden initially)
        self.prog_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.prog_frame.pack(fill="x", padx=24, pady=(0, 10))

        self.status_lbl = ctk.CTkLabel(
            self.prog_frame,
            text="",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#3C3C43", "#8E8E93")
        )
        self.status_lbl.pack(anchor="w")

        self.progress_bar = ctk.CTkProgressBar(self.prog_frame)
        self.progress_bar.set(0)

        # Buttons
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=24, pady=(0, 20))

        self.skip_btn = ctk.CTkButton(
            self.btn_frame,
            text="Remind Me Later",
            fg_color="transparent",
            text_color=("#007AFF", "#0A84FF"),
            hover_color=("#E5E5EA", "#2C2C2E"),
            command=self.destroy
        )
        self.skip_btn.pack(side="left")

        self.update_btn = ctk.CTkButton(
            self.btn_frame,
            text="Update & Restart",
            fg_color=("#007AFF", "#0A84FF"),
            hover_color=("#0062CC", "#3395FF"),
            command=self._start_download
        )
        self.update_btn.pack(side="right")

    def _start_download(self):
        if self.is_downloading:
            return
        self.is_downloading = True
        self.skip_btn.configure(state="disabled")
        self.update_btn.configure(state="disabled", text="Downloading...")
        self.progress_bar.pack(fill="x", pady=(4, 0))

        threading.Thread(target=self._download_and_install, daemon=True).start()

    def _download_and_install(self):
        try:
            temp_dir = tempfile.gettempdir()
            target_file = os.path.join(temp_dir, f"update_{self.asset_name}")

            req = urllib.request.Request(
                self.download_url,
                headers={"User-Agent": "LogiDesk-AppUpdater"}
            )

            with urllib.request.urlopen(req) as resp, open(target_file, "wb") as out:
                total_size = int(resp.info().get("Content-Length", 0))
                downloaded = 0
                block_size = 65536

                while True:
                    buffer = resp.read(block_size)
                    if not buffer:
                        break
                    downloaded += len(buffer)
                    out.write(buffer)

                    if total_size > 0:
                        pct = downloaded / total_size
                        mb_done = downloaded / (1024 * 1024)
                        mb_total = total_size / (1024 * 1024)
                        self.after(
                            0,
                            lambda p=pct, d=mb_done, t=mb_total: self._update_progress(p, f"Downloading: {d:.1f} MB / {t:.1f} MB ({int(p*100)}%)")
                        )

            self.after(0, lambda: self._apply_update(target_file))

        except Exception as e:
            self.after(0, lambda err=e: self._handle_download_error(err))

    def _update_progress(self, pct, status_text):
        self.progress_bar.set(pct)
        self.status_lbl.configure(text=status_text)

    def _handle_download_error(self, err):
        self.is_downloading = False
        self.skip_btn.configure(state="normal")
        self.update_btn.configure(state="normal", text="Retry Update")
        self.status_lbl.configure(text="Download failed.")
        messagebox.showerror("Update Error", f"Failed to download the update:\n\n{err}")

    def _apply_update(self, new_exe_path):
        """Creates a batch script to replace the current running executable and restart."""
        self.status_lbl.configure(text="Restarting and applying update...")

        current_exe = sys.executable
        # If running from python script directly during dev, do not overwrite python.exe
        if not getattr(sys, "frozen", False):
            messagebox.showinfo(
                "Development Mode",
                f"Update downloaded successfully to:\n{new_exe_path}\n\n(File swap is skipped in script dev mode)."
            )
            self.destroy()
            return

        batch_path = os.path.join(tempfile.gettempdir(), "logidesk_update.bat")
        batch_script = f"""@echo off
timeout /t 2 /nobreak > NUL
:retry
move /y "{new_exe_path}" "{current_exe}" > NUL 2>&1
if errorlevel 1 (
    timeout /t 1 /nobreak > NUL
    goto retry
)
start "" "{current_exe}"
del "%~f0"
"""
        with open(batch_path, "w") as f:
            f.write(batch_script)

        # Launch the batch script silently without window
        subprocess.Popen(
            ["cmd.exe", "/c", batch_path],
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
            close_fds=True
        )

        # Exit current application
        self.parent.quit()
        sys.exit(0)
