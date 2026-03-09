import sys
import subprocess
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox


APP_TITLE = "Customer Report Tool"
SCRIPT_NAMES = {
    "compare_all": "Customer_Attend_All.py",
    "compare_recent": "Customer_Attend_Recent.py",
    "city_points": "City_Point_Filter.py",
}
FOLDERS = ["Reports_Import", "City_Points", "Reports_Output"]


class CustomerReportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("900x650")
        self.minsize(820, 560)

        self.repo_root = Path(__file__).resolve().parent
        self.is_running = False

        self._ensure_folders_exist()
        self._build_ui()
        self._refresh_status()

    def _ensure_folders_exist(self):
        for folder in FOLDERS:
            (self.repo_root / folder).mkdir(parents=True, exist_ok=True)

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        header = ttk.Frame(self, padding=(16, 16, 16, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        title = ttk.Label(
            header,
            text="Customer Report Tool",
            font=("Segoe UI", 16, "bold"),
        )
        title.grid(row=0, column=0, sticky="w")

        subtitle = ttk.Label(
            header,
            text=(
                "Run your attendance and City Points scripts from one place. "
                "All scripts are expected to be in the same folder as this app. "
                "The app automatically creates Reports_Import, City_Points, and Reports_Output if they do not exist."
            ),
            wraplength=760,
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(6, 0))

        top = ttk.Frame(self, padding=(16, 8, 16, 8))
        top.grid(row=1, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        top.columnconfigure(1, weight=1)

        actions = ttk.LabelFrame(top, text="Actions", padding=12)
        actions.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        actions.columnconfigure(0, weight=1)

        self.btn_compare_all = ttk.Button(
            actions,
            text="Run Compare All Games",
            command=lambda: self._run_script("compare_all"),
        )
        self.btn_compare_all.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.btn_compare_recent = ttk.Button(
            actions,
            text="Run Compare Two Most Recent Games",
            command=lambda: self._run_script("compare_recent"),
        )
        self.btn_compare_recent.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        self.btn_city_points = ttk.Button(
            actions,
            text="Run City Points Export",
            command=lambda: self._run_script("city_points"),
        )
        self.btn_city_points.grid(row=2, column=0, sticky="ew")

        status_box = ttk.LabelFrame(top, text="Project Status", padding=12)
        status_box.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        status_box.columnconfigure(1, weight=1)

        self.status_labels = {}
        for idx, folder in enumerate(FOLDERS):
            ttk.Label(
                status_box,
                text=f"{folder}:"
            ).grid(row=idx, column=0, sticky="w", padx=(0, 8), pady=2)

            lbl = ttk.Label(status_box, text="Checking...")
            lbl.grid(row=idx, column=1, sticky="w", pady=2)
            self.status_labels[folder] = lbl

        ttk.Label(
            status_box,
            text="Scripts:"
        ).grid(row=len(FOLDERS), column=0, sticky="nw", padx=(0, 8), pady=(8, 2))

        self.scripts_status = ttk.Label(
            status_box,
            text="Checking...",
            justify="left"
        )
        self.scripts_status.grid(row=len(FOLDERS), column=1, sticky="w", pady=(8, 2))

        controls = ttk.Frame(self, padding=(16, 0, 16, 8))
        controls.grid(row=2, column=0, sticky="nsew")
        controls.columnconfigure(0, weight=1)
        controls.rowconfigure(1, weight=1)

        info_bar = ttk.Frame(controls)
        info_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        info_bar.columnconfigure(0, weight=1)

        self.current_action_var = tk.StringVar(value="Ready")
        self.current_action_label = ttk.Label(
            info_bar,
            textvariable=self.current_action_var
        )
        self.current_action_label.grid(row=0, column=0, sticky="w")

        self.refresh_button = ttk.Button(
            info_bar,
            text="Refresh Status",
            command=self._refresh_status
        )
        self.refresh_button.grid(row=0, column=1, sticky="e")

        log_frame = ttk.LabelFrame(controls, text="Output Log", padding=10)
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame,
            wrap="word",
            height=18,
            font=("Consolas", 10)
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        footer = ttk.Frame(self, padding=(16, 0, 16, 16))
        footer.grid(row=3, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        ttk.Label(
            footer,
            text=(
                "Tip: put your Excel imports into Reports_Import or City_Points, "
                "then run the matching action above."
            ),
        ).grid(row=0, column=0, sticky="w")

    def _append_log(self, text):
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.update_idletasks()

    def _set_running_state(self, running):
        self.is_running = running
        state = "disabled" if running else "normal"

        self.btn_compare_all.configure(state=state)
        self.btn_compare_recent.configure(state=state)
        self.btn_city_points.configure(state=state)
        self.refresh_button.configure(state=state)

    def _refresh_status(self):
        for folder in FOLDERS:
            path = self.repo_root / folder

            if path.exists() and path.is_dir():
                file_count = len(list(path.iterdir()))
                self.status_labels[folder].configure(
                    text=f"Found ({file_count} item(s))"
                )
            else:
                self.status_labels[folder].configure(text="Missing")

        lines = []
        for _, script_name in SCRIPT_NAMES.items():
            script_path = self.repo_root / script_name
            status = "Found" if script_path.exists() else "Missing"
            lines.append(f"{script_name}: {status}")

        self.scripts_status.configure(text="\n".join(lines))

    def _validate_before_run(self, action_key):
        script_name = SCRIPT_NAMES[action_key]
        script_path = self.repo_root / script_name

        if not script_path.exists():
            messagebox.showerror(
                "Missing Script",
                f"Could not find {script_name} in:\n{self.repo_root}"
            )
            return None

        if action_key in {"compare_all", "compare_recent"}:
            folder = self.repo_root / "Reports_Import"
            if not folder.exists():
                messagebox.showerror(
                    "Missing Folder",
                    f"Missing folder:\n{folder}"
                )
                return None

        elif action_key == "city_points":
            folder = self.repo_root / "City_Points"
            if not folder.exists():
                messagebox.showerror(
                    "Missing Folder",
                    f"Missing folder:\n{folder}"
                )
                return None

        return script_path

    def _run_script(self, action_key):
        if self.is_running:
            return

        script_path = self._validate_before_run(action_key)
        if script_path is None:
            return

        thread = threading.Thread(
            target=self._run_script_worker,
            args=(action_key, script_path),
            daemon=True
        )
        thread.start()

    def _run_script_worker(self, action_key, script_path):
        action_names = {
            "compare_all": "Compare All Games",
            "compare_recent": "Compare Two Most Recent Games",
            "city_points": "City Points Export",
        }
        action_name = action_names[action_key]

        self.after(0, lambda: self._set_running_state(True))
        self.after(0, lambda: self.current_action_var.set(f"Running: {action_name}"))
        self.after(
            0,
            lambda: self._append_log(
                f"\n{'=' * 70}\nRunning {action_name}\nScript: {script_path.name}\n{'=' * 70}\n"
            )
        )

        try:
            process = subprocess.Popen(
                [sys.executable, str(script_path)],
                cwd=str(self.repo_root),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )

            if process.stdout is not None:
                for line in process.stdout:
                    self.after(0, lambda line=line: self._append_log(line))

            return_code = process.wait()

            if return_code == 0:
                self.after(
                    0,
                    lambda: self._append_log(
                        f"\nFinished successfully: {action_name}\n"
                    )
                )
                self.after(
                    0,
                    lambda: self.current_action_var.set(f"Finished: {action_name}")
                )
            else:
                self.after(
                    0,
                    lambda: self._append_log(
                        f"\nFailed with exit code {return_code}: {action_name}\n"
                    )
                )
                self.after(
                    0,
                    lambda: self.current_action_var.set(f"Failed: {action_name}")
                )
                self.after(
                    0,
                    lambda: messagebox.showerror(
                        "Script Failed",
                        f"{action_name} failed. Check the output log for details."
                    )
                )

        except Exception as exc:
            self.after(0, lambda: self._append_log(f"\nError: {exc}\n"))
            self.after(0, lambda: self.current_action_var.set("Error"))
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))

        finally:
            self.after(0, self._refresh_status)
            self.after(0, lambda: self._set_running_state(False))


if __name__ == "__main__":
    app = CustomerReportApp()
    app.mainloop()