import math
import sys
import threading
import traceback
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk


APP_TITLE = "Customer Report Tool"
FOLDERS = ["Reports_Import", "City_Points", "Reports_Output"]
TOOLS = [
    "Compare All Games",
    "Compare Two Most Recent Games",
    "City Points Export",
]


class SeatOccupiedComparer:
    def __init__(
        self,
        import_folder="Reports_Import",
        output_folder="Reports_Output",
        compare_column="Seat Occupied",
    ):
        self.import_folder = Path(import_folder)
        self.output_folder = Path(output_folder)
        self.compare_column = compare_column
        self.seat_key_columns = ["Area Name", "Block", "Seat Row", "Seat Number"]

    def get_input_files(self):
        files = sorted(self.import_folder.glob("*.xlsx"))
        if len(files) < 2:
            raise ValueError(
                f"At least 2 Excel files are required in '{self.import_folder}'."
            )
        return files

    @staticmethod
    def classify_game_type(date_value):
        if pd.isna(date_value):
            return "Unknown"
        dt = pd.to_datetime(date_value, errors="coerce")
        if pd.isna(dt):
            return "Unknown"
        day_name = dt.day_name()
        if day_name in ["Tuesday", "Wednesday", "Thursday"]:
            return "Midweek"
        if day_name in ["Friday", "Saturday", "Sunday", "Monday"]:
            return "Weekend"
        return "Unknown"

    @staticmethod
    def is_attended_value(value):
        if pd.isna(value):
            return False
        return str(value).strip().lower() == "yes"

    def load_file(self, file_path):
        df = pd.read_excel(file_path, engine="openpyxl")

        if df.shape[0] < 1:
            raise ValueError(f"{file_path.name} does not contain any data rows.")
        if df.shape[1] < 2:
            raise ValueError(
                f"{file_path.name} must contain the game in column 1 and date/time in column 2."
            )

        report_game = df.iloc[0, 0]
        report_game_datetime = df.iloc[0, 1]
        report_game = "" if pd.isna(report_game) else str(report_game).strip()
        report_game_datetime = pd.to_datetime(report_game_datetime, errors="coerce")
        report_game_type = self.classify_game_type(report_game_datetime)

        required_columns = self.seat_key_columns + [self.compare_column]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(
                f"{file_path.name} is missing required columns: {', '.join(missing)}"
            )

        useful_columns = [
            "Area Name",
            "Block",
            "Seat Row",
            "Seat Number",
            self.compare_column,
            "Cust. Ref.",
            "Cust. Forename",
            "Cust. Surname",
            "Email",
            "Seat Ref.",
        ]
        existing_columns = [col for col in useful_columns if col in df.columns]
        df = df[existing_columns].copy()

        for col in self.seat_key_columns:
            df[col] = df[col].astype(str).str.strip()
        df[self.compare_column] = df[self.compare_column].astype(str).str.strip()

        for col in ["Cust. Ref.", "Cust. Forename", "Cust. Surname", "Email", "Seat Ref."]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()

        df["Game"] = report_game
        df["Game Date/Time"] = report_game_datetime
        df["Game Type"] = report_game_type

        duplicate_mask = df.duplicated(subset=self.seat_key_columns, keep=False)
        if duplicate_mask.any():
            print(
                f"Warning: {file_path.name} contains duplicate seat rows. Keeping the last occurrence for each seat."
            )
            df = df.drop_duplicates(subset=self.seat_key_columns, keep="last")

        return df

    def compare_two_files(self, old_file, new_file):
        old_df = self.load_file(old_file)
        new_df = self.load_file(new_file)

        old_name = old_file.stem
        new_name = new_file.stem

        merged = old_df.merge(
            new_df,
            on=self.seat_key_columns,
            how="outer",
            suffixes=(f"_{old_name}", f"_{new_name}"),
            indicator=True,
        )

        old_occ_col = f"{self.compare_column}_{old_name}"
        new_occ_col = f"{self.compare_column}_{new_name}"
        old_game_col = f"Game_{old_name}"
        new_game_col = f"Game_{new_name}"
        old_game_dt_col = f"Game Date/Time_{old_name}"
        new_game_dt_col = f"Game Date/Time_{new_name}"
        old_game_type_col = f"Game Type_{old_name}"
        new_game_type_col = f"Game Type_{new_name}"

        def classify_change(row):
            if row["_merge"] == "left_only":
                return "Missing in new report"
            if row["_merge"] == "right_only":
                return "New in new report"
            old_val = str(row.get(old_occ_col, "")).strip()
            new_val = str(row.get(new_occ_col, "")).strip()
            if old_val != new_val:
                return "Seat Occupied Changed"
            return "No Change"

        def get_attended_game(row):
            old_attended = self.is_attended_value(row.get(old_occ_col))
            new_attended = self.is_attended_value(row.get(new_occ_col))
            if old_attended and not new_attended:
                return row.get(old_game_col, "")
            if new_attended and not old_attended:
                return row.get(new_game_col, "")
            return ""

        def get_attended_game_datetime(row):
            old_attended = self.is_attended_value(row.get(old_occ_col))
            new_attended = self.is_attended_value(row.get(new_occ_col))
            if old_attended and not new_attended:
                return row.get(old_game_dt_col, "")
            if new_attended and not old_attended:
                return row.get(new_game_dt_col, "")
            return ""

        def get_attended_game_type(row):
            old_attended = self.is_attended_value(row.get(old_occ_col))
            new_attended = self.is_attended_value(row.get(new_occ_col))
            if old_attended and not new_attended:
                return row.get(old_game_type_col, "")
            if new_attended and not old_attended:
                return row.get(new_game_type_col, "")
            return ""

        def get_not_attended_game(row):
            old_attended = self.is_attended_value(row.get(old_occ_col))
            new_attended = self.is_attended_value(row.get(new_occ_col))
            if old_attended and not new_attended:
                return row.get(new_game_col, "")
            if new_attended and not old_attended:
                return row.get(old_game_col, "")
            return ""

        def get_not_attended_game_datetime(row):
            old_attended = self.is_attended_value(row.get(old_occ_col))
            new_attended = self.is_attended_value(row.get(new_occ_col))
            if old_attended and not new_attended:
                return row.get(new_game_dt_col, "")
            if new_attended and not old_attended:
                return row.get(old_game_dt_col, "")
            return ""

        def get_not_attended_game_type(row):
            old_attended = self.is_attended_value(row.get(old_occ_col))
            new_attended = self.is_attended_value(row.get(new_occ_col))
            if old_attended and not new_attended:
                return row.get(new_game_type_col, "")
            if new_attended and not old_attended:
                return row.get(old_game_type_col, "")
            return ""

        merged["Change Type"] = merged.apply(classify_change, axis=1)
        merged["Attended Game"] = merged.apply(get_attended_game, axis=1)
        merged["Attended Game Date/Time"] = merged.apply(get_attended_game_datetime, axis=1)
        merged["Attended Game Type"] = merged.apply(get_attended_game_type, axis=1)
        merged["Not Attended Game"] = merged.apply(get_not_attended_game, axis=1)
        merged["Not Attended Game Date/Time"] = merged.apply(get_not_attended_game_datetime, axis=1)
        merged["Not Attended Game Type"] = merged.apply(get_not_attended_game_type, axis=1)

        changes = merged[merged["Change Type"] != "No Change"].copy()

        ordered_columns = [
            "Area Name",
            "Block",
            "Seat Row",
            "Seat Number",
            "Change Type",
            "Attended Game",
            "Attended Game Date/Time",
            "Attended Game Type",
            "Not Attended Game",
            "Not Attended Game Date/Time",
            "Not Attended Game Type",
            old_game_col,
            old_game_dt_col,
            old_game_type_col,
            new_game_col,
            new_game_dt_col,
            new_game_type_col,
            old_occ_col,
            new_occ_col,
        ]

        extra_columns = [
            f"Cust. Ref._{old_name}",
            f"Cust. Ref._{new_name}",
            f"Cust. Forename_{old_name}",
            f"Cust. Forename_{new_name}",
            f"Cust. Surname_{old_name}",
            f"Cust. Surname_{new_name}",
            f"Email_{old_name}",
            f"Email_{new_name}",
            f"Seat Ref._{old_name}",
            f"Seat Ref._{new_name}",
        ]

        final_columns = [col for col in ordered_columns + extra_columns if col in changes.columns]
        return changes[final_columns]

    def compare_all(self):
        files = self.get_input_files()
        results = []
        for i in range(len(files) - 1):
            old_file = files[i]
            new_file = files[i + 1]
            changes = self.compare_two_files(old_file, new_file)
            changes.insert(0, "Compared From File", old_file.name)
            changes.insert(1, "Compared To File", new_file.name)
            results.append(changes)
        if results:
            return pd.concat(results, ignore_index=True)
        return pd.DataFrame()

    def save_split_reports(self, report):
        self.output_folder.mkdir(parents=True, exist_ok=True)
        midweek_report = report[
            report["Not Attended Game Type"].astype(str).str.strip().eq("Midweek")
        ].copy()
        weekend_report = report[
            report["Not Attended Game Type"].astype(str).str.strip().eq("Weekend")
        ].copy()

        midweek_path = self.output_folder / "missed_midweek.xlsx"
        weekend_path = self.output_folder / "missed_weekend.xlsx"
        midweek_report.to_excel(midweek_path, index=False)
        weekend_report.to_excel(weekend_path, index=False)
        print(f"Midweek report created: {midweek_path}")
        print(f"Weekend report created: {weekend_path}")

    def run(self):
        report = self.compare_all()
        if report.empty:
            print("No Seat Occupied changes found.")
        else:
            self.save_split_reports(report)
            print(report.head(20))


class RecentSeatOccupiedComparer(SeatOccupiedComparer):
    def get_report_metadata(self, file_path):
        df = pd.read_excel(file_path, engine="openpyxl")
        if df.shape[0] < 1:
            raise ValueError(f"{file_path.name} does not contain any data rows.")
        if df.shape[1] < 2:
            raise ValueError(
                f"{file_path.name} must contain the game in column 1 and date/time in column 2."
            )
        report_game = df.iloc[0, 0]
        report_game_datetime = df.iloc[0, 1]
        report_game = "" if pd.isna(report_game) else str(report_game).strip()
        report_game_datetime = pd.to_datetime(report_game_datetime, errors="coerce")
        if pd.isna(report_game_datetime):
            raise ValueError(
                f"{file_path.name} has an invalid game date/time in column 2 row 2."
            )
        return {
            "file_path": file_path,
            "game": report_game,
            "game_datetime": report_game_datetime,
            "game_type": self.classify_game_type(report_game_datetime),
        }

    def get_two_most_recent_files(self):
        files = sorted(self.import_folder.glob("*.xlsx"))
        if len(files) < 2:
            raise ValueError(
                f"At least 2 Excel files are required in '{self.import_folder}'."
            )
        file_metadata = [self.get_report_metadata(file) for file in files]
        file_metadata.sort(key=lambda x: x["game_datetime"], reverse=True)
        two_most_recent = file_metadata[:2]
        two_most_recent.sort(key=lambda x: x["game_datetime"])
        return two_most_recent[0]["file_path"], two_most_recent[1]["file_path"]

    def save_report(self, report):
        self.output_folder.mkdir(parents=True, exist_ok=True)
        output_path = self.output_folder / "missed_recent_game.xlsx"
        report.to_excel(output_path, index=False)
        print(f"Recent game report created: {output_path}")

    def run(self):
        old_file, new_file = self.get_two_most_recent_files()
        print("Comparing the two most recent games:")
        print(f"Older game file: {old_file.name}")
        print(f"Newer game file: {new_file.name}")
        report = self.compare_two_files(old_file, new_file)
        report.insert(0, "Compared From File", old_file.name)
        report.insert(1, "Compared To File", new_file.name)
        if report.empty:
            print("No Seat Occupied changes found between the two most recent games.")
        else:
            self.save_report(report)
            print(report.head(20))


class CityPointFilter:
    def __init__(
        self,
        import_folder="City_Points",
        output_folder="Reports_Output",
        seat_occupied_column="Seat Occupied",
        customer_ref_column="Cust. Ref.",
        output_file_prefix="City_Point_Add",
        max_rows_per_file=5000,
    ):
        self.import_folder = Path(import_folder)
        self.output_folder = Path(output_folder)
        self.seat_occupied_column = seat_occupied_column
        self.customer_ref_column = customer_ref_column
        self.output_file_prefix = output_file_prefix
        self.max_rows_per_file = max_rows_per_file

    def get_input_file(self):
        files = sorted(self.import_folder.glob("*.xlsx"))
        if not files:
            raise ValueError(f"No Excel files were found in '{self.import_folder}'.")
        if len(files) > 1:
            raise ValueError(
                f"Expected exactly 1 Excel file in '{self.import_folder}', but found {len(files)}."
            )
        return files[0]

    def load_file(self, file_path):
        df = pd.read_excel(file_path, engine="openpyxl")
        required_columns = [self.customer_ref_column, self.seat_occupied_column]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(
                f"{file_path.name} is missing required columns: {', '.join(missing)}"
            )

        df = df[[self.customer_ref_column, self.seat_occupied_column]].copy()
        df[self.customer_ref_column] = (
            df[self.customer_ref_column]
            .fillna("")
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
        )
        df[self.seat_occupied_column] = (
            df[self.seat_occupied_column]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )
        return df

    def filter_attended_customers(self, df):
        attended_df = df[df[self.seat_occupied_column] == "yes"].copy()
        attended_df = attended_df[attended_df[self.customer_ref_column] != ""].copy()
        attended_df = attended_df[[self.customer_ref_column]].drop_duplicates()
        return attended_df.reset_index(drop=True)

    def save_split_csvs(self, df):
        self.output_folder.mkdir(parents=True, exist_ok=True)
        total_rows = len(df)
        if total_rows == 0:
            output_path = self.output_folder / f"{self.output_file_prefix}_1.csv"
            df.to_csv(output_path, index=False, header=False, encoding="utf-8")
            print(f"CSV created: {output_path}")
            return

        total_files = math.ceil(total_rows / self.max_rows_per_file)
        for file_number in range(total_files):
            start_row = file_number * self.max_rows_per_file
            end_row = start_row + self.max_rows_per_file
            chunk_df = df.iloc[start_row:end_row].copy()
            output_path = self.output_folder / f"{self.output_file_prefix}_{file_number + 1}.csv"
            chunk_df.to_csv(output_path, index=False, header=False, encoding="utf-8")
            print(f"CSV created: {output_path} ({len(chunk_df)} customers)")

    def run(self):
        input_file = self.get_input_file()
        print(f"Using input file: {input_file.name}")
        df = self.load_file(input_file)
        output_df = self.filter_attended_customers(df)
        print(f"Unique attending customers found: {len(output_df)}")
        self.save_split_csvs(output_df)
        print("\nFirst 20 rows of export:")
        print(output_df.head(20))


class CustomerReportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("920x680")
        self.minsize(860, 580)

        if getattr(sys, "frozen", False):
            self.repo_root = Path(sys.executable).resolve().parent
        else:
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

        title = ttk.Label(header, text=APP_TITLE, font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, sticky="w")

        subtitle = ttk.Label(
            header,
            text=(
                "Standalone version. This app contains the compare-all, compare-recent, "
                "and City Points logic directly, so it can be packaged into one .exe."
            ),
            wraplength=780,
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(6, 0))

        top = ttk.Frame(self, padding=(16, 8, 16, 8))
        top.grid(row=1, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        top.columnconfigure(1, weight=1)

        actions = ttk.LabelFrame(top, text="Actions", padding=12)
        actions.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        actions.columnconfigure(0, weight=1)

        self.btn_compare_all = ttk.Button(actions, text="Run Compare All Games", command=self._run_compare_all)
        self.btn_compare_all.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.btn_compare_recent = ttk.Button(actions, text="Run Compare Two Most Recent Games", command=self._run_compare_recent)
        self.btn_compare_recent.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        self.btn_city_points = ttk.Button(actions, text="Run City Points Export", command=self._run_city_points)
        self.btn_city_points.grid(row=2, column=0, sticky="ew")

        status_box = ttk.LabelFrame(top, text="Project Status", padding=12)
        status_box.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        status_box.columnconfigure(1, weight=1)

        self.status_labels = {}
        for idx, folder in enumerate(FOLDERS):
            ttk.Label(status_box, text=f"{folder}:").grid(row=idx, column=0, sticky="w", padx=(0, 8), pady=2)
            lbl = ttk.Label(status_box, text="Checking...")
            lbl.grid(row=idx, column=1, sticky="w", pady=2)
            self.status_labels[folder] = lbl

        ttk.Label(status_box, text="Built-in tools:").grid(row=len(FOLDERS), column=0, sticky="nw", padx=(0, 8), pady=(8, 2))
        self.tools_status = ttk.Label(status_box, text="Checking...", justify="left")
        self.tools_status.grid(row=len(FOLDERS), column=1, sticky="w", pady=(8, 2))

        controls = ttk.Frame(self, padding=(16, 0, 16, 8))
        controls.grid(row=2, column=0, sticky="nsew")
        controls.columnconfigure(0, weight=1)
        controls.rowconfigure(1, weight=1)

        info_bar = ttk.Frame(controls)
        info_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        info_bar.columnconfigure(0, weight=1)

        self.current_action_var = tk.StringVar(value="Ready")
        ttk.Label(info_bar, textvariable=self.current_action_var).grid(row=0, column=0, sticky="w")
        self.refresh_button = ttk.Button(info_bar, text="Refresh Status", command=self._refresh_status)
        self.refresh_button.grid(row=0, column=1, sticky="e")

        log_frame = ttk.LabelFrame(controls, text="Output Log", padding=10)
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="word", height=18, font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        footer = ttk.Frame(self, padding=(16, 0, 16, 16))
        footer.grid(row=3, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Label(
            footer,
            text="Tip: put your Excel imports into Reports_Import or City_Points, then run the matching action above.",
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
            file_count = len(list(path.iterdir())) if path.exists() and path.is_dir() else 0
            self.status_labels[folder].configure(text=f"Found ({file_count} item(s))")
        self.tools_status.configure(text="\n".join(f"{tool}: Built-in" for tool in TOOLS))

    def _run_compare_all(self):
        self._run_worker("Compare All Games", self._do_compare_all)

    def _run_compare_recent(self):
        self._run_worker("Compare Two Most Recent Games", self._do_compare_recent)

    def _run_city_points(self):
        self._run_worker("City Points Export", self._do_city_points)

    def _run_worker(self, action_name, func):
        if self.is_running:
            return
        thread = threading.Thread(target=self._worker, args=(action_name, func), daemon=True)
        thread.start()

    def _worker(self, action_name, func):
        self.after(0, lambda: self._set_running_state(True))
        self.after(0, lambda: self.current_action_var.set(f"Running: {action_name}"))
        self.after(0, lambda: self._append_log(f"\n{'=' * 70}\nRunning {action_name}\n{'=' * 70}\n"))
        stream = StringIO()
        try:
            with redirect_stdout(stream):
                func()
            output = stream.getvalue()
            if output:
                self.after(0, lambda output=output: self._append_log(output))
            self.after(0, lambda: self._append_log(f"\nFinished successfully: {action_name}\n"))
            self.after(0, lambda: self.current_action_var.set(f"Finished: {action_name}"))
        except Exception:
            output = stream.getvalue()
            err = traceback.format_exc()
            combined = (output + "\n" + err).strip() + "\n"
            self.after(0, lambda combined=combined: self._append_log(combined))
            self.after(0, lambda: self.current_action_var.set(f"Failed: {action_name}"))
            self.after(0, lambda: messagebox.showerror("Action Failed", f"{action_name} failed. Check the output log for details."))
        finally:
            self.after(0, self._refresh_status)
            self.after(0, lambda: self._set_running_state(False))

    def _do_compare_all(self):
        comparer = SeatOccupiedComparer(
            import_folder=self.repo_root / "Reports_Import",
            output_folder=self.repo_root / "Reports_Output",
        )
        comparer.run()

    def _do_compare_recent(self):
        comparer = RecentSeatOccupiedComparer(
            import_folder=self.repo_root / "Reports_Import",
            output_folder=self.repo_root / "Reports_Output",
        )
        comparer.run()

    def _do_city_points(self):
        filter_tool = CityPointFilter(
            import_folder=self.repo_root / "City_Points",
            output_folder=self.repo_root / "Reports_Output",
        )
        filter_tool.run()


if __name__ == "__main__":
    app = CustomerReportApp()
    app.mainloop()
