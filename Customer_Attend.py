import pandas as pd
from pathlib import Path


class SeatOccupiedComparer:
    def __init__(
        self,
        import_folder="Reports_Import",
        output_folder="Reports_Output",
        compare_column="Seat Occupied"
    ):
        self.import_folder = Path(import_folder)
        self.output_folder = Path(output_folder)
        self.compare_column = compare_column

        self.seat_key_columns = [
            "Area Name",
            "Block",
            "Seat Row",
            "Seat Number",
        ]

    def get_input_files(self):
        files = sorted(self.import_folder.glob("*.xlsx"))

        if len(files) < 2:
            raise ValueError(
                f"At least 2 Excel files are required in '{self.import_folder}'."
            )

        return files

    def classify_game_type(self, date_value):
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

    def is_attended_value(self, value):
        if pd.isna(value):
            return False

        value = str(value).strip().lower()
        return value == "yes"

    def load_file(self, file_path):
        df = pd.read_excel(file_path, engine="openpyxl")

        if df.shape[0] < 1:
            raise ValueError(f"{file_path.name} does not contain any data rows.")

        if df.shape[1] < 2:
            raise ValueError(
                f"{file_path.name} must contain the game in column 1 and date/time in column 2."
            )

        # Excel row 2 = pandas index 0
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

        text_columns = [
            "Cust. Ref.",
            "Cust. Forename",
            "Cust. Surname",
            "Email",
            "Seat Ref.",
        ]

        for col in text_columns:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()

        df["Game"] = report_game
        df["Game Date/Time"] = report_game_datetime
        df["Game Type"] = report_game_type

        duplicate_mask = df.duplicated(subset=self.seat_key_columns, keep=False)
        if duplicate_mask.any():
            print(
                f"Warning: {file_path.name} contains duplicate seat rows. "
                f"Keeping the last occurrence for each seat."
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
            indicator=True
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
        changes = changes[final_columns]

        return changes

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


if __name__ == "__main__":
    comparer = SeatOccupiedComparer()
    report = comparer.compare_all()

    if report.empty:
        print("No Seat Occupied changes found.")
    else:
        comparer.save_split_reports(report)
        print(report.head(20))