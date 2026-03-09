import math
import pandas as pd
from pathlib import Path


class CityPointFilter:
    def __init__(
        self,
        import_folder="City_Points",
        output_folder="Reports_Output",
        seat_occupied_column="Seat Occupied",
        customer_ref_column="Cust. Ref.",
        output_file_prefix="City_Point_Add",
        max_rows_per_file=5000
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
            raise ValueError(
                f"No Excel files were found in '{self.import_folder}'."
            )

        if len(files) > 1:
            raise ValueError(
                f"Expected exactly 1 Excel file in '{self.import_folder}', "
                f"but found {len(files)}."
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

        attended_df = attended_df[
            attended_df[self.customer_ref_column] != ""
        ].copy()

        attended_df = attended_df[[self.customer_ref_column]].drop_duplicates()

        attended_df = attended_df.reset_index(drop=True)

        return attended_df

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

            output_path = (
                self.output_folder /
                f"{self.output_file_prefix}_{file_number + 1}.csv"
            )

            chunk_df.to_csv(
                output_path,
                index=False,
                header=False,
                encoding="utf-8"
            )

            print(
                f"CSV created: {output_path} "
                f"({len(chunk_df)} customers)"
            )

    def run(self):
        input_file = self.get_input_file()
        print(f"Using input file: {input_file.name}")

        df = self.load_file(input_file)
        output_df = self.filter_attended_customers(df)

        print(f"Unique attending customers found: {len(output_df)}")

        self.save_split_csvs(output_df)

        print("\nFirst 20 rows of export:")
        print(output_df.head(20))


if __name__ == "__main__":
    filter_tool = CityPointFilter()
    filter_tool.run()