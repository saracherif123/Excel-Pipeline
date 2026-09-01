import os
import pandas as pd
import argparse

DEFAULT_FOLDER1 = "/Users/sarasaad/Documents/Data Processing/Data/Q & FS/processed"
DEFAULT_FOLDER2 = "/Users/sarasaad/Documents/Data Processing/Data/Q & FS"
DEFAULT_COMBINED_FILE = "/Users/sarasaad/Documents/Data Processing/Q&FS_cleaned_final.xlsx"


def count_rows_in_folder(folder_path):
    total_rows = 0

    for file in os.listdir(folder_path):

        if file.startswith("~$"):  # skip temp Excel files
            continue

        if file.endswith(".xlsx") or file.endswith(".xls"):

            file_path = os.path.join(folder_path, file)

            try:
                df = pd.read_excel(file_path, engine="openpyxl")
                row_count = len(df)

                print(f"{file} -> {row_count} rows")

                total_rows += row_count

            except Exception as e:
                print(f"Error reading {file}: {e}")

    return total_rows


def compare_row_counts(folder1: str, folder2: str, combined_file: str) -> dict:
    print("\nProcessing Folder 1 (processed files)")
    total1 = count_rows_in_folder(folder1)

    print("\nProcessing Folder 2 (original files)")
    total2 = count_rows_in_folder(folder2)

    print("\nProcessing Combined File")
    df_combined = pd.read_excel(combined_file)
    total3 = len(df_combined)
    print(f"{combined_file} -> {total3} rows")

    print("\n------ FINAL COMPARISON ------")
    print(f"Folder 1 total rows (processed): {total1}")
    print(f"Folder 2 total rows (original): {total2}")
    print(f"Combined Excel rows: {total3}")

    print("\nDifferences:")
    print(f"Processed vs Original: {total1 - total2}")
    print(f"Combined vs Processed: {total3 - total1}")
    print(f"Combined vs Original: {total3 - total2}")

    return {
        "processed_rows": total1,
        "original_rows": total2,
        "combined_rows": total3,
        "processed_minus_original": total1 - total2,
        "combined_minus_processed": total3 - total1,
        "combined_minus_original": total3 - total2,
    }


def main():
    parser = argparse.ArgumentParser(description="Compare row counts between folders and a consolidated file.")
    parser.add_argument("--processed-folder", default=DEFAULT_FOLDER1, help="Folder containing processed Excel files")
    parser.add_argument("--original-folder", default=DEFAULT_FOLDER2, help="Folder containing original Excel files")
    parser.add_argument("--combined-file", default=DEFAULT_COMBINED_FILE, help="Combined Excel file path")
    args = parser.parse_args()

    compare_row_counts(args.processed_folder, args.original_folder, args.combined_file)


if __name__ == "__main__":
    main()