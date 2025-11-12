import pandas as pd
import io
import numpy as np
import json
import os
import sys

# Define the key columns for comparison (unique row identifier)
KEY_COLS = ['HQ_CODE', 'ID', 'MBR_NO']
TAB_SEP = "\t"


def load_config(config_path="config.json"):
    """Loads file paths from the configuration JSON file."""
    if not os.path.exists(config_path):
        print(f"Error: Config file not found at {config_path}")
        sys.exit(1)
    with open(config_path, "r") as f:
        return json.load(f)


def load_data_file(file_path, skip_rows=0):
    """Loads a tab-separated file into a DataFrame, ensuring string type."""
    if not os.path.exists(file_path):
        print(f"Warning: File not found at {file_path}. Skipping comparison.")
        return None

    # Use io.StringIO to handle the file content if it was loaded into memory,
    # but for reading from disk, pd.read_csv works directly.
    try:
        # read_csv with tab delimiter, skipping header/metadata rows
        df = pd.read_csv(file_path, sep=TAB_SEP, skiprows=skip_rows, dtype=str)
        # Clean column names (strip whitespace)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return None


def run_comparison_test(test_case):
    """
    Implements the row count check and proceeds to comparison if counts match.
    """
    name = test_case['name']
    before_path = test_case['before_path']
    after_path = test_case['after_path']

    print(f"\n--- Running Test: {name} ---")

    # The 'before' file has a metadata header row we need to skip (skiprows=1)
    df_before = load_data_file(before_path, skip_rows=1)
    # The 'after' file starts directly with the header (skiprows=0)
    df_after = load_data_file(after_path, skip_rows=0)

    if df_before is None or df_after is None:
        return

    # --- Test Case 1: Row Count Check ---
    if len(df_before) != len(df_after):
        print("❌ TEST FAILED: Row Count Mismatch!")
        print(
            f"The number of rows are mismatch in file {os.path.basename(before_path)} ({len(df_before)} rows) and {os.path.basename(after_path)} ({len(df_after)} rows)")
        return
    else:
        print(
            f"✅ Row Count Check Passed: Both files have {len(df_before)} data rows. Proceeding to cell-level comparison.")

    # --- Data Preparation for Comparison ---
    try:
        df_before_aligned = df_before.set_index(KEY_COLS).sort_index()
        df_after_aligned = df_after.set_index(KEY_COLS).sort_index()
    except KeyError as e:
        print(f"❌ Error: One or more key columns {KEY_COLS} not found in the files. Check headers.")
        return

    # --- Cell-Level Comparison ---
    df_diff = df_before_aligned.compare(df_after_aligned, align_axis=0)

    # --- Formatting Output (for differences) ---
    output_rows = []
    highlight_start_B = "**>>BEFORE:"
    highlight_end_B = "<<**"
    highlight_start_A = "**>>AFTER:"
    highlight_end_A = "<<**"

    for index, diff_row in df_diff.iterrows():
        output_parts = [f"KEY:{k}={v}" for k, v in zip(KEY_COLS, index)]

        for col_name in df_diff.columns.get_level_values(0).unique():
            val_before = diff_row.get((col_name, 'self'))
            val_after = diff_row.get((col_name, 'other'))

            if pd.notna(val_before) or pd.notna(val_after):
                field_output = f"{col_name}:"

                if pd.notna(val_before):
                    field_output += f"{highlight_start_B}{val_before}{highlight_end_B}"

                if pd.notna(val_after):
                    if pd.notna(val_before):
                        field_output += " | "
                    field_output += f"{highlight_start_A}{val_after}{highlight_end_A}"

                output_parts.append(field_output)

        output_rows.append(TAB_SEP.join(output_parts))

    # --- Final Result ---
    if output_rows:
        print("🚨 Cell-Level Differences Found:")
        print("Format: KEY_COLUMNS\tFIELD_NAME: **>>BEFORE:Value_B<<** | **>>AFTER:Value_A<<**")
        print("-" * 100)
        print("\n".join(output_rows))
    else:
        print("✅ Cell-Level Comparison Passed: No differences found in data fields.")


def main():
    """Main execution function to read config and run tests."""
    # Assuming config.json is in the same directory as the script
    config = load_config()

    if 'input_files' not in config or not config['input_files']:
        print("Error: 'input_files' list is missing or empty in config.json.")
        return

    for test_case in config['input_files']:
        if all(key in test_case for key in ['name', 'before_path', 'after_path']):
            run_comparison_test(test_case)
        else:
            print(f"Warning: Skipping malformed test case entry: {test_case}")


if __name__ == "__main__":
    main()