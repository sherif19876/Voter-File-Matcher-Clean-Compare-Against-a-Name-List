import pandas as pd
import re
import os


def clean_name(name):
    """
    Clean a name string by removing spaces, non-letter characters, and lowercasing.

    Parameters:
    name (any): Input name, could be NaN or string.

    Returns:
    str: Cleaned name containing only lowercase letters, or empty string if input is NaN.
    """
    if pd.isna(name):
        return ""
    # Convert to string, lowercase, then remove any character that is not a-z
    return re.sub(r'[^a-z]', '', str(name).lower())


def process_voter_file(fla_file, names_df, output_prefix):
    """
    Process a single voter Excel file: clean names, match against a reference list,
    and export matched rows and unmatched names.

    Parameters:
    fla_file (str): Path to the voter Excel file.
    names_df (DataFrame): DataFrame with cleaned reference names (column 'Full_Clean').
    output_prefix (str): Prefix for output Excel files.
    """
    print(f"\n========== PROCESSING {fla_file} ==========\n")

    try:
        # Read the Excel file without headers, all as strings
        df = pd.read_excel(fla_file, dtype=str, header=None)
        print(f"Loaded {len(df)} rows from '{fla_file}'")
    except Exception as e:
        print(f"Error reading {fla_file}: {e}")
        return

    # Define column indices (0-based) where last and first names are located
    # These are consistent across all input files
    last_col_idx = 2   # 3rd column = Last name
    first_col_idx = 4  # 5th column = First name

    print(f"\nUsing column {last_col_idx} (3rd) for LAST name")
    print(f"Using column {first_col_idx} (5th) for FIRST name")

    print("\nSample pulled values from voter file:")
    print(pd.DataFrame({
        "Raw Last": df.iloc[:5, last_col_idx],
        "Raw First": df.iloc[:5, first_col_idx]
    }))

    # Clean the extracted first and last names using the clean_name function
    df['First_Clean'] = df.iloc[:, first_col_idx].apply(clean_name)
    df['Last_Clean'] = df.iloc[:, last_col_idx].apply(clean_name)
    # Create a combined key: first_last (e.g., "john_doe")
    df['Full_Clean'] = df['First_Clean'] + "_" + df['Last_Clean']

    print("Cleaned FLA-style data for matching.")
    print("Example cleaned FLA names:", df['Full_Clean'].head(5).tolist())

    # Match: keep rows where the cleaned full name exists in the reference names set
    matched = df[df['Full_Clean'].isin(set(names_df['Full_Clean']))].drop(
        columns=['First_Clean', 'Last_Clean', 'Full_Clean']
    )
    print(f"Found {len(matched)} matching rows in {fla_file}")

    # Determine which reference names were NOT found in this voter file
    found_set = set(df[df['Full_Clean'].isin(set(names_df['Full_Clean']))]['Full_Clean'])
    not_found = names_df[~names_df['Full_Clean'].isin(found_set)]
    print(f"{len(not_found)} names not found in {fla_file}")

    # Prepare output filenames based on the input file's base name
    base_name = os.path.splitext(os.path.basename(fla_file))[0]
    matched_file = f"{output_prefix}_{base_name}.xlsx"
    notfound_file = f"{output_prefix}_{base_name}_NotFound.xlsx"

    # Export matched rows (without headers as per original requirement)
    try:
        matched.to_excel(matched_file, index=False, header=False)
        print(f"Matches saved to '{matched_file}' ({len(matched)} rows)")
    except Exception as e:
        print(f"Error writing {matched_file}: {e}")

    # Export the list of names not found (only First and Last columns)
    try:
        not_found[['First', 'Last']].to_excel(notfound_file, index=False)
        print(f"Not-found names saved to '{notfound_file}'")
    except Exception as e:
        print(f"Error writing {notfound_file}: {e}")

    print(f"\n========== COMPLETED {fla_file} ==========\n")


def match_all_voter_files(names_file):
    """
    Main orchestrator: load the reference names CSV, clean them, and process
    each voter Excel file defined in a hardcoded list.

    Parameters:
    names_file (str): Path to CSV file containing at least 'First' and 'Last' columns.
    """
    # Load the reference names CSV
    try:
        names_df = pd.read_csv(names_file, dtype=str)
        print(f"Loaded names file '{names_file}' with {len(names_df)} entries")
    except Exception as e:
        print(f"Error reading names file: {e}")
        return

    # Validate required columns
    if not all(col in names_df.columns for col in ['First', 'Last']):
        print("ERROR: The names CSV must contain columns 'First' and 'Last'.")
        return

    # Clean the reference names (same cleaning logic as voter data)
    names_df['First_Clean'] = names_df['First'].apply(clean_name)
    names_df['Last_Clean'] = names_df['Last'].apply(clean_name)
    names_df['Full_Clean'] = names_df['First_Clean'] + "_" + names_df['Last_Clean']
    print("Cleaned all names for matching.\n")

    # List of voter Excel files to process (adjust paths as needed)
    voter_files = [
        "FLA_20250909.xlsx",
        "STJ_20250909.xlsx",
        "VOL_20250909.xlsx",
        "LAK_20250909.xlsx",
        "PUT_20250909.xlsx"
    ]

    # Process each file
    for file in voter_files:
        process_voter_file(file, names_df, output_prefix="Namesfin")

    print("\nAll files processed successfully.")


# Run the full workflow when script is executed directly
if __name__ == "__main__":
    match_all_voter_files("Names_Split_into_First_and_Last.csv")
