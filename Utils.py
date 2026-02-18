import pandas as pd
import os

def load_dynamic_df(path, sheet, target_col, max_search=10):
    """Searches for the header row within the first max_search rows."""
    for i in range(max_search + 1):
        try:
            df = pd.read_excel(path, sheet_name=sheet, header=i)
            df.columns = df.columns.str.strip()
            if target_col in df.columns:
                return df
        except Exception:
            continue
    raise KeyError(f"Could not find header with column '{target_col}' in the first {max_search} rows of {path}")

def get_unique_filename(file_path):
    """Checks if a file exists and appends a numeric suffix if it does."""
    if not os.path.exists(file_path):
        return file_path

    # Split into file path/name and the .xlsx extension
    base, extension = os.path.splitext(file_path)
    counter = 1
    
    # Try 'FileName 1.xlsx', 'FileName 2.xlsx', etc.
    new_path = f"{base} {counter}{extension}"
    while os.path.exists(new_path):
        counter += 1
        new_path = f"{base} {counter}{extension}"
        
    return new_path

def col_to_idx(col_letter):
            # Converts 'A' -> 0, 'B' -> 1, etc.
            num = 0
            for c in col_letter:
                num = num * 26 + (ord(c.upper()) - ord('A') + 1)
            return num - 1

def input_with_default(prompt, default):
    """Returns the default value if the user just presses Enter."""
    user_input = input(f"{prompt} [{default}]: ").strip()
    return user_input if user_input else default