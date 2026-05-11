import os
import re
import pandas as pd
import numpy as np

FILENAME_PATTERN = re.compile(r"^GP9_2023(\d{2})\.csv$")


def read_excel_file():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, "Regional FES.xlsx")
    
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    
    try:
        import openpyxl
    except ImportError as e:
        raise ImportError("openpyxl is required to read .xlsm files. Install it with `pip install openpyxl`." ) from e

    try:
        sheet_configs = {
            "GSP info": {"usecols": "B:G", "skiprows": 4, "nrows": 376},
            "MAIN DATA": {"usecols": "A:N", "skiprows": 14, "nrows": 539},
            "DG": {"usecols": "A:H", "skiprows": 2, "nrows": 71158},
            "Sub1MW": {"usecols": "A:H", "skiprows": 2, "nrows": 152181}
        }
        all_sheets_data = {}

        with pd.ExcelFile(excel_path, engine='openpyxl') as xls:
            for sheet_name, config in sheet_configs.items():
                df = pd.read_excel(xls, sheet_name=sheet_name, **config)
                
                all_sheets_data[sheet_name] = df
                print(f"Successfully read Excel file sheet '{sheet_name}' with columns: {list(df.columns)}")
        
        return all_sheets_data
    except Exception as e:
        raise RuntimeError(f"Error reading Excel file '{excel_path}': {e}") from e


def read_ordered_gp9_csv_files():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    files = os.listdir(current_dir)

    matched_files = []
    for filename in files:
        match = FILENAME_PATTERN.match(filename)
        if match:
            month = int(match.group(1))
            matched_files.append((month, filename))

    matched_files.sort(key=lambda item: item[0])

    dataframes = []
    for month, filename in matched_files:
        file_path = os.path.join(current_dir, filename)
        try:
            df = pd.read_csv(file_path)
            dataframes.append(df)
            print(f"Successfully read: {filename}")
        except Exception as e:
            print(f"Error reading {filename}: {e}")

    if not dataframes:
        raise FileNotFoundError("No GP9_2023MM.csv files were found in the script directory.")

    combined = pd.concat(dataframes, ignore_index=True)
    print(f"Combined {len(dataframes)} files into one DataFrame with {len(combined)} rows.")
    return combined


def prompt_select_gsp_id(df: pd.DataFrame) -> pd.DataFrame:
    if "GSP Id" not in df.columns:
        raise KeyError("The dataframe does not contain a 'GSP Id' column.")

    unique_values = sorted(df["GSP Id"].dropna().unique())
    if not unique_values:
        raise ValueError("No unique GSP Id values were found in the dataframe.")

    print("\nAvailable GSP Id values:")
    for index, value in enumerate(unique_values, start=1):
        print(f"{index}. {value}")

    while True:
        selection = input("Enter the number of the GSP Id to filter by: ").strip()
        if not selection.isdigit():
            print("Please enter a valid number.")
            continue

        selection_index = int(selection)
        if 1 <= selection_index <= len(unique_values):
            selected_value = unique_values[selection_index - 1]
            filtered_df = df[df["GSP Id"] == selected_value].reset_index(drop=True)
            print(f"Filtered dataframe to GSP Id '{selected_value}' with {len(filtered_df)} rows.")
            return filtered_df

        print(f"Please enter a number between 1 and {len(unique_values)}.")


def create_datetime_index(df: pd.DataFrame) -> pd.DataFrame:
    if "Settlement Date" not in df.columns or "Settlement Period" not in df.columns:
        raise KeyError("The dataframe must contain 'Settlement Date' and 'Settlement Period' columns.")

    print(f"\nSettlement Period range: {df['Settlement Period'].min()} to {df['Settlement Period'].max()}")
    print(f"Total rows before processing: {len(df)}")
    
    # Filter out periods > 48 to handle DST (fall-back days have periods 49-50)
    df = df[df['Settlement Period'] <= 48].copy()
    print(f"Total rows after filtering (Settlement Period <= 48): {len(df)}")

    def convert_to_datetime(row):
        date_str = str(row["Settlement Date"]).zfill(8)
        period = int(row["Settlement Period"])
        
        year = int(date_str[:4])
        month = int(date_str[4:6])
        day = int(date_str[6:8])
        
        hour = (period - 1) // 2
        minute = ((period - 1) % 2) * 30
        
        return pd.Timestamp(year, month, day, hour, minute)

    df["Datetime"] = df.apply(convert_to_datetime, axis=1)
    
    # Keep only the first occurrence of each unique datetime (handles spring-forward day with 46 periods)
    df = df.drop_duplicates(subset=['Datetime'], keep='first')
    print(f"Total rows after deduplication (365*48 timesteps): {len(df)}")
    
    df = df.set_index("Datetime")
    return df


def prompt_select_year(df: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df.index, pd.DatetimeIndex):
        raise TypeError("The dataframe index must be a DatetimeIndex.")
    
    unique_years = sorted(df.index.year.unique())
    if not unique_years:
        raise ValueError("No year values found in the dataframe index.")
    
    print("\nAvailable years:")
    for index, year in enumerate(unique_years, start=1):
        print(f"{index}. {year}")
    
    while True:
        selection = input("Enter the number of the year to filter by: ").strip()
        if not selection.isdigit():
            print("Please enter a valid number.")
            continue
        
        selection_index = int(selection)
        if 1 <= selection_index <= len(unique_years):
            selected_year = unique_years[selection_index - 1]
            filtered_df = df[df.index.year == selected_year]
            print(f"Filtered dataframe to year {selected_year} with {len(filtered_df)} rows.")
            return filtered_df
        print(f"Please enter a number between 1 and {len(unique_years)}.")


def fill_missing_periods(df: pd.DataFrame, year: int) -> pd.DataFrame:
    if not isinstance(df.index, pd.DatetimeIndex):
        raise TypeError("The dataframe index must be a DatetimeIndex.")
    
    # Generate complete DatetimeIndex for the year (365 days * 48 periods)
    start_date = pd.Timestamp(year, 1, 1)
    end_date = pd.Timestamp(year, 12, 31, 23, 30)
    complete_index = pd.date_range(start=start_date, end=end_date, freq='30min')
    
    print(f"Expected periods for {year}: {len(complete_index)}")
    print(f"Actual periods in data: {len(df)}")
    
    missing_count = len(complete_index) - len(df)
    if missing_count > 0:
        print(f"Missing {missing_count} periods. Interpolating...")
        
        # Reindex to complete index and interpolate
        df = df.reindex(complete_index)
        
        # Interpolate numeric columns
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        df[numeric_cols] = df[numeric_cols].interpolate(method='linear')
        
        print(f"After interpolation: {len(df)} periods")
    else:
        print("No missing periods found.")
    
    return df


if __name__ == "__main__":
    read_csv_files = True


    # Read and save Excel file separately
    excel_data = read_excel_file()
    
    # Process CSV files
    if read_csv_files == True:
        all_data = read_ordered_gp9_csv_files()
        filtered_data = prompt_select_gsp_id(all_data)
        filtered_data = create_datetime_index(filtered_data)
        filtered_data = prompt_select_year(filtered_data)
        selected_year: int = filtered_data.index.year[0] # type: ignore
        filtered_data = fill_missing_periods(filtered_data, selected_year)
        filtered_data.to_csv("filtered_gp9_data.csv")
        print(filtered_data)
        print(len(filtered_data))
