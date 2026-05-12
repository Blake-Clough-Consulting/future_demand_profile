import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt  # type: ignore
import matplotlib.dates as mdates  # type: ignore
from datetime import timedelta

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
            "Sub1MW": {"usecols": "A:H", "skiprows": 2, "nrows": 152181},
            "DxStorage": {"usecols": "A:H", "skiprows": 2, "nrows": 12786},
            "mBattery": {"usecols": "A:H", "skiprows": 2, "nrows": 30420},
            "LV Gain": {"usecols": "A:H", "skiprows": 2, "nrows": 310388},
            "Active": {"usecols": "A:H", "skiprows": 2, "nrows": 225828},
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


def prompt_select_gsp_id(df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    if "GSP Id" not in df.columns:
        raise KeyError("The dataframe does not contain a 'GSP Id' column.")

    unique_values = sorted(df["GSP Id"].dropna().unique())
    if not unique_values:
        raise ValueError("No unique GSP Id values were found in the dataframe.")

    print("\nAvailable GSP Id (Elexon ID) values:")
    for index, value in enumerate(unique_values, start=1):
        print(f"{index}. {value}")

    while True:
        selection = input("Enter the number of the GSP Id (Elexon ID) to filter by: ").strip()
        if not selection.isdigit():
            print("Please enter a valid number.")
            continue

        selection_index = int(selection)
        if 1 <= selection_index <= len(unique_values):
            selected_value = unique_values[selection_index - 1]
            filtered_df = df[df["GSP Id"] == selected_value].reset_index(drop=True)
            print(f"Filtered dataframe to GSP Id (Elexon ID) '{selected_value}' with {len(filtered_df)} rows.")
            return filtered_df, selected_value

        print(f"Please enter a number between 1 and {len(unique_values)}.")


def filter_data_by_elexon_id(excel_data: dict, elexon_id: str) -> dict:
    """
    Filter all excel dataframes by the selected Elexon ID.
    
    Filters each sheet using the appropriate Elexon ID column:
    - MAIN DATA: 'Elexon ID' column
    - DG: 'etys_location' column
    - Sub1MW: 'etys_location' column
    
    Args:
        excel_data: Dictionary of all excel sheets from read_excel_file()
        elexon_id: The selected Elexon ID (GSP Id) to filter by
    
    Returns:
        Dictionary containing filtered dataframes for each sheet:
            - 'MAIN DATA': Filtered MAIN DATA rows
            - 'DG': Filtered DG rows
            - 'Sub1MW': Filtered Sub1MW rows
            - 'GSP info': Unfiltered (reference data)
    """
    filtered_data = {}
    
    # Filter MAIN DATA by 'Elexon ID' column
    if "MAIN DATA" in excel_data:
        main_data_df = excel_data["MAIN DATA"]
        if "Elexon ID" in main_data_df.columns:
            filtered_main_data = main_data_df[main_data_df["Elexon ID"].astype(str) == str(elexon_id)].reset_index(drop=True)
            filtered_data["MAIN DATA"] = filtered_main_data
            print(f"Filtered MAIN DATA to Elexon ID '{elexon_id}': {len(filtered_main_data)} rows")
        else:
            print(f"Warning: 'Elexon ID' column not found in MAIN DATA. Available columns: {list(main_data_df.columns)}")
            filtered_data["MAIN DATA"] = pd.DataFrame()
    
    # Filter DG by 'etys_location' column
    if "DG" in excel_data:
        dg_df = excel_data["DG"]
        if "etys_location" in dg_df.columns:
            filtered_dg = dg_df[dg_df["etys_location"].astype(str) == str(elexon_id)].reset_index(drop=True)
            filtered_data["DG"] = filtered_dg
            print(f"Filtered DG to etys_location '{elexon_id}': {len(filtered_dg)} rows")
        else:
            print(f"Warning: 'etys_location' column not found in DG. Available columns: {list(dg_df.columns)}")
            filtered_data["DG"] = pd.DataFrame()
    
    # Filter Sub1MW by 'etys_location' column
    if "Sub1MW" in excel_data:
        sub1mw_df = excel_data["Sub1MW"]
        if "etys_location" in sub1mw_df.columns:
            filtered_sub1mw = sub1mw_df[sub1mw_df["etys_location"].astype(str) == str(elexon_id)].reset_index(drop=True)
            filtered_data["Sub1MW"] = filtered_sub1mw
            print(f"Filtered Sub1MW to etys_location '{elexon_id}': {len(filtered_sub1mw)} rows")
        else:
            print(f"Warning: 'etys_location' column not found in Sub1MW. Available columns: {list(sub1mw_df.columns)}")
            filtered_data["Sub1MW"] = pd.DataFrame()

    if "DxStorage" in excel_data:
        dxstorage_df = excel_data["DxStorage"]
        if "location" in dxstorage_df.columns:
            filtered_dxstorage = dxstorage_df[dxstorage_df["location"].astype(str) == str(elexon_id)].reset_index(drop=True)
            filtered_data["DxStorage"] = filtered_dxstorage
            print(f"Filtered DxStorage to location '{elexon_id}': {len(filtered_dxstorage)} rows")
        else:
            print(f"Warning: 'location' column not found in DxStorage. Available columns: {list(dxstorage_df.columns)}")
            filtered_data["DxStorage"] = pd.DataFrame()

    if "mBattery" in excel_data:
        mbattery_df = excel_data["mBattery"]
        if "etys_location" in mbattery_df.columns:
            filtered_mbattery = mbattery_df[mbattery_df["etys_location"].astype(str) == str(elexon_id)].reset_index(drop=True)
            filtered_data["mBattery"] = filtered_mbattery
            print(f"Filtered mBattery to etys_location '{elexon_id}': {len(filtered_mbattery)} rows")
        else:
            print(f"Warning: 'etys_location' column not found in mBattery. Available columns: {list(mbattery_df.columns)}")
            filtered_data["mBattery"] = pd.DataFrame()

    if "LV Gain" in excel_data:
        lv_gain_df = excel_data["LV Gain"]
        if "GSP" in lv_gain_df.columns:
            filtered_lv_gain = lv_gain_df[lv_gain_df["GSP"].astype(str) == str(elexon_id)].reset_index(drop=True)
            filtered_data["LV Gain"] = filtered_lv_gain
            print(f"Filtered LV Gain to GSP '{elexon_id}': {len(filtered_lv_gain)} rows")
        else:
            print(f"Warning: 'GSP' column not found in LV Gain. Available columns: {list(lv_gain_df.columns)}")
            filtered_data["LV Gain"] = pd.DataFrame()

    if "Active" in excel_data:
        active_df = excel_data["Active"]
        if "GSP" in active_df.columns:
            filtered_active = active_df[active_df["GSP"].astype(str) == str(elexon_id)].reset_index(drop=True)
            filtered_data["Active"] = filtered_active
            print(f"Filtered Active to GSP '{elexon_id}': {len(filtered_active)} rows")
        else:
            print(f"Warning: 'GSP' column not found in Active. Available columns: {list(active_df.columns)}")
            filtered_data["Active"] = pd.DataFrame()    

    # Keep GSP info unfiltered (reference data)
    if "GSP info" in excel_data:
        filtered_data["GSP info"] = excel_data["GSP info"]

    return filtered_data




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


def prompt_select_year(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
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
            return filtered_df, selected_year
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


def prepare_processing_data(excel_data: dict, csv_data: pd.DataFrame) -> dict:
    """
    Consolidate excel and csv data selection/filtering into a single processing package.
    
    This function orchestrates the user selection flow and returns all necessary data
    organized and ready for profile transformation/scaling operations.
    
    Args:
        excel_data: Dictionary of all excel sheets from read_excel_file()
        csv_data: Combined CSV dataframe from read_ordered_gp9_csv_files()
    
    Returns:
        Dictionary containing:
            - 'elexon_id': The selected Elexon ID (GSP Id)
            - 'year': The selected year
            - 'csv_profile': The processed CSV data with datetime index, filtered to year
            - 'main_data': The filtered MAIN DATA row(s) for the selected Elexon ID
            - 'gsp_info': The complete GSP info sheet (for reference/lookup)
            - 'dg_data': The complete DG sheet (for reference)
            - 'sub1mw_data': The complete Sub1MW sheet (for reference)
    """
    print("\n" + "="*80)
    print("PROCESSING DATA SELECTION AND PREPARATION")
    print("="*80)
    
    # Step 1: Select Elexon ID (GSP ID)
    print("\nStep 1: Select GSP ID (Elexon ID)")
    csv_filtered, selected_elexon_id = prompt_select_gsp_id(csv_data)
    
    # Step 2: Filter all excel data by Elexon ID
    print("\nStep 2: Filtering all excel data by selected Elexon ID")
    filtered_excel_data = filter_data_by_elexon_id(excel_data, selected_elexon_id)
    
    # Step 3: Create datetime index and filter by year
    print("\nStep 3: Processing CSV data (datetime index and year selection)")
    csv_filtered = create_datetime_index(csv_filtered)
    csv_filtered, selected_year = prompt_select_year(csv_filtered)
    
    # Step 4: Fill missing periods
    print("\nStep 4: Filling missing periods")
    csv_filtered = fill_missing_periods(csv_filtered, selected_year)
    
    # Assemble processing package
    processing_data = {
        'elexon_id': selected_elexon_id,
        'year': selected_year,
        'csv_profile': csv_filtered,
        'main_data': filtered_excel_data.get('MAIN DATA'),
        'dg_data': filtered_excel_data.get('DG'),
        'sub1mw_data': filtered_excel_data.get('Sub1MW'),
        'gsp_info': filtered_excel_data.get('GSP info'),
    }
    
    print("\n" + "="*80)
    print("PROCESSING DATA READY")
    print(f"Elexon ID: {selected_elexon_id}")
    print(f"Year: {selected_year}")
    print(f"CSV Profile shape: {processing_data['csv_profile'].shape}")
    if processing_data['main_data'] is not None and len(processing_data['main_data']) > 0:
        print(f"MAIN DATA shape: {processing_data['main_data'].shape}")
    if processing_data['dg_data'] is not None and len(processing_data['dg_data']) > 0:
        print(f"DG shape: {processing_data['dg_data'].shape}")
    if processing_data['sub1mw_data'] is not None and len(processing_data['sub1mw_data']) > 0:
        print(f"Sub1MW shape: {processing_data['sub1mw_data'].shape}")
    print("="*80 + "\n")
    
    return processing_data


def plot_one_day_profile(df: pd.DataFrame, elexon_id: str, year: int) -> None:
    """
    Plot one day of data at half-hourly resolution.
    Prompts user to select which day to plot.
    
    Args:
        df: Dataframe with DatetimeIndex and data columns
        elexon_id: The selected Elexon ID (for plot title)
        year: The selected year (for plot title)
    """
    if not isinstance(df.index, pd.DatetimeIndex):
        raise TypeError("The dataframe index must be a DatetimeIndex.")
    
    # Get unique dates
    unique_dates = sorted(pd.Series(df.index.date).unique().tolist())
    if not unique_dates:
        print("No data available to plot.")
        return
    
    print(f"\n{'='*80}")
    print("PLOT ONE DAY - HALF HOURLY RESOLUTION")
    print(f"{'='*80}")
    print(f"\nAvailable dates ({len(unique_dates)} total):")
    
    # Show first 5 and last 5 dates
    if len(unique_dates) > 10:
        for idx in range(min(5, len(unique_dates))):
            print(f"{idx + 1}. {unique_dates[idx]}")
        print(f"... ({len(unique_dates) - 10} more dates) ...")
        for idx in range(max(5, len(unique_dates) - 5), len(unique_dates)):
            print(f"{idx + 1}. {unique_dates[idx]}")
    else:
        for idx in range(len(unique_dates)):
            print(f"{idx + 1}. {unique_dates[idx]}")
    
    while True:
        selection = input("\nEnter the number of the date to plot (or 'q' to skip): ").strip()
        if selection.lower() == 'q':
            print("Skipping single day plot.")
            return
        
        if not selection.isdigit():
            print("Please enter a valid number.")
            continue
        
        selection_index = int(selection)
        if 1 <= selection_index <= len(unique_dates):
            selected_date = unique_dates[selection_index - 1]
            day_data = df[df.index.date == selected_date]
            
            if len(day_data) == 0:
                print(f"No data for date {selected_date}")
                continue
            
            # Create plot
            fig, ax = plt.subplots(figsize=(14, 6))
            
            # Plot Meter Volume column
            if "Meter Volume" in day_data.columns:
                ax.plot(day_data.index, day_data["Meter Volume"], marker='o', label="Meter Volume", linewidth=2, markersize=4)
                ax.set_xlabel("Time (Half-hourly)", fontsize=12)
                ax.set_ylabel("Meter Volume", fontsize=12)
            else:
                print(f"Column 'Meter Volume' not found. Available columns: {list(day_data.columns)}")
                continue
            ax.set_title(f"Daily Profile - {selected_date}\nElexon ID: {elexon_id}, Year: {year}", fontsize=14, fontweight='bold')
            ax.grid(True, alpha=0.3)
            ax.legend(loc='best')
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))
            ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
            fig.autofmt_xdate(rotation=45, ha='right')
            plt.tight_layout()
            plt.show()
            break
        
        print(f"Please enter a number between 1 and {len(unique_dates)}.")


def plot_yearly_rolling_average(df: pd.DataFrame, elexon_id: str, year: int, window_days: int = 3) -> None:
    """
    Plot the entire year with a rolling average.
    
    Args:
        df: Dataframe with DatetimeIndex and data columns
        elexon_id: The selected Elexon ID (for plot title)
        year: The selected year (for plot title)
        window_days: Number of days for rolling average (default 3)
    """
    if not isinstance(df.index, pd.DatetimeIndex):
        raise TypeError("The dataframe index must be a DatetimeIndex.")
    
    print(f"\n{'='*80}")
    print(f"PLOT YEARLY DATA - {window_days}-DAY ROLLING AVERAGE")
    print(f"{'='*80}")
    
    # Check if Meter Volume column exists
    if "Meter Volume" not in df.columns:
        print(f"Column 'Meter Volume' not found. Available columns: {list(df.columns)}")
        return
    
    # Calculate rolling average (window in hours: days * 48 half-hourly periods)
    window_periods = window_days * 48
    df_rolling = df[["Meter Volume"]].rolling(window=window_periods, center=True, min_periods=1).mean()
    
    # Create plot
    fig, ax = plt.subplots(figsize=(16, 7))
    
    # Plot rolling average
    ax.plot(df_rolling.index, df_rolling["Meter Volume"], label=f"Meter Volume ({window_days}-day MA)", linewidth=2)
    
    ax.set_xlabel("Date", fontsize=12)
    ax.set_ylabel(f"Meter Volume ({window_days}-day Rolling Average)", fontsize=12)
    ax.set_title(f"Yearly Profile with {window_days}-Day Rolling Average\nElexon ID: {elexon_id}, Year: {year}", 
                 fontsize=14, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.legend(loc='best')
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
    fig.autofmt_xdate(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()
    
    print(f"✓ Yearly rolling average plot displayed")
    print(f"  Total data points: {len(df)}")
    print(f"  Rolling window: {window_days} days ({window_periods} periods)")


if __name__ == "__main__":
    read_csv_files = True

    # Read excel data
    print("Loading excel data...")
    excel_data = read_excel_file()
    
    # Process CSV files and prepare all data for transformation
    if read_csv_files == True:
        print("\nLoading CSV data...")
        all_csv_data = read_ordered_gp9_csv_files()
        
        # Consolidate data selection and filtering
        processing_data = prepare_processing_data(excel_data, all_csv_data)
        
        # Extract components for easier access
        csv_profile = processing_data['csv_profile']
        main_data = processing_data['main_data']
        elexon_id = processing_data['elexon_id']
        year = processing_data['year']
        
        # Save CSV profile
        csv_profile.to_csv("filtered_gp9_data.csv")
        print(f"Saved filtered CSV profile to 'filtered_gp9_data.csv'")
        print(f"\nProfile data preview:\n{csv_profile.head()}")
        
        # Display MAIN DATA if available
        if main_data is not None:
            print(f"\nMAIN DATA for Elexon ID {elexon_id}:\n{main_data}")
        
        print(f"\n✓ Data preparation complete and ready for profile transformation")
        print(f"  - Use 'processing_data' dict to access all excel and csv data")
        print(f"  - processing_data['csv_profile']: Contains the profile data with set points")
        print(f"  - processing_data['main_data']: Contains the max/min set points from excel")
        
        # Plot the selected data
        print(f"\n{'='*80}")
        print("VISUALIZATION")
        print(f"{'='*80}")
        
        plot_one_day_profile(csv_profile, elexon_id, year)
        plot_yearly_rolling_average(csv_profile, elexon_id, year, window_days=3)

