import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt  # type: ignore
import matplotlib.dates as mdates  # type: ignore
from datetime import timedelta

FILENAME_PATTERN = re.compile(r"^GP9_2023(\d{2})\.csv$")
ACTIVE_REQUIRED_COLUMNS = {"scenario", "year", "DemandPk", "DemandAM", "DemandPM", "type"}
ACTIVE_DEMAND_COLUMNS = ["DemandPk", "DemandAM", "DemandPM"]
GROSS_DEMAND_SETPOINTS = {
    "winter_peak": {
        "label": "Winter Peak",
        "active_column": "DemandPk",
        "lv_gain_column": "Peak",
        "storage_column": "wintpk",
    },
    "summer_min_am": {
        "label": "Summer Min AM",
        "active_column": "DemandAM",
        "lv_gain_column": "SummerAM",
        "storage_column": "summam",
    },
    "summer_min_pm": {
        "label": "Summer Min PM",
        "active_column": "DemandPM",
        "lv_gain_column": "SummerPM",
        "storage_column": "summpm",
    },
}
FES_CONTRIBUTION_SHEETS = {
    "LV Gain": {
        "scenario_column": "Scenario",
        "year_column": "Year",
        "value_columns": {
            "winter_peak": "Peak",
            "summer_min_am": "SummerAM",
            "summer_min_pm": "SummerPM",
        },
    },
    "mBattery": {
        "scenario_column": "scenario",
        "year_column": "year",
        "value_columns": {
            "winter_peak": "wintpk",
            "summer_min_am": "summam",
            "summer_min_pm": "summpm",
        },
    },
    "DxStorage": {
        "scenario_column": "scenario",
        "year_column": "year",
        "value_columns": {
            "winter_peak": "wintpk",
            "summer_min_am": "summam",
            "summer_min_pm": "summpm",
        },
    },
}


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


def prompt_select_fes_scenario(active_df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    if "scenario" not in active_df.columns:
        raise KeyError("The Active dataframe does not contain a 'scenario' column.")

    unique_scenarios = sorted(active_df["scenario"].dropna().astype(str).unique())
    if not unique_scenarios:
        raise ValueError("No scenario values were found in the filtered Active dataframe.")

    print("\nAvailable FES scenarios:")
    for index, scenario in enumerate(unique_scenarios, start=1):
        print(f"{index}. {scenario}")

    while True:
        selection = input("Enter the number of the FES scenario to use for scaling: ").strip()
        if not selection.isdigit():
            print("Please enter a valid number.")
            continue

        selection_index = int(selection)
        if 1 <= selection_index <= len(unique_scenarios):
            selected_scenario = unique_scenarios[selection_index - 1]
            scenario_df = active_df[active_df["scenario"].astype(str) == selected_scenario].reset_index(drop=True)
            print(f"Selected FES scenario '{selected_scenario}' with {len(scenario_df)} Active rows.")
            return scenario_df, selected_scenario

        print(f"Please enter a number between 1 and {len(unique_scenarios)}.")


def prompt_select_fes_year(active_df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    if "year" not in active_df.columns:
        raise KeyError("The Active dataframe does not contain a 'year' column.")

    unique_years = sorted(active_df["year"].dropna().unique())
    if not unique_years:
        raise ValueError("No year values were found in the filtered Active dataframe.")

    print("\nAvailable FES years:")
    for index, year in enumerate(unique_years, start=1):
        print(f"{index}. {int(year)}")

    while True:
        selection = input("Enter the number of the FES year to use for scaling: ").strip()
        if not selection.isdigit():
            print("Please enter a valid number.")
            continue

        selection_index = int(selection)
        if 1 <= selection_index <= len(unique_years):
            selected_year = unique_years[selection_index - 1]
            year_df = active_df[active_df["year"] == selected_year].reset_index(drop=True)
            selected_year = int(selected_year)
            print(f"Selected FES year {selected_year} with {len(year_df)} Active rows.")
            return year_df, selected_year

        print(f"Please enter a number between 1 and {len(unique_years)}.")


def calculate_active_demand_totals(active_df: pd.DataFrame) -> dict:
    missing_columns = sorted(ACTIVE_REQUIRED_COLUMNS - set(active_df.columns))
    if missing_columns:
        raise KeyError(f"The Active dataframe is missing required columns: {missing_columns}")

    if active_df.empty:
        raise ValueError("No Active rows were found for the selected GSP, FES scenario, and FES year.")

    demand_values = active_df[ACTIVE_DEMAND_COLUMNS].apply(pd.to_numeric, errors="coerce")
    if demand_values.isna().any().any():
        null_columns = demand_values.columns[demand_values.isna().any()].tolist()
        raise ValueError(f"The Active dataframe contains non-numeric demand values in: {null_columns}")

    totals = demand_values.sum()
    return {column: float(totals[column]) for column in ACTIVE_DEMAND_COLUMNS}


def calculate_fes_sheet_contributions(
    sheet_name: str,
    df: pd.DataFrame,
    scenario: str,
    year: int,
) -> dict:
    if sheet_name not in FES_CONTRIBUTION_SHEETS:
        raise KeyError(f"No FES contribution mapping has been configured for sheet '{sheet_name}'.")

    config = FES_CONTRIBUTION_SHEETS[sheet_name]
    required_columns = {
        config["scenario_column"],
        config["year_column"],
        *config["value_columns"].values(),
    }
    missing_columns = sorted(required_columns - set(df.columns))
    if missing_columns:
        raise KeyError(f"The {sheet_name} dataframe is missing required columns: {missing_columns}")

    year_values = pd.to_numeric(df[config["year_column"]], errors="coerce")
    matching_rows = df[
        (df[config["scenario_column"]].astype(str) == str(scenario))
        & (year_values == int(year))
    ].reset_index(drop=True)

    if matching_rows.empty:
        print(f"No {sheet_name} rows found for scenario '{scenario}', FES year {year}; using 0 contribution.")
        return {
            setpoint_key: {
                "value": 0.0,
                "row_count": 0,
                "source_column": source_column,
            }
            for setpoint_key, source_column in config["value_columns"].items()
        }

    contributions = {}
    for setpoint_key, source_column in config["value_columns"].items():
        values = pd.to_numeric(matching_rows[source_column], errors="coerce")
        if values.isna().any():
            raise ValueError(f"The {sheet_name} dataframe contains non-numeric values in '{source_column}'.")

        contributions[setpoint_key] = {
            "value": float(values.sum()),
            "row_count": len(matching_rows),
            "source_column": source_column,
        }

    return contributions


def calculate_gross_demand_setpoints(
    filtered_excel_data: dict,
    active_demand_totals: dict,
    scenario: str,
    year: int,
) -> dict:
    gross_setpoints = {}
    sheet_contributions = {}

    for sheet_name in FES_CONTRIBUTION_SHEETS:
        sheet_df = filtered_excel_data.get(sheet_name)
        if sheet_df is None:
            raise KeyError(f"The {sheet_name} sheet was not loaded or filtered.")

        sheet_contributions[sheet_name] = calculate_fes_sheet_contributions(
            sheet_name,
            sheet_df,
            scenario,
            year,
        )

    for setpoint_key, setpoint_config in GROSS_DEMAND_SETPOINTS.items():
        active_value = float(active_demand_totals[setpoint_config["active_column"]])
        components = {"Active": active_value}

        for sheet_name in FES_CONTRIBUTION_SHEETS:
            components[sheet_name] = sheet_contributions[sheet_name][setpoint_key]["value"]

        gross_setpoints[setpoint_key] = {
            "label": setpoint_config["label"],
            "components": components,
            "gross_demand": float(sum(components.values())),
        }

    return gross_setpoints


def print_gross_demand_setpoints(gross_demand_setpoints: dict) -> None:
    print("\nGross demand setpoints:")
    for setpoint_key in GROSS_DEMAND_SETPOINTS:
        setpoint = gross_demand_setpoints[setpoint_key]
        print(f"  {setpoint['label']}: {setpoint['gross_demand']:,.6f}")
        for component_name, value in setpoint["components"].items():
            print(f"    {component_name}: {value:,.6f}")


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
            - 'year': The selected CSV source year
            - 'csv_year': The selected CSV source year
            - 'fes_scenario': The selected Regional FES scenario
            - 'fes_year': The selected Regional FES target year
            - 'active_demand_totals': Active DemandPk/DemandAM/DemandPM totals for scaling
            - 'gross_demand_setpoints': Gross demand setpoints with Active/LV Gain/storage components
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
    print("\nStep 3: Processing CSV data (datetime index and source year selection)")
    csv_filtered = create_datetime_index(csv_filtered)
    csv_filtered, selected_csv_year = prompt_select_year(csv_filtered)
    
    # Step 4: Fill missing periods
    print("\nStep 4: Filling missing periods")
    csv_filtered = fill_missing_periods(csv_filtered, selected_csv_year)

    # Step 5: Select Regional FES scenario/year and calculate Active demand totals
    print("\nStep 5: Selecting Regional FES scenario/year and calculating Active demand totals")
    active_data = filtered_excel_data.get("Active")
    if active_data is None:
        raise KeyError("The Active sheet was not loaded or filtered.")
    if active_data.empty:
        raise ValueError(f"No Active rows were found for selected GSP '{selected_elexon_id}'.")

    missing_active_columns = sorted(ACTIVE_REQUIRED_COLUMNS - set(active_data.columns))
    if missing_active_columns:
        raise KeyError(f"The Active dataframe is missing required columns: {missing_active_columns}")

    active_scenario_data, selected_fes_scenario = prompt_select_fes_scenario(active_data)
    active_year_data, selected_fes_year = prompt_select_fes_year(active_scenario_data)
    active_demand_totals = calculate_active_demand_totals(active_year_data)
    gross_demand_setpoints = calculate_gross_demand_setpoints(
        filtered_excel_data,
        active_demand_totals,
        selected_fes_scenario,
        selected_fes_year,
    )
    
    # Assemble processing package
    processing_data = {
        'elexon_id': selected_elexon_id,
        'year': selected_csv_year,
        'csv_year': selected_csv_year,
        'fes_scenario': selected_fes_scenario,
        'fes_year': selected_fes_year,
        'active_demand_totals': active_demand_totals,
        'gross_demand_setpoints': gross_demand_setpoints,
        'csv_profile': csv_filtered,
        'main_data': filtered_excel_data.get('MAIN DATA'),
        'dg_data': filtered_excel_data.get('DG'),
        'sub1mw_data': filtered_excel_data.get('Sub1MW'),
        'dxstorage_data': filtered_excel_data.get('DxStorage'),
        'mbattery_data': filtered_excel_data.get('mBattery'),
        'lv_gain_data': filtered_excel_data.get('LV Gain'),
        'active_data': active_year_data,
        'gsp_info': filtered_excel_data.get('GSP info'),
    }
    
    print("\n" + "="*80)
    print("PROCESSING DATA READY")
    print(f"Elexon ID: {selected_elexon_id}")
    print(f"CSV source year: {selected_csv_year}")
    print(f"FES scenario: {selected_fes_scenario}")
    print(f"FES target year: {selected_fes_year}")
    print(f"CSV Profile shape: {processing_data['csv_profile'].shape}")
    print("Active demand totals for scaling:")
    for column in ACTIVE_DEMAND_COLUMNS:
        print(f"  {column}: {active_demand_totals[column]:,.6f}")
    print_gross_demand_setpoints(gross_demand_setpoints)
    if processing_data['main_data'] is not None and len(processing_data['main_data']) > 0:
        print(f"MAIN DATA shape: {processing_data['main_data'].shape}")
    if processing_data['dg_data'] is not None and len(processing_data['dg_data']) > 0:
        print(f"DG shape: {processing_data['dg_data'].shape}")
    if processing_data['sub1mw_data'] is not None and len(processing_data['sub1mw_data']) > 0:
        print(f"Sub1MW shape: {processing_data['sub1mw_data'].shape}")
    if processing_data['dxstorage_data'] is not None and len(processing_data['dxstorage_data']) > 0:
        print(f"DxStorage shape: {processing_data['dxstorage_data'].shape}")
    if processing_data['mbattery_data'] is not None and len(processing_data['mbattery_data']) > 0:
        print(f"mBattery shape: {processing_data['mbattery_data'].shape}")
    if processing_data['lv_gain_data'] is not None and len(processing_data['lv_gain_data']) > 0:
        print(f"LV Gain shape: {processing_data['lv_gain_data'].shape}")
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
        fes_scenario = processing_data['fes_scenario']
        fes_year = processing_data['fes_year']
        active_demand_totals = processing_data['active_demand_totals']
        gross_demand_setpoints = processing_data['gross_demand_setpoints']
        
        # Save CSV profile
        csv_profile.to_csv("filtered_gp9_data.csv")
        print(f"Saved filtered CSV profile to 'filtered_gp9_data.csv'")
        print(f"\nProfile data preview:\n{csv_profile.head()}")
        
        # Display MAIN DATA if available
        if main_data is not None:
            print(f"\nMAIN DATA for Elexon ID {elexon_id}:\n{main_data}")

        print(f"\nActive demand totals for FES scenario {fes_scenario}, FES year {fes_year}:")
        for column in ACTIVE_DEMAND_COLUMNS:
            print(f"  {column}: {active_demand_totals[column]:,.6f}")

        print_gross_demand_setpoints(gross_demand_setpoints)
        
        print(f"\n✓ Data preparation complete and ready for profile transformation")
        print(f"  - Use 'processing_data' dict to access all excel and csv data")
        print(f"  - processing_data['csv_profile']: Contains the profile data with set points")
        print(f"  - processing_data['main_data']: Contains the max/min set points from excel")
        print(f"  - processing_data['active_demand_totals']: Contains Active DemandPk/DemandAM/DemandPM totals")
        print(f"  - processing_data['gross_demand_setpoints']: Contains the three gross demand setpoints")
        
        # Plot the selected data
        print(f"\n{'='*80}")
        print("VISUALIZATION")
        print(f"{'='*80}")
        
        plot_one_day_profile(csv_profile, elexon_id, year)
        plot_yearly_rolling_average(csv_profile, elexon_id, year, window_days=3)

