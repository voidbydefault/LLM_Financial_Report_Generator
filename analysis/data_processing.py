#### analysis/data_processing.py
####
#### This script uses Pandas to load, preprocess, and transform raw data from an Excel file.
#### It performs the following key tasks:
####   - Loads data from the specified Excel file.
####   - Reshapes the data from a wide to a long format, making it suitable for analysis.
####   - Calculates annual revenue by property and tenant.
####   - Identifies the top tenants for each property.
####   - Analyzes revenue changes between years to pinpoint key drivers.
####   - Exports analysis results to an Excel file for further review.

import pandas as pd
import os
from typing import Tuple


def load_and_preprocess_data(file_path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Loads data from Excel, reshapes it, and performs initial calculations."""
    try:
        df = pd.read_excel(file_path)
        print("Columns in the dataset:", df.columns)

        df_long = pd.melt(df, id_vars=['tenant', 'property'], var_name='Date', value_name='Revenue')
        df_long['Date'] = pd.to_datetime(df_long['Date'], errors='coerce')
        df_long['Year'] = df_long['Date'].dt.year

        return df_long, df
    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
        return pd.DataFrame(), pd.DataFrame()  # Return empty DataFrames
    except Exception as e:
        print(f"Error loading and preprocessing data: {e}")
        return pd.DataFrame(), pd.DataFrame()


def calculate_annual_revenue(df_long: pd.DataFrame) -> pd.DataFrame:
    """Calculates annual revenue by property."""
    annual_revenue = df_long.groupby(['property', 'Year'])['Revenue'].sum().reset_index()
    annual_revenue = annual_revenue.sort_values(['property', 'Year'])
    annual_revenue['Revenue_Change'] = annual_revenue.groupby('property')['Revenue'].diff()  # Annual Change
    return annual_revenue


def calculate_top_tenants(df_long: pd.DataFrame) -> pd.DataFrame:
    """Calculates top 10 tenants per property."""
    tenant_revenue = df_long.groupby(['property', 'tenant'])['Revenue'].sum().reset_index()
    return tenant_revenue


def tenant_change_analysis(tenant_annual_rev: pd.DataFrame, prop: str, year: int) -> pd.DataFrame:
    """Analyzes tenant revenue changes between two years for a given property."""
    current = tenant_annual_rev[(tenant_annual_rev['property'] == prop) & (tenant_annual_rev['Year'] == year)]
    previous = tenant_annual_rev[(tenant_annual_rev['property'] == prop) & (tenant_annual_rev['Year'] == year - 1)]
    comparison = pd.merge(current, previous, on='tenant', how='outer', suffixes=('_curr', '_prev')).fillna(0)
    comparison['Revenue_Change'] = comparison['Revenue_curr'] - comparison['Revenue_prev']
    return comparison.sort_values(by='Revenue_Change', ascending=False)


def calculate_tenant_changes(df_long: pd.DataFrame) -> dict:
    """Calculates tenant revenue changes for each property and year."""
    tenant_annual_rev = df_long.groupby(['property', 'Year', 'tenant'])['Revenue'].sum().reset_index()
    tenant_change_results = {}
    for prop in tenant_annual_rev['property'].unique():
        years = sorted(tenant_annual_rev[tenant_annual_rev['property'] == prop]['Year'].unique())
        for i in range(1, len(years)):
            change_df = tenant_change_analysis(tenant_annual_rev, prop, years[i])
            key = f"{prop}: {years[i - 1]} to {years[i]}"
            tenant_change_results[key] = change_df
    return tenant_change_results


def export_analysis_results(annual_revenue: pd.DataFrame, tenant_change_results: dict, output_dir: str) -> None:
    """Exports analysis results to an Excel file."""
    output_file = os.path.join(output_dir, 'analysis_results.xlsx')
    try:
        with pd.ExcelWriter(output_file) as writer:
            annual_revenue.to_excel(writer, sheet_name='Annual_Revenue', index=False)
            for key, change_df in tenant_change_results.items():
                sheet_name = key.replace(" ", "").replace(":", "_")
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                change_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"\nAnalysis results exported to '{output_file}'.")
    except Exception as e:
        print(f"Error exporting analysis results: {e}")