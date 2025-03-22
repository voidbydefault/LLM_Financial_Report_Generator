# analysis/visualizations.py
import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from typing import Dict, List


def generate_visualizations(annual_revenue: pd.DataFrame, tenant_revenue: pd.DataFrame,
                            tenant_change_results: dict, output_dir: str) -> Dict[str, List[str]]:
    """Generates visualizations and saves them to the output directory."""

    viz_dir = os.path.join(output_dir, "visualizations")
    os.makedirs(viz_dir, exist_ok=True)

    visualization_paths = {
        'annual_revenue': [],
        'top_tenants': [],
        'revenue_changes': []
    }

    # 1. Annual Revenue by Property
    plt.figure(figsize=(12, 6))
    sns.barplot(data=annual_revenue, x='property', y='Revenue', hue='Year', palette='viridis')
    plt.title('Annual Revenue by Property')
    plt.xlabel('Property')
    plt.ylabel('Revenue (USD)')
    plt.xticks(rotation=45, ha="right")  # Rotate x-axis labels for better readability
    plt.tight_layout()
    annual_filename = 'annual_revenue.png'
    plt.savefig(os.path.join(viz_dir, annual_filename))
    plt.close()
    visualization_paths['annual_revenue'].append(f"visualizations/{annual_filename}")  # Relative path

    # 2. Top Tenants per Property
    for prop in tenant_revenue['property'].unique():
        plt.figure(figsize=(10, 6))
        prop_data = tenant_revenue[tenant_revenue['property'] == prop].nlargest(10, 'Revenue')
        sns.barplot(data=prop_data, x='Revenue', y='tenant', hue='tenant', palette='rocket', legend=False)
        plt.title(f'Top 10 Tenants - {prop}')
        plt.xlabel('Total Revenue (USD)')
        plt.ylabel('Tenant')
        plt.tight_layout()
        filename = f'top_tenants_{prop}.png'.replace(" ", "_")
        plt.savefig(os.path.join(viz_dir, filename))
        plt.close()
        visualization_paths['top_tenants'].append(f"visualizations/{filename}")  # Relative path

    # 3. Revenue Change Analysis
    for key in tenant_change_results:
        prop, years = key.split(": ")
        df = tenant_change_results[key]
        top_changes = df.reindex(df['Revenue_Change'].abs().sort_values(ascending=False).index).head(10)

        if not top_changes.empty:
            plt.figure(figsize=(10, 6))
            top_changes = top_changes.sort_values('Revenue_Change', ascending=False)
            sns.barplot(data=top_changes,
                        x='Revenue_Change',
                        y='tenant',
                        palette='coolwarm',
                        hue=np.sign(top_changes['Revenue_Change']),
                        dodge=False,
                        legend=False)  # disable legend
            plt.title(f'Top 10 Revenue Changes: {prop} ({years})')
            plt.xlabel('Revenue Change (USD)')
            plt.ylabel('Tenant')
            plt.tight_layout()
            filename = f'top10_changes_{prop}_{years}.png'.replace(" ", "_").replace(":", "_")
            plt.savefig(os.path.join(viz_dir, filename))
            plt.close()
            visualization_paths['revenue_changes'].append(f"visualizations/{filename}")  # Relative path

    return visualization_paths