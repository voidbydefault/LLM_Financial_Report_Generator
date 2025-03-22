#### analysis/report_generation.py
####
#### This script orchestrates the entire report generation process.
#### It calls functions from the other modules to:
####   - Load and preprocess the data.
####   - Perform data analysis and calculations.
####   - Generate visualizations.
####   - Query the LLM for insights and commentary.
####   - Combine all of the above to generate a markdown-formatted executive report.


import os
import pandas as pd
from typing import Dict, List
from analysis.data_processing import load_and_preprocess_data, calculate_annual_revenue, \
    calculate_top_tenants, calculate_tenant_changes, export_analysis_results
from analysis.llm_interaction import query_llm
from analysis.visualizations import generate_visualizations
import logging


def generate_markdown_report(annual_revenue: pd.DataFrame, tenant_revenue: pd.DataFrame, visualization_paths: Dict[str, List[str]],
                             tenant_change_results: Dict, model_name: str, temperature: float, max_tokens: int) -> str:
    """Generates the markdown report content."""

    # Basic Statistics
    total_revenue = annual_revenue.groupby('Year')['Revenue'].sum()
    top_properties = annual_revenue.groupby('property')['Revenue'].sum().nlargest(5)

    # Generate LLM commentary for annual revenue with direct data
    annual_prompt_data = "\n".join(
        [f"- {year}: ${rev:,.0f} ({rev / total_revenue.sum() * 100:.1f}% of total)"
         for year, rev in total_revenue.items()]
    )

    annual_commentary = query_llm(
        f"""Analyze these annual revenue trends using ONLY the provided data:
        {annual_prompt_data}

        Required analysis:
        1. Year-over-year changes in total revenue
        2. Property performance relative to each other
        3. Notable percentage contributions
        4. State "No notable changes" if under 5% variance

        Output format:
        - Start with overall trend summary
        - Bullet points of key observations
        - End with largest contributor percentage""", model_name, temperature, max_tokens
    )

    # Create markdown content
    md_content = f"""# Executive Sales Report

## Annual Revenue Overview
![Annual Revenue by Property](visualizations/annual_revenue.png)

{annual_commentary or '*Revenue commentary not available*'}

- Total Revenue by Year:
{chr(10).join(f'  - **{year}**: ${rev:,.2f}' for year, rev in total_revenue.items())}

- Top 5 Properties by Total Revenue:
{chr(10).join(f'  - **{prop}**: ${rev:,.2f}' for prop, rev in top_properties.items())}

## Tenant Performance
"""

    # Add top tenants visualizations with commentary
    md_content += "\n## Top Tenants by Property\n"
    for path in visualization_paths['top_tenants']:
        filename = os.path.basename(path).replace('.png', '')
        prop_name = ' '.join(filename.split('_')[2:])
        prop_data = tenant_revenue[tenant_revenue['property'] == prop_name].nlargest(10, 'Revenue')

        if prop_data.empty:
            continue

        total_prop_rev = prop_data['Revenue'].sum()
        tenant_list = "\n".join(
            [f"{row['tenant']}: ${row['Revenue']:,.0f} ({row['Revenue'] / total_prop_rev * 100:.1f}%)"
             for _, row in prop_data.iterrows()]
        )

        tenant_commentary = query_llm(
            f"""Analyze tenant distribution for {prop_name} using:
            Total property revenue: ${total_prop_rev:,.0f}
            Tenant breakdown:
            {tenant_list}

            Required analysis:
            1. Top 3 tenant contributions as percentages
            2. Concentration risk assessment
            3. Compare top/bottom performer amounts
            4. State "Balanced distribution" if top 3 < 60%

            Output format:
            - Summary statement
            - Key percentages
            - Risk assessment""", model_name, temperature, max_tokens
        )

        md_content += f"\n### {prop_name}\n"
        md_content += f"![Top Tenants - {prop_name}]({path})\n"
        md_content += f"{tenant_commentary or '*Tenant analysis unavailable*'}\n"

    # Add revenue change analysis with direct data
    md_content += "\n## Significant Revenue Changes\n"
    for key in tenant_change_results:
        try:
            prop, years = key.split(": ")
            df = tenant_change_results[key]

            # Get top 5 positive and negative changes
            top_gains = df.nlargest(3, 'Revenue_Change')
            top_losses = df.nsmallest(3, 'Revenue_Change')

            change_data = "Significant Changes:\n"
            change_data += "Growth Contributors:\n" + "\n".join(
                [f"{row['tenant']}: +${row['Revenue_Change']:,.0f}"
                 for _, row in top_gains.iterrows() if row['Revenue_Change'] > 1000]
            ) + "\n"
            change_data += "Significant Losses:\n" + "\n".join(
                [f"{row['tenant']}: -${abs(row['Revenue_Change']):,.0f}"
                 for _, row in top_losses.iterrows() if row['Revenue_Change'] < -1000]
            )

            change_commentary = query_llm(
                f"""Analyze revenue changes for {prop} ({years}):
                {change_data}

                Required analysis:
                1. Largest absolute changes
                2. Net impact calculation
                3. Notable loss/gain ratios
                4. State "No significant changes" if all < $1k

                Output format:
                - Net change summary
                - Top 3 contributors
                - Loss mitigation suggestions""", model_name, temperature, max_tokens
            )

            # Find matching visualization
            safe_prop = prop.replace(" ", "_")
            safe_years = years.replace(" ", "_").replace(":", "_")
            pattern = f"top10_changes_{safe_prop}_{safe_years}.png"
            matches = [p for p in visualization_paths['revenue_changes'] if pattern in p]

            viz_section = ""
            if matches:
                path = matches[0]
                viz_section = f"![Revenue Changes - {prop} ({years})]({path})\n"

            md_content += f"\n### {prop} ({years})\n"
            md_content += viz_section
            md_content += f"{change_commentary or '*Change analysis unavailable*'}\n"

        except Exception as e:
            print(f"Skipping {key} due to error: {str(e)}")
            logging.error(f"Skipping {key} due to error: {str(e)}")
            continue

    return md_content


def generate_report(file_path: str, output_dir: str, model_name: str, temperature: float, max_tokens: int) -> str | None:
    """Generates the full report."""
    # Load and preprocess data
    df_long, df = load_and_preprocess_data(file_path)

    if df_long.empty or df.empty:
        print("Failed to load and preprocess data.  Check the file and its format.")
        return None

    # Perform calculations
    annual_revenue = calculate_annual_revenue(df_long)
    tenant_revenue = calculate_top_tenants(df_long)
    tenant_change_results = calculate_tenant_changes(df_long)

    # Export analysis results to Excel
    export_analysis_results(annual_revenue, tenant_change_results, output_dir)

    # Generate visualizations
    visualization_paths = generate_visualizations(annual_revenue, tenant_revenue, tenant_change_results, output_dir)

    # Generate markdown report
    report_md = generate_markdown_report(annual_revenue, tenant_revenue, visualization_paths,
                                          tenant_change_results, model_name, temperature, max_tokens)
    return report_md