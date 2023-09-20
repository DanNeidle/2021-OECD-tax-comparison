#!/usr/bin/env python

from openpyxl import load_workbook
import json
import os
import plotly.graph_objects as go
import imageio
from PIL import Image
import io
import numpy as np

wb_file = "all_oecd_composition_data_1990_to_2021.xlsx"
start_year = 1990
end_year = 2021
logo_file = "logo_full_white_on_blue.jpg"


# note to stop gif looping
# ffmpeg -i tax_composition_OECD.gif -loop 0 -vcodec copy tax_composition_OECD_noloop.gif

# a lookup totals showing which columns are summed to produce tax (negative means you minus it)
tax_types = {
    "total_revenue": [0],
    "income_tax": [2, -4, 8],   # for some countries IT income is missing, so let's use total personal tax less CGT
    "NI": [9, 22],
    "VAT": [41],
    "non_VAT_sales": [38, -41],
    "corporate": [5],
    "property": [4, 23],
}

# order of bar stacking, from bottom
tax_order = ["income_tax", "NI", "VAT", "non_VAT_sales", "corporate", "property"]

tax_labels = {
    "income_tax": "Personal income tax",
    "NI": "National insurance/social security",
    "corporate": "Corporate tax",
    "property": "Property and wealth taxes",
    "VAT": "Value Added Tax",
    "non_VAT_sales": "Other taxes on goods/services"
}

def load_in_workbook(file):
    print("Importing data from excel")

    wb = load_workbook(filename=wb_file, read_only=True, data_only=True)
    ws = wb['OECD.Stat export']

    output = []
    for row in ws.rows:
        output.append(row)
        
    wb.close()
    return output

def save_to_json(filename, data):
    data_str_keys = {str(key): value for key, value in data.items()}
    
    with open(f"{filename}.json", 'w') as f:
        json.dump(data_str_keys, f, default=str)
        
def load_from_json(filename):
    with open(f"{filename}.json", 'r') as f:
        return json.load(f)
    

def has_content(cell):
    if cell is None or cell == "":
        return False
    else:
        return True
    

def process_raw_data(excel_data):
    tax_revenue = {}
    for excel_row in excel_data[11:]:
        
        # is it OECD or non-OECD?
        if has_content(excel_row[0].value):
            country = excel_row[0].value.strip()
            tax_revenue[country] = {"OECD": True}
        elif has_content(excel_row[1].value):
            country = excel_row[1].value.strip()
            tax_revenue[country] = {"OECD": False}
        else:
            continue
        
        has_data = False

        for year in range(start_year, end_year + 1):
            
            start_column = (year - 1990) * 65 + 3
            
            # if no data then use the previous year's data (if we have it!)
            if has_data and (excel_row[start_column].value == 0 or not has_content(excel_row[start_column].value)):
                tax_revenue[country][str(year)] = tax_revenue[country][str(year - 1)]
                continue
            
            has_data = True
            tax_revenue[country][str(year)] = {}
            
            
            for tax in tax_types:
                total = 0
                for column in tax_types[tax]:
                    
                    # where the column is negative that means we subtract that number!
                    
                    if column >= 0:
                        if has_content(excel_row[start_column + column].value):
                            total += excel_row[start_column + column].value
                    else:
                        if has_content(excel_row[start_column - column].value):
                            total -= excel_row[start_column - column].value
                
                tax_revenue[country][str(year)][tax] = total
    
    return tax_revenue


def load_logo(logo_path):
    return dict(
        source=Image.open(logo_path),
        xref="paper", yref="paper",
        x=1.0, y=1.01,
        sizex=0.08, sizey=0.08,
        xanchor="right", yanchor="bottom"
    )


def create_gif(data, mode, highlight_country):
    filename = f"tax_composition_{mode}.gif"
    frame_duration = 0.1  # in seconds for each intermediate frame

    gif_writer = imageio.get_writer(filename, mode='I', duration=frame_duration, loop=1)

    n_frames_between_years = 5  # Number of intermediate frames

    for year in range(start_year, end_year):
        print(f"Generating for transition {year} to {year+1}")

        # Interpolate data for smoother transition
        for i in range(n_frames_between_years + 1):
            weight = i / n_frames_between_years  # Weight for linear interpolation
            interpolated_data = interpolate_data(data, year, weight, mode)

            fig = plot_tax_data(interpolated_data, year, mode, highlight_country) 

            img_bytes = fig.to_image(format="png", width=800, height=600)
            img = Image.open(io.BytesIO(img_bytes))

            gif_writer.append_data(np.array(img))

    gif_writer.close()

def interpolate_data(data, year, weight, mode):
    """
    Interpolate data between the given year and the next year based on the weight.
    Weight is between 0 and 1 where 0 represents the given year and 1 represents the next year.
    """
    interpolated_data = {}

    for country, country_data in data.items():
        interpolated_data[country] = {"OECD": country_data["OECD"], str(year): {}}
        for tax, value in country_data[str(year)].items():
            if str(year+1) in country_data:
                next_value = country_data[str(year+1)][tax]
                # Linear interpolation
                interpolated_value = (1 - weight) * value + weight * next_value
                interpolated_data[country][str(year)][tax] = interpolated_value


    return interpolated_data

# before interpolating!
def old_create_gif(data, mode):
    
    filename = f"tax_composition_{mode}.gif"
    # Set the animation frame duration
    frame_duration = 0.1  # in seconds

    gif_writer = imageio.get_writer(filename, mode='I', duration=frame_duration, loop=0)

    
    
    for year in range(start_year, end_year + 1):
        print(f"Generating {year}")
        fig = plot_tax_data(data, year, mode)

        img_bytes = fig.to_image(format="png", width=1024, height=768)
        img = Image.open(io.BytesIO(img_bytes))
        # Append the image to the GIF
        gif_writer.append_data(np.array(img))

    gif_writer.close()


def plot_tax_data(oecd_data, year, mode, highlight_country):
    year = str(year)
    
    # Filter countries based on the mode
    if mode == "OECD":
        filtered_data = {k: v for k, v in oecd_data.items() if v.get("OECD", False) == True}
    elif mode == "Non-OECD":
        filtered_data = {k: v for k, v in oecd_data.items() if v.get("OECD", False) == False}
    else:  # mode == "both"
        filtered_data = oecd_data
        
     # Sort countries by total_revenue for the specified year
    sorted_countries = sorted(filtered_data.keys(), key=lambda x: filtered_data[x].get(year, {}).get("total_revenue", 0), reverse=True)
    filtered_data = {country: filtered_data[country] for country in sorted_countries}

    # Extract tax data for the specified year
    countries = []
    tax_data = {"income_tax": [], "NI": [], "corporate": [], "property": [], "VAT": [], "non_VAT_sales": []}
    for country, data in filtered_data.items():
        if year in data:
            countries.append(country)
            for tax, values in tax_data.items():
                values.append(data[year].get(tax, 0))
    
    # Plot
    tax_colors = {
        "income_tax": "#002060",
        "NI": "#0070C0",
        "corporate": "#00B050",
        "property": "#FFC000",
        "VAT": "#E62F33",
        "non_VAT_sales": "#C00000"
    }
    
    logo = load_logo(logo_file)
    
    fig = go.Figure()
    for tax in tax_order:
        fig.add_trace(go.Bar(
            x=countries,
            y=tax_data[tax],
            name=tax_labels[tax],
            marker=dict(color=tax_colors[tax]),
        ))
        
    styled_ticklabels = [label if label != highlight_country else f'<span style="color:red; font-weight:bold">{label}</span>'  for label in countries]
    
    fig.update_layout(
        barmode='stack',
        images=[logo],
        title=f"{mode} tax system composition (% of GDP, {year})",
        title_x=0.5,
        title_font=dict(size=24),
        xaxis_tickangle=-45, 
        legend=dict(
            x=0.98,  # Set x position to the far right
            y=0.95,  # Set y position to the very top
            xanchor='right',  # Anchor the right side of the legend
            yanchor='top',  # Anchor the top of the legend
            font=dict(size=12) 
        ),
        xaxis=dict(
            ticktext=styled_ticklabels,  # Use the styled tick labels
            tickvals=list(range(len(countries)))  # Use the indices as tick values
        ),
        yaxis=dict(
            range=[0, 50],          # Set the range from 0 to 40
            dtick=5,                # Set the interval between ticks to 5
            title="% of GDP",       # Your y-axis title
            title_font=dict(size=20)
        ),
        margin=dict(t=60, r=20, b=50, l=50)
    )

    return fig

if os.path.exists(f"OECD_composition_totals.json"):
    print("Loading pre-generated data")
    oecd_data = load_from_json("OECD_composition_totals")
    
else:
    print("Loading in data from Excel file")
    raw_data = load_in_workbook(wb_file)
    print("Processing data")
    oecd_data = process_raw_data(raw_data)
    save_to_json("OECD_composition_totals", oecd_data)


# for a one off chart:
# fig = plot_tax_data(oecd_data, 2020, "OECD", "United Kingdom")
# fig.show()

create_gif(oecd_data, "Non-OECD", "United Kingdom")
