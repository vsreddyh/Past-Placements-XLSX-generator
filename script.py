import json
import re
import math
import requests
from lxml import html
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from gspread_formatting import (
    CellFormat,
    format_cell_range,
    get_conditional_format_rules,
    ConditionalFormatRule,
    BooleanCondition,
    BooleanRule,
    Color,
)
from google.oauth2.service_account import Credentials


# ---------------- AUTH ---------------- #
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

creds = Credentials.from_service_account_file("service_account.json", scopes=scopes)
gc = gspread.authorize(creds)


# ---------------- LOAD JSON ---------------- #
# Read from local file
url = "https://kmit.in/placements/placement.php"
headers = {"User-Agent": "Mozilla/5.0"}

response = requests.get(url, headers=headers)
tree = html.fromstring(response.content)

# Extract years (e.g., "2024-2025")
year_elements = tree.xpath("//div[@id='campus']/div/ul/li/a/b/text()")
years = [y.strip() for y in year_elements if y.strip()]
data = {}


# Loop through each year's corresponding div
globalset = set()
for i, year in enumerate(years, start=1):
    table_xpath = f"//div[@id='cp{year}']//table/tbody/tr"
    rows = tree.xpath(table_xpath)

    data[year] = []
    namedict = {}
    for row in rows:
        cols = [c.text_content().strip() for c in row.xpath("./td")]
        if len(cols) < 5:
            continue

        # Example row: [ "1", "Microsoft", "1", "125000", "54" ]
        _, name, selections, internship, ctc = cols
        if name in namedict:
            globalset.add(name)
            namedict[name] = namedict[name] + 1
            name = name + "-" + str(namedict[name])
        else:
            namedict[name] = 1
            name = name + "-1"
        data[year].append(
            {
                "name": name,
                "selections": int(selections) if selections.isdigit() else selections,
                "internship": internship,
                "ctc": ctc,
            }
        )
print("JSON Read")


companies = {}
for year, companies_list in data.items():
    for company in companies_list:
        name = company["name"]
        x = name.split("-", 1)[0]
        if x not in globalset:
            name = x
        if name not in companies:
            companies[name] = {}
        companies[name][year] = {
            "selections": company.get("selections", None),
            "internship": company.get("internship", None),
            "ctc": company.get("ctc", None),
        }
print("Structured Headers")

years = sorted(data.keys(), reverse=True)
print("Sorted years")

header_parent = ["Academic Year"]
header_child = ["Company"]
df_child = ["Company"]

for year in years:
    header_parent.extend([year, None, None])
    header_child.extend(["Selections", "Stipend", "CTC"])
    df_child.extend(["Selections", "Stipend", year + "_CTC"])

rows = []
for company, year_data in companies.items():
    row = []
    row.append(company)
    for year in years:
        if year in year_data:
            row.append(year_data[year].get("selections"))
            row.append(year_data[year].get("internship"))
            ctc = year_data[year].get("ctc")
            match = re.search(r"[-+]?\d*\.\d+|\d+", str(ctc))
            if match:
                num_str = match.group()
                if ctc is not None and math.isclose(float(num_str), 1.22, abs_tol=1e-6):
                    row.append(str(122))
                    continue
                row.append(str(float(num_str)))
            else:
                row.append(str(0))
        else:
            row.extend([None, None, None])
    rows.append(row)

df = pd.DataFrame(rows, columns=df_child)
print("Dataframe created")

ctc_cols = [(col) for col in df.columns if col.endswith("CTC")]
for col in ctc_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce")
df_sorted = df.sort_values(by=ctc_cols, ascending=[False] * len(ctc_cols))
print("Data sorted")


# ---------------- CREATE SHEET ---------------- #
sh = gc.open_by_key("1o-05XFY0tgZU9MY2mHT4rtp3J0lwZFcQDI1FBefj5aY")
worksheet = sh.get_worksheet(0)
worksheet.clear()
worksheet.update(range_name="A1", values=[header_parent])
worksheet.update(range_name="A2", values=[header_child])
set_with_dataframe(worksheet, df_sorted, row=3, col=1, include_column_header=False)
print("Data Uploaded")


# ---------------- FORMAT SHEET ---------------- #
# Merge and center align parent year headers
col = 2
for year in years:
    start_cell = gspread.utils.rowcol_to_a1(1, col)
    end_cell = gspread.utils.rowcol_to_a1(1, col + 2)
    rng = f"{start_cell}:{end_cell}"
    worksheet.merge_cells(rng, merge_type="MERGE_ALL")
    col += 3
print("Years merged")

num_rows = len(df)
data_col_count = len(df.columns)

start_cell = gspread.utils.rowcol_to_a1(1, 2)
end_cell = gspread.utils.rowcol_to_a1(num_rows, data_col_count)
rng = f"{start_cell}:{end_cell}"
format_cell_range(worksheet, rng, CellFormat(horizontalAlignment="CENTER"))
print("Center aligned")

# Conditional formatting for blanks
colour = Color(252 / 255, 232 / 255, 178 / 255)
boolean_rule = BooleanRule(
    condition=BooleanCondition("BLANK"),
    format=CellFormat(backgroundColor=colour),
)
cf_rule = ConditionalFormatRule(
    ranges=[
        {
            "sheetId": worksheet._properties["sheetId"],
            "startRowIndex": 2,  # Row 3
            "startColumnIndex": 1,
            "endRowIndex": num_rows + 2,
            "endColumnIndex": data_col_count,
        }
    ],
    booleanRule=boolean_rule,
)

rules = get_conditional_format_rules(worksheet)
rules.append(cf_rule)
rules.save()
print("Conditional formatting rule added")
print("Completed")
