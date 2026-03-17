# KPI Report

A Streamlit dashboard for visualizing and analyzing KPI performance data against 45th and 60th percentile benchmarks.

---

## Features

- **File upload** — drag and drop a CSV or XLSX file to load your data
- **Sidebar filters** — filter the entire dashboard by Title, KPI type, and Beat/Announcement
- **Summary cards** — at-a-glance counts of rows above target, on target, and below target; click any card to filter the data table to just those rows
- **Data table** — color-coded rows with a status column per actuals column, a live search bar, and left-aligned columns for readability
- **Trend analysis chart** — cumulative line chart of actuals vs. the 45th–60th percentile target range, aggregated by Beat

---

## Color Coding

| Color | Meaning | Threshold |
|-------|---------|-----------|
| 🟢 Green | Above target | Actual >= 60th percentile |
| 🟠 Orange | On target | Actual >= 45th percentile and < 60th percentile |
| 🔴 Red | Below target | Actual < 45th percentile |

---

## Expected Data Format

Your CSV or XLSX file should contain the following columns (names are matched case-insensitively):

| Column | Description |
|--------|-------------|
| `Title` | Title or grouping label (optional, used for sidebar filter) |
| `KPI` | KPI type name (optional, used for sidebar filter and chart) |
| `Beat` | Announcement or beat label (optional, used for sidebar filter and chart x-axis) |
| `*actual*` | One or more columns containing the word "actual" — these are the values being evaluated |
| `*45*` or `*45th*` | Column containing the 45th percentile benchmark values |
| `*60*` or `*60th*` | Column containing the 60th percentile benchmark values |

Numeric values can be formatted as plain numbers (`1234`), comma-separated (`1,234`), or percentages (`60%`). Percentages are automatically normalized to decimals (`0.60`).

---

## Project Structure

```
your-project/
├── kpi_report.py        # Main Streamlit app
├── requirements.txt     # Python dependencies
├── README.md            # This file
└── .streamlit/
    └── config.toml      # Forces dark mode for all users
```

---

## Setup & Running Locally

**1. Clone or download the project files**

**2. Install dependencies**

```bash
pip install -r requirements.txt
```

**3. Run the app**

```bash
streamlit run kpi_report.py
```

The app will open at `http://localhost:8501` in your browser.

---

## Deploying to Streamlit Community Cloud

1. Push the project to a GitHub repository (include all files, including `.streamlit/config.toml`)
2. Go to [share.streamlit.io](https://share.streamlit.io) and connect your repo
3. Set the main file path to `kpi_report.py`
4. Click **Deploy**

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `streamlit` | App framework and UI |
| `pandas` | Data loading and manipulation |
| `numpy` | Numeric helpers |
| `plotly` | Interactive trend chart |
| `openpyxl` | Excel (.xlsx) file reading |