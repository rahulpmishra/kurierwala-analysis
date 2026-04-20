# Kurier Analysis

Lightweight Streamlit dashboard for analyzing courier booking and payment data from Google Sheets and Excel workbooks.

[![Kurier Analysis Demo](assets/demo.gif)](https://drive.google.com/file/d/16voHjFiuKkF64F52Zde-RJN2i_3wIYN6/view?usp=sharing)

## Why It Matters

This project turns a notebook-based logistics analysis workflow into a simple reusable app for month-wise operational reporting. It is designed for quick day-to-day use: load a workbook, pick a month, choose a report, and drill down into the numbers.

## Demo

- Quick preview GIF is shown at the top of this README
- Full demo video: [Watch the project demo](https://drive.google.com/file/d/16voHjFiuKkF64F52Zde-RJN2i_3wIYN6/view?usp=sharing)

## What It Does

- Accepts a Google Sheet URL or Excel upload
- Detects valid month-year sheets automatically
- Shows the spreadsheet title after analysis
- Saves the latest working source URLs for faster reuse
- Supports drill-down from date summaries to sender-level details

## Reports

- Date-wise packet count
- Packets booked per sender
- Packets booked per mode
- Payment received per month
  - Cash amount
  - UPI amount
  - Credit amount
  - Credit count
  - Transaction count
  - Sender-wise breakdown for a selected date

## Example Input Workbook

Use this sample file to understand the expected workbook structure and test the app quickly:

[Download sample workbook](examples/kurier_sample_input.xlsx)

Included sheets:

- `Instructions`
- `JAN 2026`
- `FEB 2026`
- `MARCH 2026`

Core columns used by the app:

| Column | Example | Used For |
| --- | --- | --- |
| `DATE` | `01-03-2026` | Date-wise analysis and sorting |
| `AWB NO.` | `MAR001` | Packet counting |
| `SENDER NAME` | `TRACKON` | Sender-wise summaries |
| `MODE` | `AIR` / `SURFACE` | Mode-wise packet analysis |
| `CREDIT OR CASH` | `CASH` / `UPI` / `CREDIT` | Payment split logic |
| `AMOUNT` | `140` / `monthly` | Payment totals and credit count |

## Tech Stack

- Python
- Streamlit
- Pandas
- OpenPyXL

## Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Project Files

- `app.py` - Streamlit application
- `data_analysis_kurierwala.ipynb` - original notebook used as the logic base
- `examples/kurier_sample_input.xlsx` - sample multi-sheet input workbook
- `requirements.txt` - project dependencies
