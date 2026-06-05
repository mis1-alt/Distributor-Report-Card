# Accurex Distributor Order Report Card Generator

Generates PDF report cards for all distributors from the monthly Excel sales file.
Output matches the finalized SNEH report format exactly — same fonts (DejaVu Sans with ₹ symbol), same layout, same KPI cards, same tables.

## Deploy to Streamlit Cloud (free, 5 minutes)

1. Create a free account at https://github.com and create a new repository
2. Upload these 4 files to the repository:
   - `app.py`
   - `requirements.txt`
   - `packages.txt`
   - `README.md`
3. Go to https://share.streamlit.io and sign in with GitHub
4. Click **New app** → select your repository → set main file to `app.py` → Deploy
5. Share the URL with your team — done!

## How to use

1. Open the app URL
2. Set **Months elapsed in FY** in the sidebar (April=1, May=2, June=3...)
3. Optionally upload the Accurex logo
4. Upload your Excel sales file
5. Select distributors to generate
6. Click **Generate PDFs**
7. Download individually or as a ZIP

## Excel format expected

Each sheet = one distributor
- Col A: LY Product name
- Col B: LY Quantity
- Col C: LY Amount
- Col D: (blank separator)
- Col E: CY Product name
- Col F: CY Quantity
- Col G: CY Amount
- Row 1: Headers (skipped)
- Row 2: Distributor totals (skipped)
- Row 3 onwards: Product data
