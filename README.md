# Accurex Distributor Order Report Card Generator

Generates PDF report cards for all distributors from the monthly Excel sales file.
Same fonts (DejaVu Sans with ₹ symbol) and KPI cards as before. The report now has a single, simplified table — **"Order to be Placed"** — comparing Last Year vs This Year qty/amount directly (no monthly averaging), sorted value-wise from highest to lowest order required. Rows needing an order are shown in red; rows with no action needed (already ordered) are shown in green.

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
2. Optionally upload the Accurex logo
3. Upload your Excel sales file
4. Select distributors to generate
5. Click **Generate PDFs**
6. Download individually or as a ZIP

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
