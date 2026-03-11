# No Strings Attached — Analytics Dashboard

A static analytics dashboard for **No Strings Attached** (BH 24 Salt Lake City Sector 2), built from monthly Petpooja exports.

## Files

| File | Purpose |
|---|---|
| `dashboard.html` | Single-file dashboard — all charts, tables, and styling |
| `process_data.py` | Reads Excel exports and generates `dashboard_data.json` |
| `dashboard_data.json` | Generated data file loaded by the dashboard |

## Requirements

```bash
pip3 install openpyxl
```

---

## Monthly Refresh (end of each month)

### Step 1 — Export from Petpooja

Export two reports from the Petpooja backend covering the **last 12 months**:

1. **Order Listing report**
   - Reports → Order Reports → Order Listing
   - File will be named: `Order_Listing_*.xlsx`

2. **Item Wise Sales report**
   - Reports → Item Reports → Item Wise Sales Report
   - File will be named: `Item_Wise_Sales_Report_*.xlsx`

Drop both files into this project folder.

### Step 2 — Regenerate the data

```bash
python3 process_data.py
```

This auto-detects both files and writes `dashboard_data.json`.

### Step 3 — Preview locally (optional)

```bash
python3 -m http.server 8080
```

Open [http://localhost:8080](http://localhost:8080) to verify everything looks correct.

### Step 4 — Publish

```bash
git add dashboard_data.json
git commit -m "Data refresh: <Month Year>"
git push
```

The live dashboard updates automatically within a minute.

---

## Dashboard Charts

1. **Monthly Revenue** — bar + line (revenue + order count per month). Click any bar to drill into daily breakdown and top items for that month.
2. **Revenue Distribution by Day of Week** — box plot showing Q1/median/Q3/whiskers/mean per day.
3. **Sales by Time of Day** — 4-hour windows; noon–midnight highlighted.
4. **Top Selling Items (by order appearances)** — sortable table with monthly sparklines and avg order value.
5. **Top Items by Revenue** — horizontal bar chart from item-wise sales data.
6. **Top Items by Quantity Sold** — horizontal bar chart from item-wise sales data.
7. **Top Food Items by Revenue** — same as above with drinks excluded.
8. **Top Food Items by Quantity Sold** — same as above with drinks excluded.

## Notes

- The dashboard must be served over HTTP (`python3 -m http.server`) — opening `dashboard.html` directly as a file won't work due to `fetch()`.
- `Order_Listing_*.xlsx` and `Item_Wise_Sales_Report_*.xlsx` are excluded from git (business data). Only `dashboard_data.json` is committed.
- If you accumulate multiple monthly order listing files, `process_data.py` merges them automatically and deduplicates orders.
- Item counts in the Top Selling Items table reflect **order appearances**, not units sold. Use the Item Wise Sales charts for actual unit counts.
