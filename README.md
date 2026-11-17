# Smart MIS

Upload a sales Excel/CSV file. Smart MIS cleans it and tells a small shop or online seller what to stock, fix and promote.

## Run it
```
pip install -r requirements.txt
streamlit run app.py
```
Then upload `demo_sales_full_messy.xlsx` (10,000 synthetic rows with duplicates, blanks, errors and typos).

## Files
| File | Purpose |
|---|---|
| `app.py` | The Streamlit screens |
| `mis_core.py` | Cleaning + analytics (no Streamlit, easy to test) |
| `demo_sales_full_messy.xlsx` | Synthetic demo data |

## Columns
Required: order date, order ID, product, quantity, unit selling price.
Optional (each unlocks a report): order status, category, state/city, channel, customer ID, cost price.

## What it checks and fixes
Blank cells, Excel errors (#N/A...), text dates (read day-first), impossible dates, negative/zero/text quantities and prices,
prices 10x above a product's median, exact duplicate rows (first copy kept), same order ID + product with different details,
case/space/typo variants of statuses and product names.
Cancelled, returned, refunded, RTO, failed orders are excluded from net sales.

## Reports
Action Plan, sales overview, product ranking (A/B/C, trend), returns & cancellations, category/state/channel,
repeat customers, profit & margin, and a full Excel MIS report.

## Limits
- Action Plan suggestions are rule-based hints, not guarantees.
- Profit ignores shipping, marketplace fees and ad cost.
- Text dates are read day-first (DD-MM-YYYY).
- Demo data is synthetic. Test with a real file before relying on it.
