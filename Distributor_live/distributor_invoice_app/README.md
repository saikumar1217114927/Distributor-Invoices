# Distributor Invoice Generator

A small Flask web app that turns your distributor Excel sheet into one
Word (.docx) TAX INVOICE per distributor, automatically, with GST
calculated per your rules.

## Run locally

```bash
cd distributor_invoice_app
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000 in your browser.

## Put it online for your team (free) — Render

[Render](https://render.com) gives you a free, always-public URL with
no credit card. The free tier "sleeps" after ~15 minutes of no
traffic, so the first request after a break takes 30-60 seconds to
wake up — fine for an internal tool.

1. **Push this folder to a GitHub repo** (public or private both work).
2. Go to [render.com](https://render.com) → sign up (GitHub login is
   fastest) → **New +** → **Web Service** → connect your repo.
3. Render should auto-detect Python. Set:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
   - **Instance Type:** Free
4. (Recommended, since invoices contain bank account numbers) Add two
   environment variables under **Environment**:
   - `TEAM_PASSWORD` — a password your team will type in before using
     the app. Leave this unset if you want it open to anyone with the
     link.
   - `SECRET_KEY` — any random string (used to sign the login
     session cookie).
5. Click **Create Web Service**. In a couple of minutes you'll get a
   URL like `https://distributor-invoices.onrender.com` — share that
   with your team.

That's it — no server to manage. Every push to your GitHub repo
redeploys automatically.

### Alternative: PythonAnywhere

Also free, no credit card, and doesn't sleep — good if you'd rather
not use Git. Trade-off: limited daily CPU seconds on the free plan, so
it's best for occasional/light use (generating a batch of invoices a
few times a month is well within limits). Sign up at
[pythonanywhere.com](https://www.pythonanywhere.com), create a Flask
web app from the dashboard, upload this folder via their Files tab or
`git clone`, point the WSGI config at `app.py`'s `app` object, and
install `requirements.txt` in a virtualenv via their in-browser
console.

## How it works

1. Upload the distributor Excel file with columns: Partner Code,
   Partner Name, Total, GST, Address, PAN, Bank Name,
   Bank Account Number, ifsc code. (Address and PAN are optional —
   if a distributor's cell is blank, that line is simply left off
   their invoice.)
2. Pick the payout month and the invoice date using the built-in
   month/date pickers (they default to today) — these apply to every
   invoice in that batch. They're rendered on the invoice as
   `Apr-2023` and `30-Apr-2023` respectively.
3. Click **Generate invoices** — you get a ZIP with one `.docx` per
   distributor, named `<PartnerCode>_<Month>.docx`, plus a
   `summary.csv` recapping every distributor's base/GST/total.

## GST rule applied to every distributor row

The `Total` column in your sheet is treated as the **final,
GST-inclusive** payout amount. GST is extracted back out of it:

| GST No. on file            | Tax charged                  |
|-----------------------------|-------------------------------|
| Blank / `0`                 | None — invoice = Total as-is |
| Present, doesn't start "33" | IGST @ 18%                   |
| Present, starts with "33"   | SGST @ 9% + CGST @ 9%        |

```
base (Particulars amount) = round(Total / 1.18)
tax  (IGST, or SGST+CGST)  = round(base * rate)
```

This mirrors the sample invoice exactly (Total 8,386 → base 7,107,
IGST 1,279).

Myalternates' own details (name, address, GSTIN, place of supply, SAC
code) are fixed on every invoice, same as the sample.

## Layout / design details

- **Myalternates logo** in the top-right corner of every invoice
  (`assets/myalternates_logo.png` — swap this file to change it).
- **Tight, consistent spacing** throughout — no stray gaps between
  rows in the From/To blocks.
- **Address** wraps at roughly half the page width and continues on
  extra lines below if long, instead of stretching edge to edge.
- **Bank Details** laid out in a clean, borderless aligned column
  (label, colon, value) rather than tab-separated text.
- **Particulars table** has generously padded, taller rows for a more
  professional look.
- **Net Amount Payable to Distributor** and **Total Tax (GST)
  Included Above** summary strip directly below the particulars table.
- Every invoice fits on **a single page**.

## Files

- `app.py` — Flask routes, optional team-password gate, month/date
  formatting, upload → generate → zip
- `invoice_engine.py` — GST math, number-to-words, and the .docx builder
- `templates/index.html` — upload page with month/date pickers
- `templates/login.html` — password gate page (only used if
  `TEAM_PASSWORD` is set)
- `assets/myalternates_logo.png` — logo placed top-right on every invoice
- `requirements.txt`

## Customizing

- To change the fixed Myalternates details, SAC code, or GST rates,
  edit the constants at the top of `invoice_engine.py`.
- To accept a different set of Excel column names, edit
  `COLUMN_ALIASES` in `app.py`.
- To change how wide the address wraps, adjust the `right_indent` set
  on the address paragraph in `build_invoice_docx`.
- To resize/reposition the logo, edit `_add_logo_header` in
  `invoice_engine.py`.
- To turn off the password gate, just don't set `TEAM_PASSWORD` on
  your host.
