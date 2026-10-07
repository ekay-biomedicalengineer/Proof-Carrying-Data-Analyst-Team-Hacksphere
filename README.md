# Proof-Carrying Data Analyst

A local desktop app that answers questions about messy business tables, gives **re-runnable code for every number**, and **refuses** when the data cannot support a reliable answer.

- No external AI, no internet, no API key. Your data never leaves your computer.
- Works on CSV and Excel files (each Excel sheet is treated as its own table).
- Rule-based on purpose, so it is deterministic and auditable.

---

## Requirements

- Python 3.9 or newer (from python.org; tick **"Add python.exe to PATH"** when installing)
- Libraries: `pandas` and `openpyxl`
- Tkinter (the window toolkit). It is included with the python.org installer on Windows and macOS.

## Setup

1. Put `app.py` in a folder, for example `hackathon`.
2. Open that folder in VS Code (File -> Open Folder).
3. Open the terminal (Ctrl + `) and install the libraries:

```
python -m pip install pandas openpyxl
```

## Run

```
python app.py
```

If `python` is not recognised, use `py app.py`.

## How to use it

1. **Upload your tables.** Click **Add files...** and select your CSV or Excel files. An optional `.txt` notes file is also read.
2. **Check the detected columns.** The app guesses which table is orders, customers and refunds, and which columns are ids, dates, amounts, currencies and countries. Change a dropdown only if a guess is wrong.
3. **Choose a question** from the menu and set any extra choice (currency, country, customer, month and year).
4. Click **Get answer**.

Other buttons:

- **Copy code** copies the proof code so you can paste it into `check.py` and run it.
- **Save report...** saves every answer from the session to `report.md`.
- Click the header of a step card to fold or unfold it. Steps 1 and 2 fold automatically after an answer to give the output more room.

## The 12 questions

| # | Question | Extra choice |
|---|---|---|
| 1 | How many unique orders are there? | - |
| 2 | What is the total sales in a chosen currency? | Currency |
| 3 | What is the average order amount in a chosen currency? | Currency |
| 4 | Which customer has the most orders? | - |
| 5 | How many refunds are in the data? | - |
| 6 | How many orders were shipped to a chosen country? | Country |
| 7 | How many orders were placed in a chosen month and year? | Month, year |
| 8 | What was the profit? | - |
| 9 | What is the total revenue in one number? | - |
| 10 | What is the total amount refunded? | - |
| 11 | What country is a chosen customer in? | Customer ID |
| 12 | What is the average order amount expressed in a chosen currency (converted)? | Currency |

## How it works

1. **Load.** Reads CSV and Excel, drops blank rows and columns, trims spaces, and turns text amounts such as `$1,200.50` into numbers.
2. **Classify.** Decides which table is orders, customers and refunds, and matches columns to roles.
3. **Check the traps.** Tests for duplicates, missing values, mixed currencies, ambiguous dates, conflicting tables, and missing concepts such as cost or exchange rates.
4. **Decide.** Chooses one of three statuses.
5. **Prove.** Writes standalone pandas code and runs it twice in fresh Python processes. The answer is marked **verified** only if both runs agree.

### Statuses

| Status | Meaning |
|---|---|
| **ANSWERED** (green) | The data supports exactly one reliable value. |
| **AMBIGUOUS** (amber) | More than one valid reading, mixed units, or sources that conflict. |
| **CANNOT DETERMINE** (red) | Required data is missing or too incomplete. The app says exactly what. |

### Traps it detects

- **Duplicate rows:** exact duplicates are removed and counted. The same id with different content is flagged.
- **Mixed currencies:** never added or converted without an exchange rate in the data. Outside rates are never used.
- **Missing values:** never treated as zero.
- **Ambiguous dates:** slash dates like `08/07/2025` (day/month or month/day) produce a range instead of one number. ISO dates (`2025-07-03`) and slash dates with a number above 12 (`14/09/2025`) are unambiguous.
- **Conflicting tables:** for example, a customer's account country differs from the country their orders shipped to.
- **Missing concepts:** profit is refused when no cost data exists.
- **Hidden units:** refund amounts have no currency column, so the app looks up the currency of each refunded order.

### Output sections

| Section | What it shows |
|---|---|
| Question line | The question and the choice made |
| Status badge | Coloured status and the verification note |
| Answer | The result in large text |
| WHY | The reasoning for the status |
| MISSING / AMBIGUOUS | What is absent, conflicting or open to several readings |
| EVIDENCE | Row counts, duplicates removed, ids, code output |
| ASSUMPTIONS | Choices the program made |
| RE-RUNNABLE CODE | A standalone script whose last line prints the number |

A section is shown only when it has content. A refusal has no code because there is no number to prove.

## Sample data

The `test_data` folder contains a small shop dataset with deliberate traps:

| File | Contents | Traps |
|---|---|---|
| `orders.csv` | 13 rows, 11 real orders | Duplicate rows (B101, B105), USD and GBP mixed, missing amount (B106), ambiguous date (B103 = `08/07/2025`) |
| `clients.csv` | 6 customers | K03 is India on the account but both orders shipped to Canada |
| `returns.csv` | 4 refunds | Missing refund amount (RF3), no currency column |
| `readme.txt` | Notes from the data team | No cost data and no exchange rate anywhere |

### Expected results on the sample data

| Question | Choice | Result |
|---|---|---|
| 1. Unique orders | - | ANSWERED: 11 |
| 2. Total sales | USD | ANSWERED: 1,660.50 |
| 3. Average order | GBP | CANNOT DETERMINE (B106 has no amount) |
| 4. Most orders | - | ANSWERED: K01 with 3 |
| 5. Refund count | - | ANSWERED: 4 |
| 6. Shipped to | USA | ANSWERED: 4 |
| 7. Orders in month | July 2025 | AMBIGUOUS: between 2 and 3 |
| 8. Profit | - | CANNOT DETERMINE (no cost data) |
| 9. Total revenue | - | AMBIGUOUS (USD and GBP, no rate) |
| 10. Total refunded | - | AMBIGUOUS (mixed currencies, RF3 missing) |
| 11. Customer country | K03 | AMBIGUOUS (India vs Canada) |
| 12. Average converted | USD | CANNOT DETERMINE (no GBP to USD rate) |

## Expected table layout

The app expects tables shaped like these, with any column names. It matches them by keywords such as `order`, `amount`, `currency`, `date`, `ship`, `country`, `customer` or `client`, and `refund`.

- **Orders table:** order id, customer id, order date, amount, currency, ship country
- **Customers table:** customer id, country (a name column is optional)
- **Refunds table:** refund id, order id, refund amount

If a needed table or column is missing, the app refuses the question and names exactly what is missing.

## Limits

- It supports the 12 question types above. A new type of question needs a new rule in `app.py`.
- Exchange-rate tables are detected but not applied. The app says so and does not guess a conversion.
- Old `.xls` Excel files may need `pip install xlrd`.
- The proof code contains file paths from the computer it was created on. To run it elsewhere, edit the paths or keep the files in the same location.
- Each table needs a header row. Title rows above the header can confuse column detection; use the dropdowns to correct it.

## Troubleshooting

| Problem | Fix |
|---|---|
| `Python was not found` | Install Python from python.org, tick "Add python.exe to PATH", restart VS Code. Optionally switch off the Store aliases under Settings -> Apps -> Advanced app settings -> App execution aliases. |
| `No module named 'pandas'` | Run `python -m pip install pandas openpyxl`. |
| `No module named 'tkinter'` | Reinstall Python from python.org and keep "tcl/tk and IDLE" ticked. |
| Wrong column detected | Change the dropdown in step 2. |
| Question refused as "column is missing" | Choose the right table or column in step 2, or upload the missing table. |

## Project files

```
hackathon/
├── app.py          the whole application
├── README.md       this file
└── test_data/
    ├── orders.csv
    ├── clients.csv
    ├── returns.csv
    └── readme.txt
```

