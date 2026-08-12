# ClickHouse Export GUI

A small Tkinter GUI to run one or more SQL queries against ClickHouse and save the results as Excel or CSV.

## Features

- Connection settings (host, port, database, user, password) with a test button
- Free-form SQL editor accepting **several queries separated by `;`**, executed one after another
- Export to `.xlsx` with **one sheet per query**, or to CSV (one file per query)
- Optional sheet naming through a `-- tab: My name` comment above each query
- Streamed writing, so results are never fully loaded in memory
- Auto-splits a result over 1,048,576 rows onto an extra sheet that repeats the header
- Runs in a background thread, so the UI stays responsive

## Requirements

- Python 3 (Tkinter included)
- `openpyxl` -> only for Excel export: `pip install openpyxl`

## Usage

```bash
python clickhouse_export_gui.pyw
```

On Windows you can just double-click the file to open it without a console window.

Connects over the ClickHouse HTTP interface (default port `8123`).

### Multiple queries

Separate the queries with `;`. Semicolons inside string literals, quoted
identifiers and comments are not treated as separators.

```sql
-- tab: Daily totals
SELECT
    toDate(event_time) AS day,
    count() AS events
FROM events
GROUP BY day
ORDER BY day;

-- tab: Top users
SELECT
    user_id,
    count() AS events
FROM events
GROUP BY user_id
ORDER BY events DESC
LIMIT 100
```

This produces a workbook with the sheets `Daily totals` and `Top users`.
Without a `-- tab:` comment a query gets the sheet name `Query 1`,
`Query 2`, and so on. Duplicate names get a `(2)` suffix, and characters
Excel forbids are replaced.

With the CSV format each query is written to its own file, named
`<chosen name>_01_<sheet name>.csv`, `<chosen name>_02_<sheet name>.csv`, …
A single query keeps exactly the file name you chose.

If a query fails, the run stops there and the file is still saved with the
results collected so far; the dialog reports which query failed and what
was saved.

## License

MIT

---

Not affiliated with ClickHouse
