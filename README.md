# ClickHouse Export GUI

A small Tkinter GUI to run a SQL query against ClickHouse and save the result as CSV or Excel.

## Features

- Connection settings (host, port, database, user, password) with a test button
- Free-form SQL editor
- Export to CSV (streamed, no size limit) or `.xlsx` (auto-splits over 1,048,576 rows)
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

## License

MIT

---

Not affiliated with ClickHouse
