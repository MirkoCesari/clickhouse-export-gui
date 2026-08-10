#!/usr/bin/env python3

import csv
import io
import os
import threading
import urllib.error
import urllib.parse
import urllib.request
import tkinter as tk

from tkinter import ttk, filedialog, messagebox


class ClickHouseExporter:
    def __init__(self, root):
        self.root = root
        self.root.title("ClickHouse Query Exporter")
        self.root.geometry("950x720")
        self.root.minsize(800, 600)

        self.create_ui()

    def create_ui(self):

        # ------------------------------------------------------------
        # Connection
        # ------------------------------------------------------------

        connection_frame = ttk.LabelFrame(
            self.root,
            text="ClickHouse Connection"
        )

        connection_frame.pack(
            fill="x",
            padx=10,
            pady=10
        )

        ttk.Label(
            connection_frame,
            text="Host:"
        ).grid(
            row=0,
            column=0,
            padx=5,
            pady=5,
            sticky="w"
        )

        self.host_var = tk.StringVar(value="localhost")

        ttk.Entry(
            connection_frame,
            textvariable=self.host_var,
            width=25
        ).grid(
            row=0,
            column=1,
            padx=5,
            pady=5
        )

        ttk.Label(
            connection_frame,
            text="Port:"
        ).grid(
            row=0,
            column=2,
            padx=5,
            pady=5
        )

        self.port_var = tk.StringVar(value="8123")

        ttk.Entry(
            connection_frame,
            textvariable=self.port_var,
            width=8
        ).grid(
            row=0,
            column=3,
            padx=5,
            pady=5
        )

        ttk.Label(
            connection_frame,
            text="Database:"
        ).grid(
            row=0,
            column=4,
            padx=5,
            pady=5
        )

        self.database_var = tk.StringVar(value="default")

        ttk.Entry(
            connection_frame,
            textvariable=self.database_var,
            width=18
        ).grid(
            row=0,
            column=5,
            padx=5,
            pady=5
        )

        ttk.Label(
            connection_frame,
            text="User:"
        ).grid(
            row=1,
            column=0,
            padx=5,
            pady=5,
            sticky="w"
        )

        self.user_var = tk.StringVar(value="default")

        ttk.Entry(
            connection_frame,
            textvariable=self.user_var,
            width=25
        ).grid(
            row=1,
            column=1,
            padx=5,
            pady=5
        )

        ttk.Label(
            connection_frame,
            text="Password:"
        ).grid(
            row=1,
            column=2,
            padx=5,
            pady=5
        )

        self.password_var = tk.StringVar()

        self.password_entry = ttk.Entry(
            connection_frame,
            textvariable=self.password_var,
            show="*",
            width=25
        )

        self.password_entry.grid(
            row=1,
            column=3,
            columnspan=2,
            padx=5,
            pady=5,
            sticky="w"
        )

        self.show_password_var = tk.BooleanVar()

        ttk.Checkbutton(
            connection_frame,
            text="Show",
            variable=self.show_password_var,
            command=self.toggle_password
        ).grid(
            row=1,
            column=5,
            padx=5
        )

        self.test_button = ttk.Button(
            connection_frame,
            text="Test Connection",
            command=self.test_connection
        )

        self.test_button.grid(
            row=2,
            column=0,
            columnspan=2,
            padx=5,
            pady=8,
            sticky="w"
        )

        # ------------------------------------------------------------
        # Query
        # ------------------------------------------------------------

        query_frame = ttk.LabelFrame(
            self.root,
            text="SQL Query"
        )

        query_frame.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=5
        )

        self.query_text = tk.Text(
            query_frame,
            wrap="none",
            font=("Consolas", 10)
        )

        query_scroll_y = ttk.Scrollbar(
            query_frame,
            orient="vertical",
            command=self.query_text.yview
        )

        query_scroll_x = ttk.Scrollbar(
            query_frame,
            orient="horizontal",
            command=self.query_text.xview
        )

        self.query_text.configure(
            yscrollcommand=query_scroll_y.set,
            xscrollcommand=query_scroll_x.set
        )

        self.query_text.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        query_scroll_y.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        query_scroll_x.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        query_frame.rowconfigure(0, weight=1)
        query_frame.columnconfigure(0, weight=1)

        self.query_text.insert(
            "1.0",
            "SELECT\n"
            "    now() AS current_time,\n"
            "    version() AS server_version"
        )

        # ------------------------------------------------------------
        # Output
        # ------------------------------------------------------------

        output_frame = ttk.LabelFrame(
            self.root,
            text="Output"
        )

        output_frame.pack(
            fill="x",
            padx=10,
            pady=10
        )

        ttk.Label(
            output_frame,
            text="Format:"
        ).grid(
            row=0,
            column=0,
            padx=5,
            pady=5
        )

        self.format_var = tk.StringVar(value="CSV")

        self.format_combo = ttk.Combobox(
            output_frame,
            textvariable=self.format_var,
            values=["CSV", "Excel"],
            state="readonly",
            width=10
        )

        self.format_combo.grid(
            row=0,
            column=1,
            padx=5,
            pady=5
        )

        self.export_button = ttk.Button(
            output_frame,
            text="Run Query and Save",
            command=self.start_export
        )

        self.export_button.grid(
            row=0,
            column=2,
            padx=15,
            pady=5
        )

        # ------------------------------------------------------------
        # Progress
        # ------------------------------------------------------------

        self.progress = ttk.Progressbar(
            self.root,
            mode="indeterminate"
        )

        self.progress.pack(
            fill="x",
            padx=10,
            pady=(0, 5)
        )

        self.status_var = tk.StringVar(
            value="Ready"
        )

        self.status_label = ttk.Label(
            self.root,
            textvariable=self.status_var
        )

        self.status_label.pack(
            fill="x",
            padx=10,
            pady=(0, 10)
        )

    # ------------------------------------------------------------
    # Password
    # ------------------------------------------------------------

    def toggle_password(self):

        if self.show_password_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")

    # ------------------------------------------------------------
    # HTTP connection
    # ------------------------------------------------------------

    def execute_query(
        self,
        query,
        output_format="CSVWithNames"
    ):

        host = self.host_var.get().strip()
        port = self.port_var.get().strip()
        database = self.database_var.get().strip()
        user = self.user_var.get().strip()
        password = self.password_var.get()

        params = {
            "database": database,
            "default_format": output_format
        }

        url = (
            f"http://{host}:{port}/?"
            + urllib.parse.urlencode(params)
        )

        request = urllib.request.Request(
            url,
            data=query.encode("utf-8"),
            method="POST"
        )

        if user:
            request.add_header(
                "X-ClickHouse-User",
                user
            )

        if password:
            request.add_header(
                "X-ClickHouse-Key",
                password
            )

        try:
            return urllib.request.urlopen(
                request,
                timeout=60
            )

        except urllib.error.HTTPError as e:

            error = e.read().decode(
                "utf-8",
                errors="replace"
            )

            raise Exception(error)

        except urllib.error.URLError as e:
            raise Exception(
                f"Connection error:\n{e}"
            )

    # ------------------------------------------------------------
    # Test connection
    # ------------------------------------------------------------

    def test_connection(self):

        self.status_var.set(
            "Testing connection..."
        )

        self.test_button.config(
            state="disabled"
        )

        self.progress.start()

        threading.Thread(
            target=self.test_connection_thread,
            daemon=True
        ).start()

    def test_connection_thread(self):

        try:

            response = self.execute_query(
                "SELECT version()"
            )

            result = response.read().decode(
                "utf-8"
            ).strip()

            # CSVWithNames returns:
            # version()
            # 26.x.x

            lines = result.splitlines()

            version = (
                lines[-1]
                if lines
                else "Unknown"
            )

            self.root.after(
                0,
                lambda: self.connection_success(version)
            )

        except Exception as e:

            self.root.after(
                0,
                lambda: self.connection_error(str(e))
            )

    def connection_success(self, version):

        self.progress.stop()

        self.test_button.config(
            state="normal"
        )

        self.status_var.set(
            f"Connected - ClickHouse {version}"
        )

        messagebox.showinfo(
            "Connection successful",
            f"Connected successfully.\n\n"
            f"ClickHouse version:\n{version}"
        )

    def connection_error(self, error):

        self.progress.stop()

        self.test_button.config(
            state="normal"
        )

        self.status_var.set(
            "Connection failed"
        )

        messagebox.showerror(
            "Connection error",
            error
        )

    # ------------------------------------------------------------
    # Export
    # ------------------------------------------------------------

    def start_export(self):

        query = self.query_text.get(
            "1.0",
            "end"
        ).strip()

        if not query:

            messagebox.showwarning(
                "Query missing",
                "Insert a SQL query."
            )

            return

        output_format = self.format_var.get()

        if output_format == "CSV":

            filename = filedialog.asksaveasfilename(
                title="Save CSV",
                defaultextension=".csv",
                filetypes=[
                    ("CSV files", "*.csv"),
                    ("All files", "*.*")
                ]
            )

        else:

            filename = filedialog.asksaveasfilename(
                title="Save Excel",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel files", "*.xlsx"),
                    ("All files", "*.*")
                ]
            )

        if not filename:
            return

        self.export_button.config(
            state="disabled"
        )

        self.test_button.config(
            state="disabled"
        )

        self.progress.start()

        self.status_var.set(
            "Executing query..."
        )

        threading.Thread(
            target=self.export_thread,
            args=(
                query,
                filename,
                output_format
            ),
            daemon=True
        ).start()

    def export_thread(
        self,
        query,
        filename,
        output_format
    ):

        try:

            response = self.execute_query(
                query,
                "CSVWithNames"
            )

            if output_format == "CSV":

                self.export_csv(
                    response,
                    filename
                )

            else:

                self.export_excel(
                    response,
                    filename
                )

            self.root.after(
                0,
                lambda: self.export_success(filename)
            )

        except Exception as e:

            self.root.after(
                0,
                lambda: self.export_error(str(e))
            )

    # ------------------------------------------------------------
    # CSV
    # ------------------------------------------------------------

    def export_csv(
        self,
        response,
        filename
    ):

        total_bytes = 0

        with open(filename, "wb") as file:

            while True:

                chunk = response.read(
                    1024 * 1024
                )

                if not chunk:
                    break

                file.write(chunk)

                total_bytes += len(chunk)

                mb = total_bytes / 1024 / 1024

                self.root.after(
                    0,
                    lambda value=mb:
                    self.status_var.set(
                        f"Downloading: "
                        f"{value:.1f} MB"
                    )
                )

    # ------------------------------------------------------------
    # Excel
    # ------------------------------------------------------------

    def export_excel(
        self,
        response,
        filename
    ):

        try:
            from openpyxl import Workbook

        except ImportError:

            raise Exception(
                "Excel export requires openpyxl.\n\n"
                "Install it with:\n"
                "pip install openpyxl"
            )

        workbook = Workbook(
            write_only=True
        )

        worksheet = workbook.create_sheet(
            "Query Result"
        )

        text_stream = io.TextIOWrapper(
            response,
            encoding="utf-8",
            newline=""
        )

        reader = csv.reader(
            text_stream
        )

        row_count = 0

        # Excel maximum rows per worksheet
        excel_row_limit = 1_048_576

        sheet_number = 1

        for row in reader:

            if row_count >= excel_row_limit:

                sheet_number += 1

                worksheet = workbook.create_sheet(
                    f"Query Result {sheet_number}"
                )

                row_count = 0

            worksheet.append(row)

            row_count += 1

            if row_count % 10000 == 0:

                self.root.after(
                    0,
                    lambda value=row_count:
                    self.status_var.set(
                        f"Writing Excel: "
                        f"{value:,} rows"
                    )
                )

        workbook.save(filename)

    # ------------------------------------------------------------
    # Results
    # ------------------------------------------------------------

    def export_success(
        self,
        filename
    ):

        self.progress.stop()

        self.export_button.config(
            state="normal"
        )

        self.test_button.config(
            state="normal"
        )

        self.status_var.set(
            f"Export completed: {filename}"
        )

        messagebox.showinfo(
            "Export completed",
            f"File successfully created:\n\n"
            f"{filename}"
        )

    def export_error(
        self,
        error
    ):

        self.progress.stop()

        self.export_button.config(
            state="normal"
        )

        self.test_button.config(
            state="normal"
        )

        self.status_var.set(
            "Export failed"
        )

        messagebox.showerror(
            "Error",
            error
        )


def main():

    root = tk.Tk()

    app = ClickHouseExporter(root)

    root.mainloop()


if __name__ == "__main__":
    main()