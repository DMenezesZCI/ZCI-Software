#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Aug 21 08:13:42 2025

@author: davidmenezes
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Procore Timecard Downloader + Excel pasting
- Gets access token from refresh token
- Downloads timecard entries for a date range
- Removes "Cost Code Long Name" (if present), keeps columns A-K
- Writes rows (row 2..end, cols A..K) for each employee into the sheet named for that employee
  in an existing Excel workbook (preserves cell formatting by only writing values).
"""

import os
import json
import re
import io
import csv
import webbrowser
import requests
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from urllib.parse import urlparse, parse_qs
import traceback
from pathlib import Path
from http.server import BaseHTTPRequestHandler, HTTPServer
import threading

# -------------------- CONFIG --------------------
CLIENT_ID     = 'yHTFNlKjEMAjFcnEbGNIXRN13Bq6ezt6YJVM1YLeQyY'
CLIENT_SECRET = 'bsKFfy4jbFkq9avooml9uFq2tlL97VfMGFhaaFej9l8'
# IMPORTANT: This must EXACTLY match one of the Redirect URIs registered for your app
REFRESH_TOKEN = None   # e.g. '5dR...'; set to None to force interactive (or let saved token be used)
# in the Procore Developer Portal (My Apps -> Edit App -> Redirect URIs).
# It can be a dummy like "https://example.com/callback" — you will copy the code from the browser.
REDIRECT_URI  = "http://localhost:8080/callback"
COMPANY_ID    = '9993'
TOKEN_FILE    = "procore_token.json"
DOWNLOADS = Path.home() / "Downloads"
DOWNLOADS.mkdir(parents=True, exist_ok=True)

OUT_CSV = DOWNLOADS / "timecards.csv"
# ------------------------------------------------

# ------------------ Utilities -------------------
def log_error_to_file(error: Exception):
    """
    Saves error details to a timestamped log file in ~/Documents/ErrorLogs/
    """
    docs = Path.home() / "Documents" / "ErrorLogs"
    docs.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file = docs / f"error_{timestamp}.txt"

    with open(log_file, "w", encoding="utf-8") as f:
        f.write("=== Procore Timecard Downloader Error ===\n")
        f.write(f"Timestamp: {timestamp}\n\n")
        f.write("Error Message:\n")
        f.write(str(error) + "\n\n")
        f.write("Traceback:\n")
        f.write(traceback.format_exc())

    return log_file

def save_tokens(tokens):
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f)
    print(f"Saved tokens to {TOKEN_FILE}")

def load_tokens():
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r") as f:
            return json.load(f)
    return None

def sanitize_sheet_name(name: str) -> str:
    if not isinstance(name, str):
        name = str(name)
    name = re.sub(r'[:\\/?\*\[\]]+', ' ', name).strip()
    return name[:31] if len(name) > 31 else (name or "Unknown")
# ------------------------------------------------

# --------------- OAuth helpers ------------------
def exchange_refresh_for_access(refresh_token):
    url = "https://login.procore.com/oauth/token"
    data = {
        "grant_type": "refresh_token",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": refresh_token
    }
    return requests.post(url, data=data)

def exchange_code_for_tokens(code):
    url = "https://login.procore.com/oauth/token"
    data = {
        "grant_type": "authorization_code",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "redirect_uri": REDIRECT_URI
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()

def interactive_get_new_tokens():
    class OAuthHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            parsed = urlparse(self.path)
            qs = parse_qs(parsed.query)
            self.server.auth_code = qs.get("code", [None])[0]

            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b"<html><body><h2>Authorization successful. You can close this window.</h2></body></html>")

        def log_message(self, format, *args):
            return  # suppress logging to stdout

    auth_url = (
        "https://login.procore.com/oauth/authorize"
        f"?response_type=code&client_id={CLIENT_ID}&redirect_uri={REDIRECT_URI}"
    )
    webbrowser.open(auth_url)

    # Start local server in another thread
    server = HTTPServer(('localhost', 8080), OAuthHandler)
    thread = threading.Thread(target=server.handle_request)
    thread.start()
    thread.join()

    code = server.auth_code
    if not code:
        raise SystemExit("No authorization code received. Try again.")

    tokens = exchange_code_for_tokens(code)
    save_tokens(tokens)
    return tokens

def get_access_token():
    # priority: REFRESH_TOKEN var -> saved token file -> interactive flow
    if REFRESH_TOKEN:
        print("Using REFRESH_TOKEN from script variable to get access token...")
        r = exchange_refresh_for_access(REFRESH_TOKEN)
        if r.status_code == 200:
            tokens = r.json()
            save_tokens(tokens)
            return tokens["access_token"]
        else:
            print("Refresh-token-in-script failed:", r.status_code, r.text)

    tokens = load_tokens()
    if tokens and "refresh_token" in tokens:
        print("Attempting to refresh access token using saved refresh_token...")
        r = exchange_refresh_for_access(tokens["refresh_token"])
        if r.status_code == 200:
            new_tokens = r.json()
            save_tokens(new_tokens)
            return new_tokens["access_token"]
        else:
            print("Saved refresh failed:", r.status_code, r.text)

    print("No valid refresh token available — starting manual auth...")
    tokens = interactive_get_new_tokens()
    return tokens["access_token"]
# ------------------------------------------------

# ---------------- API calls (company-level) -----------
def try_timecards_endpoint(access_token, start_date, end_date):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json"
    }
    url = f"https://api.procore.com/rest/v1.0/companies/{COMPANY_ID}/timecard_entries"
    params = {
        "company_id": COMPANY_ID,
        "start_date": start_date,
        "end_date": end_date,
        "include_totals": "true"
    }
    r = requests.get(url, headers=headers, params=params)
    print(f"API Response: {r.status_code} - {r.text}")  # More logging
    if r.status_code == 200:
        try:
            data = r.json()
            if not data.get('timecards'):
                print("No timecards found for this date range.")
                return None
            df = pd.json_normalize(data['timecards'])
            return df
        except Exception as e:
            log_error_to_file(e)
            return None
    else:
        print(f"/timecards returned {r.status_code}: {r.text}")
        log_error_to_file(f"API error: {r.text}")
        return None



def try_timecard_entries_endpoint(access_token, start_date, end_date):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json"
    }
    url = f"https://api.procore.com/rest/v1.0/companies/{COMPANY_ID}/timecard_entries"
    params = {
        "company_id": COMPANY_ID,
        "start_date": start_date,
        "end_date": end_date,
        "include_totals": "true"
    }
    r = requests.get(url, headers=headers, params=params)
    print(f"API Response: {r.status_code} - {r.text}")  # Add more details here
    if r.status_code == 200:
        try:
            data = r.json()
            df = pd.json_normalize(data)
            print(f"Timecard Data: {data}")  # Inspect the data returned
            return df
        except Exception as e:
            print(f"Error parsing JSON: {str(e)}")
            with open(OUT_CSV_RAW, "wb") as f:
                f.write(r.content)
            try:
                return pd.read_csv(OUT_CSV_RAW)
            except Exception as e:
                print(f"Error reading CSV: {str(e)}")
                return None
    else:
        print(f"/timecards returned {r.status_code}: {r.text}")
        return None


def try_reports_export(access_token, start_date, end_date):
    """
    Export the 'Timecard Report' from Procore Reports API as CSV,
    normalize headers/encoding, drop Cost Code Long Name (col G), keep A..K,
    then save to OUT_CSV and return the cleaned DataFrame.
    """
    headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
    list_url = "https://api.procore.com/rest/v1.0/reports"
    r = requests.get(list_url, headers=headers, params={"company_id": COMPANY_ID})
    if r.status_code != 200:
        print(f"/reports list returned {r.status_code}: {r.text}")
        return None

    reports = r.json()
    candidates = [rep for rep in reports if "timecard" in (rep.get("name","") or "").lower()]
    if not candidates:
        print("No 'Timecard Report' found.")
        return None

    report_id = candidates[0].get("id") or candidates[0].get("report_id")
    export_url = f"https://api.procore.com/rest/v1.0/reports/{report_id}/export"
    params = {
        "company_id": COMPANY_ID,
        "start_date": start_date,
        "end_date": end_date,
        "format": "csv"
    }

    r2 = requests.get(export_url, headers={"Authorization": f"Bearer {access_token}"}, params=params, timeout=60)
    if r2.status_code != 200:
        print(f"Report export failed {r2.status_code}: {r2.text}")
        return None

    content = r2.content

    # 1) Try pandas with utf-8-sig
    df = None
    try:
        df = pd.read_csv(io.BytesIO(content), encoding="utf-8-sig")
    except Exception:
        # 2) Fallback: detect delimiter and try again with text decode
        text = None
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                text = content.decode(enc)
                break
            except Exception:
                continue
        if text is None:
            print("Unable to decode CSV from report export.")
            return None
        sample = "\n".join(text.splitlines()[:10])
        try:
            delim = csv.Sniffer().sniff(sample, delimiters=[",",";","\t","|"]).delimiter
        except Exception:
            delim = ","
        try:
            df = pd.read_csv(io.StringIO(text), sep=delim, engine="python")
        except Exception as e:
            print("Failed to parse exported CSV:", e)
            return None

    # Normalize column names: strip whitespace
    df.columns = [str(c).strip() for c in df.columns]

    # Drop columns matching Cost Code Long Name (case-insensitive)
    cols_to_drop = [c for c in df.columns if "cost" in c.lower() and "long" in c.lower() and "name" in c.lower()]
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop, errors='ignore')
    else:
        # defensive: if 7th column looks like cost code long name, drop it
        if df.shape[1] >= 7:
            c7 = df.columns[6]
            if "cost" in str(c7).lower() and "long" in str(c7).lower():
                df = df.drop(columns=[c7], errors='ignore')

    # Truncate to A..K (first 11 columns) if more exist
    if df.shape[1] > 11:
        df = df.iloc[:, :11]

    # Save final CSV with BOM so Excel opens cleanly
    try:
        df.to_csv(str(OUT_CSV), index=False, encoding="utf-8-sig")
    except Exception as e:
        print("Failed writing cleaned CSV:", e)
        return None

    print(f"Saved cleaned report to {OUT_CSV}")
    return df


# ----------------------------------------------------

# --------------- Clean and paste -------------------
def clean_timecard_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    cols_to_drop = [c for c in df.columns if c.lower().strip() in (
        "cost code long name", "cost_code_long_name", "cost_code_longname")]
    for c in cols_to_drop:
        df = df.drop(columns=[c], errors='ignore')
    if df.shape[1] > 11:
        df = df.iloc[:, :11]
    employee_cols = [c for c in df.columns if c.lower().strip() in ("employee name","employee_name","employee")]
    if not employee_cols:
        fallback = [c for c in df.columns if 'employee' in c.lower()]
        if fallback:
            employee_cols = [fallback[0]]
    if employee_cols:
        df = df.rename(columns={employee_cols[0]: "Employee Name"})
    else:
        df["Employee Name"] = "Unknown"
    return df

def paste_into_workbook(df: pd.DataFrame, workbook_path: str):
    if df is None or df.empty:
        print("No data to paste.")
        return
    if not os.path.exists(workbook_path):
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    wb = load_workbook(workbook_path)
    grouped = df.groupby("Employee Name")
    for employee, group in grouped:
        sheet_name = sanitize_sheet_name(employee)
        if sheet_name not in wb.sheetnames:
            print(f"Creating sheet: {sheet_name}")
            wb.create_sheet(sheet_name)
        ws = wb[sheet_name]
        if ws.max_row >= 2:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.value = None
        rows = group.values.tolist()
        start_row = 2
        for r_idx, row_vals in enumerate(rows, start=start_row):
            for c_idx, value in enumerate(row_vals, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(workbook_path)
    print(f"Pasted timecards into workbook: {workbook_path}")
# ----------------------------------------------------

# -------------------- Main -------------------------
def main():
    print("1) Obtain access token (will refresh or run manual auth if needed)...")
    access_token = get_access_token()

    start_date = input("Enter pay period START date (YYYY-MM-DD, Monday): ").strip()
    end_date   = input("Enter pay period END date (YYYY-MM-DD, Sunday): ").strip()
    try:
        datetime.strptime(start_date, "%Y-%m-%d")
        datetime.strptime(end_date, "%Y-%m-%d")
    except Exception:
        print("Invalid date format. Use YYYY-MM-DD.")
        return

    print("2) Trying company-level /timecards endpoint...")
    df = try_timecards_endpoint(access_token, start_date, end_date)

    if df is None or df.empty:
        print("3) Falling back to /timecard_entries endpoint...")
        df = try_timecard_entries_endpoint(access_token, start_date, end_date)

    if df is None or df.empty:
        print("4) Falling back to Reports API export (searching for a 'Timecard' report)...")
        df = try_reports_export(access_token, start_date, end_date)

    if df is None or df.empty:
        print("No timecard data returned from API. If your account requires the UI 'Timecard Report' export,")
        print("please export the CSV manually and put it at:", OUT_CSV_RAW)
        return


    excel_path = input("Enter path to Excel workbook to paste into (e.g. timecards.xlsx): ").strip()
    if not excel_path:
        print("No workbook provided; exiting.")
        return

    try:
        paste_into_workbook(df_clean, excel_path)
    except Exception as e:
        print("Failed to paste into workbook:", e)
        print("You can manually open:", OUT_CSV_CLEAN)
        return

    print("Done.")



import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from datetime import datetime
import threading

# import your existing functions (from your code above)
# e.g. get_access_token, try_timecards_endpoint, try_timecard_entries_endpoint, try_reports_export,
# clean_timecard_df, paste_into_workbook, etc.

# ------------- GUI Helpers ----------------
def run_timecard_process(start_date, end_date, excel_path, log_box):
    try:
        log_box.insert(tk.END, "1) Getting access token...\n")
        access_token = get_access_token()

        # validate dates
        try:
            datetime.strptime(start_date, "%Y-%m-%d")
            datetime.strptime(end_date, "%Y-%m-%d")
        except Exception:
            messagebox.showerror("Date Error", "Dates must be in YYYY-MM-DD format")
            return

        log_box.insert(tk.END, "2) Trying /timecards endpoint...\n")
        df = try_timecards_endpoint(access_token, start_date, end_date)

        if df is None or df.empty:
            log_box.insert(tk.END, "Falling back to /timecard_entries...\n")
            df = try_timecard_entries_endpoint(access_token, start_date, end_date)

        if df is None or df.empty:
            log_box.insert(tk.END, "Falling back to Reports API export...\n")
            df = try_reports_export(access_token, start_date, end_date)

        if df is None or df.empty:
            messagebox.showwarning("No Data", "No timecard data returned from API.\n"
                                  "Please export CSV manually.")
            return

        df.to_csv(OUT_CSV, index=False)
        log_box.insert(tk.END, f"Saved cleaned report to {OUT_CSV}\n")


        if not excel_path:
            messagebox.showwarning("No Workbook", "No Excel workbook selected")
            return

        paste_into_workbook(df_clean, excel_path)
        log_box.insert(tk.END, f"Pasted data into {excel_path}\n")
        messagebox.showinfo("Success", "Timecards successfully processed!")

    except Exception as e:
        log_path = log_error_to_file(e)
        messagebox.showerror("Error", f"Process failed:\n{e}\n\nError saved to:\n{log_path}")


def start_process(start_entry, end_entry, excel_var, log_box):
    start_date = start_entry.get().strip()
    end_date = end_entry.get().strip()
    excel_path = excel_var.get().strip()

    # run in thread so UI doesn’t freeze
    threading.Thread(
        target=run_timecard_process,
        args=(start_date, end_date, excel_path, log_box),
        daemon=True
    ).start()


def select_excel(excel_var):
    path = filedialog.askopenfilename(
        title="Select Excel or CSV File", 
        filetypes=[("All Files", "*.*"),  # Allows all file types
                  ("CSV files", "*.csv"),  # Allow CSV files
                  ("Excel files", "*.xlsx *.xls *.xlsm *.xltx")]  # Allow all Excel formats
    )
    if path:
        excel_var.set(path)

# ------------- Main UI --------------------
def launch_ui():
    root = tk.Tk()
    root.title("Procore Timecard Downloader")

    tk.Label(root, text="Start Date (YYYY-MM-DD):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    start_entry = tk.Entry(root, width=15)
    start_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(root, text="End Date (YYYY-MM-DD):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    end_entry = tk.Entry(root, width=15)
    end_entry.grid(row=1, column=1, padx=5, pady=5)

    excel_var = tk.StringVar()
    tk.Label(root, text="Excel Workbook:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    tk.Entry(root, textvariable=excel_var, width=40).grid(row=2, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse...", command=lambda: select_excel(excel_var)).grid(row=2, column=2, padx=5, pady=5)

    log_box = scrolledtext.ScrolledText(root, width=60, height=15, state="normal")
    log_box.grid(row=3, column=0, columnspan=3, padx=5, pady=10)

    run_btn = tk.Button(
        root,
        text="Run",
        command=lambda: start_process(start_entry, end_entry, excel_var, log_box)
    )
    run_btn.grid(row=4, column=1, pady=10)

    root.mainloop()

# Replace your main() with:
if __name__ == "__main__":
    launch_ui()
