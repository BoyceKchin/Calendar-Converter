import pandas as pd
import re
from ics import Calendar, Event
from datetime import datetime
import os
from tkinter import Tk, filedialog, messagebox
import traceback
from zoneinfo import ZoneInfo

# Set local timezone (developer message: user timezone = America/New_York)
LOCAL_TZ = ZoneInfo("America/New_York")


def main():
    # -----------------------------
    # 1️⃣ Open file dialog to select CSV
    # -----------------------------
    root = Tk()
    root.withdraw()  # Hide the main tkinter window

    input_file = filedialog.askopenfilename(
        title="Select CSV File",
        filetypes=[("CSV Files", "*.csv")]
    )

    if not input_file:
        messagebox.showwarning("No File Selected", "No CSV file was chosen. Exiting program.")
        return

    # Output ICS file (same name as input CSV)
    output_file = os.path.splitext(input_file)[0] + ".ics"



    # -----------------------------
    # 2️⃣ Read and clean CSV
    # -----------------------------
    df = pd.read_csv(input_file, skiprows=3)

    # Strip column names
    df.columns = df.columns.str.strip()

    # Drop unwanted columns by letter
    cols_to_drop = ["B", "D", "E", "F", "G", "K", "M", "N", "O", "P", "Q", "R"]
    drop_cols = [df.columns[ord(c) - ord("A")] for c in cols_to_drop if ord(c) - ord("A") < len(df.columns)]
    df = df.drop(columns=drop_cols, errors="ignore")

    # Rename the 6th column (F) to Description
    if len(df.columns) >= 6:
        df = df.rename(columns={df.columns[5]: "Description"})

    # Merge Work Location into Work Activity
    if {"Work Activity", "Work Location"}.issubset(df.columns):
        df["Work Activity"] = (
            df["Work Activity"].fillna("").astype(str).str.strip()
            + " "
            + df["Work Location"].fillna("").astype(str).str.strip()
        ).str.strip()

    # Merge Work Location + Meeting Location into Location
    if {"Work Location", "Meeting Location"}.issubset(df.columns):
        df["Location"] = (
            df["Work Location"].fillna("").astype(str).str.strip()
            + " "
            + df["Meeting Location"].fillna("").astype(str).str.strip()
        ).str.strip()
        df = df.drop(columns=["Work Location", "Meeting Location"])

    # Forward fill Date column
    if "Date" in df.columns:
        df["Date"] = df["Date"].ffill()

    # Fix Time column
    if "Time" in df.columns:
        # Drop rows with empty Time
        df = df[df["Time"].notna() & (df["Time"].astype(str).str.strip() != "")]

        def fix_time_format(t):
            t = str(t).strip()

            # Normalize spacing: insert a space before AM/PM if missing
            t = re.sub(r"(\d)(am|pm)", r"\1 \2", t, flags=re.IGNORECASE)

            # Normalize dashes (hyphen, en-dash, em-dash → "-")
            t = re.sub(r"[–—]", "-", t)

            # Remove extra spaces
            t = re.sub(r"\s+", " ", t)

            # Case 1: "9-11 AM" → "9 AM - 11 AM"
            match = re.match(r"^(\d+)\s*-\s*(\d+)\s*(AM|PM)$", t, re.IGNORECASE)
            if match:
                start, end, meridian = match.groups()
                return f"{start} {meridian.upper()} - {end} {meridian.upper()}"

            # Case 2: "10 PM - 5 AM" or "9 AM - 5 PM"
            match = re.match(r"^(\d+\s*(?:AM|PM))\s*-\s*(\d+\s*(?:AM|PM))$", t, re.IGNORECASE)
            if match:
                start, end = match.groups()
                return f"{start.upper()} - {end.upper()}"

            return t

        df["Time"] = df["Time"].apply(fix_time_format)

        # Split into Start/End Dates and Times
        df["Start Date"] = df["Date"].astype(str).str.strip()
        df["End Date"] = df["Start Date"]

       # Extract Start and End Times from cleaned string
        df["Start Time"] = df["Time"].str.extract(
            r"^(\d+\s*(?:AM|PM))", expand=False
        ).fillna("").str.strip()

        df["End Time"] = df["Time"].str.extract(
            r"-\s*(\d+\s*(?:AM|PM))$", expand=False
        ).fillna("").str.strip()

        # Drop original Date and Time
        df = df.drop(columns=["Date", "Time"], errors="ignore")

    # -----------------------------
    # 3️⃣ Convert to ICS
    # -----------------------------
    cal = Calendar()

    for _, row in df.iterrows():
        event = Event()
        event.name = str(row.get("Work Activity", ""))
        event.description = str(row.get("Description", ""))
        event.location = str(row.get("Location", ""))

        start_str = f"{row['Start Date']} {row['Start Time']}"
        end_str = f"{row['End Date']} {row['End Time']}"

        try:
            start_dt = datetime.strptime(start_str, "%m/%d/%Y %I %p")
            end_dt = datetime.strptime(end_str, "%m/%d/%Y %I %p")

            # Handle overnight shifts
            if end_dt <= start_dt:
                end_dt = end_dt + pd.Timedelta(days=1)

            # Ensure they are Python datetimes (not pandas) and tz-aware
            # If they are pandas.Timestamp, convert:
            if hasattr(start_dt, "to_pydatetime"):
                start_dt = start_dt.to_pydatetime()
            if hasattr(end_dt, "to_pydatetime"):
                end_dt = end_dt.to_pydatetime()

            # If still naive, attach local tz
            if start_dt.tzinfo is None:
                start_dt = start_dt.replace(tzinfo=LOCAL_TZ)
            if end_dt.tzinfo is None:
                end_dt = end_dt.replace(tzinfo=LOCAL_TZ)


            event.begin = start_dt
            event.end = end_dt
        except Exception:
            continue  # Skip rows that can't be parsed

        cal.events.add(event)

    # -----------------------------
    # 4️⃣ Save ICS file
    # -----------------------------
    with open(output_file, "w", encoding="utf-8") as f:
        f.writelines(cal)

    # Success popup
    messagebox.showinfo("Conversion Successful", f"ICS calendar saved as:\n{output_file}")




if __name__ == "__main__":
    try:
        main()
    except Exception:
        error_msg = traceback.format_exc()
        messagebox.showerror("Error", f"An error occurred:\n{error_msg}")
