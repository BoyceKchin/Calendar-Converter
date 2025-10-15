let pyodide;

async function initPyodideEnv() {
    // Initialize Pyodide and required packages
    pyodide = await loadPyodide();
    await pyodide.loadPackage("pandas");
    await pyodide.loadPackage("micropip");

    // Install PyPI packages
    await pyodide.runPythonAsync(`
        import micropip
        await micropip.install(["ics", "pytz", "openpyxl"])
    `);
    console.log("✅ Pyodide ready with pandas, ics, pytz, openpyxl");
}

async function runPython(file) {
    if (!file) throw new Error("No Excel file provided");

    // Read the Excel file as binary
    const arrayBuffer = await file.arrayBuffer();
    const uint8Array = new Uint8Array(arrayBuffer);

    // Write it into Pyodide’s virtual filesystem
    pyodide.FS.writeFile(file.name, uint8Array);

    // Main Python script (Excel → ICS)
    const pyCode = `
import pandas as pd
import re
from ics import Calendar, Event
import pytz
from datetime import datetime

LOCAL_TZ = pytz.timezone("America/New_York")

input_file = "${file.name}"
output_file = input_file.replace(".xlsx", ".ics")

# Read Excel file (skip first 3 rows)
df = pd.read_excel(input_file, skiprows=3, engine="openpyxl")
df.columns = df.columns.str.strip()

# Drop unwanted columns
cols_to_drop = ["B","D","E","F","G","K","M","N","O","P","Q","R"]
drop_cols = [df.columns[ord(c)-ord("A")] for c in cols_to_drop if ord(c)-ord("A")<len(df.columns)]
df = df.drop(columns=drop_cols, errors="ignore")

# Rename column L → Description if present
if len(df.columns) >= 6:
    df = df.rename(columns={df.columns[5]: "Description"})

# Combine Work Activity and Work Location
if {"Work Activity", "Work Location"}.issubset(df.columns):
    df["Work Activity"] = (
        df["Work Activity"].fillna("").astype(str).str.strip() + " " +
        df["Work Location"].fillna("").astype(str).str.strip()
    ).str.strip()

# Combine Meeting/Work Locations
if {"Work Location", "Meeting Location"}.issubset(df.columns):
    df["Location"] = (
        df["Work Location"].fillna("").astype(str).str.strip() + " " +
        df["Meeting Location"].fillna("").astype(str).str.strip()
    ).str.strip()
    df = df.drop(columns=["Work Location","Meeting Location"], errors="ignore")

# Fill missing dates
if "Date" in df.columns:
    df["Date"] = df["Date"].ffill()

# Fix and extract time ranges
if "Time" in df.columns:
    df = df[df["Time"].notna() & (df["Time"].astype(str).str.strip()!="")]

    def fix_time_format(t):
        t=str(t).strip()
        t = re.sub(r"(\\d)(am|pm)", r"\\1 \\2", t, flags=re.IGNORECASE)
        t = re.sub(r"[–—]","-",t)
        t = re.sub(r"\\s+"," ",t)
        match = re.match(r"^(\\d+)\\s*-\\s*(\\d+)\\s*(AM|PM)$", t, flags=re.IGNORECASE)
        if match:
            start,end,meridian = match.groups()
            return f"{start} {meridian.upper()} - {end} {meridian.upper()}"
        match = re.match(r"^(\\d+\\s*(?:AM|PM))\\s*-\\s*(\\d+\\s*(?:AM|PM))$", t, flags=re.IGNORECASE)
        if match:
            start,end = match.groups()
            return f"{start.upper()} - {end.upper()}"
        return t

    df["Time"] = df["Time"].apply(fix_time_format)
    df["Start Date"] = df["Date"].astype(str).str.strip()
    df["End Date"] = df["Start Date"]
    df["Start Time"] = df["Time"].str.extract(r"^(\\d+\\s*(?:AM|PM))", expand=False).fillna("").str.strip()
    df["End Time"] = df["Time"].str.extract(r"-\\s*(\\d+\\s*(?:AM|PM))$", expand=False).fillna("").str.strip()
    df = df.drop(columns=["Date","Time"], errors="ignore")

# Build calendar
cal = Calendar()
for _, row in df.iterrows():
    event = Event()
    event.name = str(row.get("Work Activity",""))
    event.description = str(row.get("Description",""))
    event.location = str(row.get("Location",""))
    start_str = f"{row['Start Date']} {row['Start Time']}"
    end_str = f"{row['End Date']} {row['End Time']}"
    try:
        start_dt = datetime.strptime(start_str,"%m/%d/%Y %I %p")
        end_dt = datetime.strptime(end_str,"%m/%d/%Y %I %p")
        if end_dt <= start_dt:
            end_dt = end_dt + pd.Timedelta(days=1)
        if start_dt.tzinfo is None:
            start_dt = start_dt.replace(tzinfo=LOCAL_TZ)
        if end_dt.tzinfo is None:
            end_dt = end_dt.replace(tzinfo=LOCAL_TZ)
        event.begin = start_dt
        event.end = end_dt
    except Exception:
        continue
    cal.events.add(event)

with open(output_file, "w", encoding="utf-8") as f:
    f.writelines(cal)

output_file
`;
