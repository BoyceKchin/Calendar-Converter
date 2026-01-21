let pyodide;
let pyodideReady = initPyodide();

const convertBtn = document.getElementById("convertBtn");

pyodideReady.then(() => {
    convertBtn.disabled = false;
    convertBtn.innerText = "Convert";
});


async function initPyodide() {
    pyodide = await loadPyodide();

    await pyodide.loadPackage("pandas");
    await pyodide.loadPackage("micropip");

    await pyodide.runPythonAsync(`
import micropip

# Pin versions compatible with Python 3.11 (Pyodide)
await micropip.install([
    "ics==0.6.2",
    "tatsu==5.8.3",
    "pytz",
    "python-dateutil",
])
    `);
}



async function runPython() {
    await pyodideReady;
    const fileInput = document.getElementById("csvInput");
    const status = document.getElementById("status");
    const downloadLink = document.getElementById("downloadLink");

    if (!fileInput.files.length) {
        alert("Please select a CSV file!");
        return;
    }

    const file = fileInput.files[0];
    const text = await file.text();
    pyodide.FS.writeFile(file.name, text);

    status.innerText = "Processing...";

    const pyCode = `
import pandas as pd
import re
from ics import Calendar, Event
import pytz
from datetime import datetime
from dateutil import parser


LOCAL_TZ = pytz.timezone("America/New_York")

input_file = "${file.name}"
output_file = input_file.replace(".csv", ".ics")

df = pd.read_csv(input_file, skiprows=3)
df.columns = df.columns.str.strip()

cols_to_drop = ["B","D","E","F","G","K","M","N","O","P","Q","R"]
drop_cols = [df.columns[ord(c)-ord("A")] for c in cols_to_drop if ord(c)-ord("A")<len(df.columns)]
df = df.drop(columns=drop_cols, errors="ignore")

if len(df.columns) >=6:
    df = df.rename(columns={df.columns[5]: "Description"})

if {"Work Activity", "Work Location"}.issubset(df.columns):
    df["Work Activity"] = (df["Work Activity"].fillna("").astype(str).str.strip() + " " + df["Work Location"].fillna("").astype(str).str.strip()).str.strip()

if {"Work Location", "Meeting Location"}.issubset(df.columns):
    df["Location"] = (df["Work Location"].fillna("").astype(str).str.strip() + " " + df["Meeting Location"].fillna("").astype(str).str.strip()).str.strip()
    df = df.drop(columns=["Work Location","Meeting Location"])

if "Date" in df.columns:
    df["Date"] = df["Date"].ffill()

if "Time" in df.columns:
    df = df[df["Time"].notna() & (df["Time"].astype(str).str.strip()!="")]

    def fix_time_format(t):
        t = str(t).strip()
        t = re.sub(r"(\\d)(am|pm)", r"\\1 \\2", t, flags=re.IGNORECASE)
        t = re.sub(r"[–—]", "-", t)
        t = re.sub(r"\\s+", " ", t)
        
        # Match "10:30 - 5 PM" or "10 - 5 PM"
        match = re.match(r"^(\\d{1,2}(?::\\d{2})?)\\s*-\\s*(\\d{1,2}(?::\\d{2})?)\\s*(AM|PM)$", t, flags=re.IGNORECASE)
        if match:
            start, end, meridian = match.groups()
            if not re.search(r"(AM|PM)", start, flags=re.IGNORECASE):
                start = start + f" {meridian.upper()}"
            end = end + f" {meridian.upper()}"
            return f"{start} - {end}"
        
        # Match "10:30 AM - 5:00 PM" format
        match = re.match(r"^(\\d{1,2}(?::\\d{2})?\\s*(?:AM|PM))\\s*-\\s*(\\d{1,2}(?::\\d{2})?\\s*(?:AM|PM))$", t, flags=re.IGNORECASE)
        if match:
            start, end = match.groups()
            return f"{start.upper()} - {end.upper()}"
        
        return t

    df["Time"] = df["Time"].apply(fix_time_format)
    df["Start Date"] = df["Date"].astype(str).str.strip()
    df["End Date"] = df["Start Date"]
    df["Start Time"] = df["Time"].str.extract(r"^(\\d{1,2}(?::\\d{2})?\\s*(?:AM|PM))", expand=False).fillna("").str.strip()
    df["End Time"] = df["Time"].str.extract(r"-\\s*(\\d{1,2}(?::\\d{2})?\\s*(?:AM|PM))$", expand=False).fillna("").str.strip()
    df = df.drop(columns=["Date","Time"], errors="ignore")

cal = Calendar()
for _, row in df.iterrows():
    event = Event()
    event.name = str(row.get("Work Activity",""))
    event.description = str(row.get("Description",""))
    event.location = str(row.get("Location",""))
    
    start_str = f"{row['Start Date']} {row['Start Time']}"
    end_str = f"{row['End Date']} {row['End Time']}"
    
    try:
        start_dt = parser.parse(start_str)
        end_dt = parser.parse(end_str)
        
        if end_dt <= start_dt:
            end_dt += pd.Timedelta(days=1)
        
        if start_dt.tzinfo is None:
            start_dt = LOCAL_TZ.localize(start_dt)
        if end_dt.tzinfo is None:
            end_dt = LOCAL_TZ.localize(end_dt)
        
        event.begin = start_dt
        event.end = end_dt
    except:
        continue
    
    cal.events.add(event)

with open(output_file, "w", encoding="utf-8") as f:
    f.writelines(cal)

output_file
`;

    try {
        const outputFile = await pyodide.runPythonAsync(pyCode);
        const icsData = pyodide.FS.readFile(outputFile, { encoding: "utf8" });

        const blob = new Blob([icsData], { type: "text/calendar" });
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = outputFile;
        downloadLink.style.display = "inline";
        downloadLink.innerText = "Download ICS File";

        status.innerText = "Conversion complete!";
    } catch (err) {
        status.innerText = "Error: " + err;
    }
}




