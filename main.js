<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>CSV → ICS Converter</title>
</head>
<body>

<h1>CSV → ICS Converter</h1>

<input type="file" id="csvInput" accept=".csv">
<button id="convertBtn">Convert</button>

<p id="status"></p>
<a id="downloadLink" style="display:none">Download ICS</a>

<script src="https://cdn.jsdelivr.net/pyodide/v0.23.4/full/pyodide.js"></script>
<script>
let pyodide;

async function initPyodide() {
    pyodide = await loadPyodide();
    await pyodide.loadPackage("pandas");
    await pyodide.loadPackage("micropip");

    await pyodide.runPythonAsync(`
import micropip
await micropip.install(["ics","pytz"])
`);
}
initPyodide();

document.getElementById("convertBtn").addEventListener("click", runPython);

async function runPython() {
    const fileInput = document.getElementById("csvInput");
    const status = document.getElementById("status");
    const downloadLink = document.getElementById("downloadLink");

    if (!fileInput.files.length) {
        alert("Please select a CSV file!");
        return;
    }

    const file = fileInput.files[0];
    const text = await file.text();

    // Write the file to Pyodide FS
    pyodide.FS.writeFile(file.name, text);

    status.innerText = "Processing...";

    const pyCode = `
import pandas as pd
import re
from ics import Calendar, Event
import pytz
from datetime import datetime

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
        t = re.sub(r"[–—]", "-", t)
        t = re.sub(r"\s+", " ", t)
        t = re.sub(r"(\d)(am|pm)", r"\\1 \\2", t, flags=re.IGNORECASE)
        t = re.sub(r"(\d)(AM|PM)", r"\\1 \\2", t, flags=re.IGNORECASE)
        t = re.sub(r"\s*-\s*", " - ", t)
        pattern = r"(?i)^([0-9]{1,2}(?::[0-9]{2})?\\s*(?:AM|PM))\\s*-\\s*([0-9]{1,2}(?::[0-9]{2})?\\s*(?:AM|PM)?)$"
        m = re.match(pattern, t)
        if m:
            start, end = m.groups()
            if not re.search(r"(AM|PM)$", end, re.IGNORECASE):
                meridian = re.search(r"(AM|PM)$", start, re.IGNORECASE)
                if meridian:
                    end += " " + meridian.group(1)
            return f"{start.upper()} - {end.upper()}"
        return t

    df["Time"] = df["Time"].apply(fix_time_format)
    df["Start Date"] = df["Date"].astype(str).str.strip()
    df["End Date"] = df["Start Date"]
    df["Start Time"] = df["Time"].str.extract(r"^([0-9]{1,2}(?::[0-9]{2})?\\s*(?:AM|PM))", expand=False).fillna("").str.strip()
    df["End Time"] = df["Time"].str.extract(r"-\\s*([0-9]{1,2}(?::[0-9]{2})?\\s*(?:AM|PM))$", expand=False).fillna("").str.strip()
    df = df.drop(columns=["Date","Time"], errors="ignore")

cal = Calendar()

def parse_time(date_str, time_str):
    for fmt in ["%m/%d/%Y %I %p","%m/%d/%Y %I:%M %p"]:
        try:
            return datetime.strptime(f"{date_str} {time_str}", fmt)
        except:
            continue
    raise ValueError(f"Unrecognized time format: {time_str}")

for _, row in df.iterrows():
    event = Event()
    event.name = str(row.get("Work Activity",""))
    event.description = str(row.get("Description",""))
    event.location = str(row.get("Location",""))
    try:
        start_dt = parse_time(row['Start Date'], row['Start Time'])
        end_dt = parse_time(row['End Date'], row['End Time'])
        if end_dt <= start_dt:
            end_dt = end_dt + pd.Timedelta(days=1)
        start_dt = LOCAL_TZ.localize(start_dt.replace(tzinfo=None))
        end_dt = LOCAL_TZ.localize(end_dt.replace(tzinfo=None))
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
</script>

</body>
</html>
