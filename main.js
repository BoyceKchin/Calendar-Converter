let pyodide;

async function initPyodide() {
  pyodide = await loadPyodide();
  // Load your script.py into the Pyodide environment
  await pyodide.runPythonAsync(await (await fetch("format_calendar2.py")).text());
}
initPyodide();

async function runPython() {
  let input = document.getElementById("userInput").value;
  let result = await pyodide.runPythonAsync(`process(${input})`);
  document.getElementById("output").innerText = result;
}
