let filteredData = [];
let outputWorkbook = null;

document.getElementById('compareButton').addEventListener('click', compareFiles);
document.getElementById('downloadButton').addEventListener('click', downloadFile);

// Normalize values → lowercase, no spaces
function normalize(value) {
  return value.toString().toLowerCase().replace(/\s+/g, '');
}

function updateProgress(percent) {
  const bar = document.getElementById("progress-bar");
  const container = document.getElementById("progress-container");
  container.style.display = "block";
  bar.style.width = percent + "%";
}

function hideProgress() {
  setTimeout(() => {
    updateProgress(0);
    document.getElementById("progress-container").style.display = "none";
  }, 800);
}

async function compareFiles() {
  updateProgress(10);
  document.getElementById("download-controls").style.display = "none";
  document.getElementById("status-message").textContent = "";
  document.getElementById("status-message").className = "";

  const file1 = document.getElementById('file1').files[0];
  const file2 = document.getElementById('file2').files[0];

  if (!file1 || !file2) {
    alert('Please upload both files.');
    updateProgress(0);
    return;
  }

  const [data1, data2] = await Promise.all([readFile(file1), readFile(file2)]);
  updateProgress(30);

  const wb1 = XLSX.read(data1, { type: 'binary' });
  const wb2 = XLSX.read(data2, { type: 'binary' });

  const sheet1 = wb1.Sheets[wb1.SheetNames[0]];
  const sheet2 = wb2.Sheets[wb2.SheetNames[0]];

  const json1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
  const json2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });

  const header1 = json1[0];
  const header2 = json2[0];

  const handleColIndex = header1.indexOf("Handle");
  const bodyFabricColIndex = header2.indexOf("Body/Fabric");

  if (handleColIndex === -1) {
    alert('Could not find "Handle" column in Existing Inventory.');
    updateProgress(0);
    return;
  }

  if (bodyFabricColIndex === -1) {
    alert('Could not find "Body/Fabric" column in New Inventory.');
    updateProgress(0);
    return;
  }

  // Build style number set from New Inventory
  const styleSet = new Set();
  for (let i = 1; i < json2.length; i++) {
    const val = json2[i][bodyFabricColIndex];
    if (val !== undefined && val !== null) {
      styleSet.add(normalize(val));
    }
  }

  updateProgress(50);

  // Filter matching rows
  filteredData = [header1];
  let matchCount = 0;

  for (let i = 1; i < json1.length; i++) {
    const row = json1[i];
    const cell = row[handleColIndex];
    if (cell && styleSet.has(normalize(cell))) {
      filteredData.push(row);
      matchCount++;
    }
  }

  const message = document.getElementById('status-message');
  message.className = '';

  if (matchCount === 0) {
    message.textContent = "❌ No matches found. No file will be downloaded.";
    message.classList.add("error");
    updateProgress(100);
    hideProgress();
    return;
  } else {
    message.textContent = `✔ ${matchCount} style # match${matchCount > 1 ? "es" : ""} found.`;
    message.classList.add("success");
    document.getElementById("download-controls").style.display = "block";
  }

  // Prepare the workbook for later download
  const newSheet = XLSX.utils.aoa_to_sheet(filteredData);
  outputWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outputWorkbook, newSheet, "Matches Only");

  updateProgress(90);
  setTimeout(() => {
    updateProgress(100);
    hideProgress();
  }, 500);
}

function downloadFile() {
  if (!filteredData.length || !outputWorkbook) {
    alert("No data to download. Please run a comparison first.");
    return;
  }

  const format = document.getElementById("outputFormat").value;
  let blob, filename;

  if (format === 'xlsx') {
    const wbout = XLSX.write(outputWorkbook, { bookType: 'xlsx', type: 'binary' });
    blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    filename = "matching_styles.xlsx";
  } else if (format === 'csv') {
    const csv = XLSX.utils.sheet_to_csv(outputWorkbook.Sheets["Matches Only"]);
    blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    filename = "matching_styles.csv";
  } else {
    const txt = XLSX.utils.sheet_to_txt(outputWorkbook.Sheets["Matches Only"]);
    blob = new Blob([txt], { type: "text/plain;charset=utf-8" });
    filename = "matching_styles.txt";
  }

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.target.result);
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}
