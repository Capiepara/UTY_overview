let workbook;
let rawData = [];
let filteredData = [];
let headers = [];

let charts = {};

const excelFile = document.getElementById("excelFile");
const sheetSelect = document.getElementById("sheetSelect");
const generateBtn = document.getElementById("generateBtn");
const fileName = document.getElementById("fileName");
const dashboard = document.getElementById("dashboard");

const seasonFilter = document.getElementById("seasonFilter");
const stageFilter = document.getElementById("stageFilter");
const factoryFilter = document.getElementById("factoryFilter");
const materialFilter = document.getElementById("materialFilter");
const resetBtn = document.getElementById("resetBtn");

excelFile.addEventListener("change", handleFile);
generateBtn.addEventListener("click", generateDashboard);
resetBtn.addEventListener("click", resetFilters);

[seasonFilter, stageFilter, factoryFilter, materialFilter].forEach(filter => {
  filter.addEventListener("change", applyFilters);
});

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  fileName.textContent = file.name;

  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    workbook = XLSX.read(data, { type: "array" });

    sheetSelect.innerHTML = "";

    workbook.SheetNames.forEach(sheet => {
      const option = document.createElement("option");
      option.value = sheet;
      option.textContent = sheet;
      sheetSelect.appendChild(option);
    });

    const defaultSheet = workbook.SheetNames.find(s =>
      s.toLowerCase().includes("development")
    );

    if (defaultSheet) sheetSelect.value = defaultSheet;

    generateBtn.disabled = false;
  };

  reader.readAsArrayBuffer(file);
}

function generateDashboard() {
  const sheetName = sheetSelect.value;
  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  const headerRowIndex = findHeaderRow(rows);
  headers = rows[headerRowIndex].map(h => String(h).trim());

  rawData = rows.slice(headerRowIndex + 1)
    .filter(row => row.some(cell => cell !== ""))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] ?? "";
      });
      return obj;
    });

  buildFilters();
  applyFilters();

  dashboard.classList.remove("hidden");
}

function findHeaderRow(rows) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].map(x => String(x).trim().toLowerCase());
    if (
      row.includes("season") &&
      row.includes("dev. code") &&
      row.includes("dev. stage")
    ) {
      return i;
    }
  }

  return 0;
}

function getCol(row, names) {
  for (const name of names) {
    const key = Object.keys(row).find(k =>
      k.trim().toLowerCase() === name.trim().toLowerCase()
    );
    if (key) return row[key];
  }
  return "";
}

function buildFilters() {
  fillSelect(seasonFilter, uniqueValues(rawData, ["Season"]));
  fillSelect(stageFilter, uniqueValues(rawData, ["DEV. Stage"]));
  fillSelect(factoryFilter, uniqueValues(rawData, ["Factory"]));
  fillSelect(materialFilter, uniqueValues(rawData, ["Material Indicator"]));
}

function fillSelect(select, values) {
  select.innerHTML = "";

  values.forEach(value => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    option.selected = true;
    select.appendChild(option);
  });
}

function uniqueValues(data, colNames) {
  return [...new Set(
    data.map(row => String(getCol(row, colNames)).trim())
      .filter(v => v !== "")
  )].sort();
}

function selectedValues(select) {
  return Array.from(select.selectedOptions).map(o => o.value);
}

function applyFilters() {
  const seasons = selectedValues(seasonFilter);
  const stages = selectedValues(stageFilter);
  const factories = selectedValues(factoryFilter);
  const materials = selectedValues(materialFilter);

  filteredData = rawData.filter(row => {
    const season = String(getCol(row, ["Season"])).trim();
    const stage = String(getCol(row, ["DEV. Stage"])).trim();
    const factory = String(getCol(row, ["Factory"])).trim();
    const material = String(getCol(row, ["Material Indicator"])).trim();

    return (
      seasons.includes(season) &&
      stages.includes(stage) &&
      factories.includes(factory) &&
      materials.includes(material)
    );
  });

  updateDashboard();
}

function resetFilters() {
  [seasonFilter, stageFilter, factoryFilter, materialFilter].forEach(select => {
    Array.from(select.options).forEach(option => option.selected = true);
  });

  applyFilters();
}

function distinctDev(data) {
  return new Set(
    data.map(row => String(getCol(row, ["DEV. Code"])).trim())
      .filter(v => v !== "")
  ).size;
}

function countDistinctBy(data, colNames) {
  const map = {};

  data.forEach(row => {
    const key = String(getCol(row, colNames)).trim() || "Blank";
    const dev = String(getCol(row, ["DEV. Code"])).trim();

    if (!dev) return;

    if (!map[key]) map[key] = new Set();
    map[key].add(dev);
  });

  return Object.fromEntries(
    Object.entries(map).map(([k, v]) => [k, v.size])
  );
}

function countStage(stageName, status = null) {
  const data = filteredData.filter(row => {
    const stage = String(getCol(row, ["DEV. Stage"])).trim().toUpperCase();
    const activeDrop = String(getCol(row, ["Active/Drop"])).trim().toUpperCase();

    if (stage !== stageName.toUpperCase()) return false;

    if (status) return activeDrop.includes(status.toUpperCase());

    return true;
  });

  return distinctDev(data);
}

function calcWaterfall() {
  const lr2Total = countStage("LR2");
  const lr2Dropped = countStage("LR2", "Dropped");

  const flcTotal = Math.max(lr2Total - lr2Dropped, 0);
  const flcDropped = countStage("FLC", "Dropped");

  const smsTotal = Math.max(flcTotal - flcDropped, 0);
  const smsDropped = countStage("SMS", "Dropped");

  const cfmTotal = Math.max(smsTotal - smsDropped, 0);

  return {
    labels: [
      "LR2 Total",
      "LR2 Dropped",
      "FLC Total",
      "FLC Dropped",
      "SMS Total",
      "SMS Dropped",
      "CFM Total"
    ],
    values: [
      lr2Total,
      -lr2Dropped,
      flcTotal,
      -flcDropped,
      smsTotal,
      -smsDropped,
      cfmTotal
    ],
    cfmTotal
  };
}

function calcOntimeByStage() {
  const stages = ["LR2", "FLC", "SMS", "CFM"];
  const result = {};

  stages.forEach(stageName => {
    const rows = filteredData.filter(row =>
      String(getCol(row, ["DEV. Stage"])).trim().toUpperCase() === stageName
    );

    const total = distinctDev(rows);

    const ontime = rows.filter(row => {
      const v = getCol(row, ["Shipment ontime", "Shipment Ontime"]);
      return Number(v) === 1 || String(v).trim().toLowerCase() === "yes";
    });

    const ontimeCount = distinctDev(ontime);
    result[stageName] = total === 0 ? 0 : Math.round((ontimeCount / total) * 100);
  });

  return result;
}

function updateDashboard() {
  const totalDev = distinctDev(filteredData);
  const dropped = distinctDev(
    filteredData.filter(row =>
      String(getCol(row, ["Active/Drop"])).trim().toUpperCase().includes("DROPPED")
    )
  );

  const ontime = filteredData.filter(row => {
    const v = getCol(row, ["Shipment ontime", "Shipment Ontime"]);
    return Number(v) === 1 || String(v).trim().toLowerCase() === "yes";
  });

  const ontimePercent = totalDev === 0 ? 0 : Math.round((distinctDev(ontime) / totalDev) * 100);

  const waterfall = calcWaterfall();

  document.getElementById("kpiDev").textContent = totalDev;
  document.getElementById("kpiOntime").textContent = ontimePercent + "%";
  document.getElementById("kpiDropped").textContent = dropped;
  document.getElementById("kpiCfm").textContent = waterfall.cfmTotal;

  drawWaterfall(waterfall);
  drawOntime(calcOntimeByStage());

  drawBar("seasonChart", "DEV Code by Season", countDistinctBy(filteredData, ["Season"]));
  drawBar("factoryChart", "DEV Code by Factory", countDistinctBy(filteredData, ["Factory"]));
  drawBar("stageChart", "DEV Code by Stage", countDistinctBy(filteredData, ["DEV. Stage"]));
  drawBar("materialChart", "DEV Code by Material", countDistinctBy(filteredData, ["Material Indicator"]));

  renderTable();
}

function destroyChart(id) {
  if (charts[id]) {
    charts[id].destroy();
    charts[id] = null;
  }
}

function drawBar(canvasId, title, dataObj) {
  destroyChart(canvasId);

  const labels = Object.keys(dataObj);
  const values = Object.values(dataObj);

  charts[canvasId] = new Chart(document.getElementById(canvasId), {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: title,
        data: values,
        backgroundColor: "#2563eb",
        borderRadius: 8
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { precision: 0 }
        },
        x: {
          ticks: {
            autoSkip: false,
            maxRotation: 45,
            minRotation: 0
          }
        }
      }
    }
  });
}

function drawOntime(dataObj) {
  destroyChart("ontimeChart");

  charts.ontimeChart = new Chart(document.getElementById("ontimeChart"), {
    type: "bar",
    data: {
      labels: Object.keys(dataObj),
      datasets: [{
        label: "Ontime %",
        data: Object.values(dataObj),
        backgroundColor: "#14b8a6",
        borderRadius: 8
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      scales: {
        y: {
          beginAtZero: true,
          max: 100,
          ticks: {
            callback: value => value + "%"
          }
        }
      }
    }
  });
}

function drawWaterfall(wf) {
  destroyChart("waterfallChart");

  let running = 0;

  const floatingData = wf.values.map(value => {
    if (value >= 0) {
      const start = 0;
      const end = value;
      running = value;
      return [start, end];
    } else {
      const start = running;
      const end = running + value;
      running = end;
      return [end, start];
    }
  });

  charts.waterfallChart = new Chart(document.getElementById("waterfallChart"), {
    type: "bar",
    data: {
      labels: wf.labels,
      datasets: [{
        label: "DEV Code",
        data: floatingData,
        backgroundColor: wf.values.map(v => v < 0 ? "#fb923c" : "#2563eb"),
        borderRadius: 8
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: context => {
              const value = wf.values[context.dataIndex];
              return value;
            }
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { precision: 0 }
        }
      }
    }
  });
}

function renderTable() {
  const table = document.getElementById("detailTable");

  const showHeaders = [
    "Season",
    "Active/Drop",
    "DEV. Stage",
    "DEV. Code",
    "Factory",
    "Model",
    "Gender / Model",
    "Colorway",
    "Material Indicator",
    "HQ XFD",
    "Actual XFD",
    "Shipment ontime",
    "Buy Ready"
  ];

  const availableHeaders = showHeaders.filter(h =>
    headers.some(x => x.trim().toLowerCase() === h.trim().toLowerCase())
  );

  let html = "<thead><tr>";
  availableHeaders.forEach(h => html += `<th>${h}</th>`);
  html += "</tr></thead><tbody>";

  filteredData.forEach(row => {
    html += "<tr>";
    availableHeaders.forEach(h => {
      html += `<td>${getCol(row, [h])}</td>`;
    });
    html += "</tr>";
  });

  html += "</tbody>";
  table.innerHTML = html;
}
