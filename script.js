let workbook = null;
let rawData = [];
let filteredData = [];
let headers = [];
let choiceControls = {};
let searchTerm = "";

const COLORS = {
  orange: "#F6A623",
  coral: "#EF5B5B",
  cyan: "#35A9CF",
  teal: "#59C3A5",
  navy: "#405B73",
  cream: "#FFF8F3",
  gray: "#C7D0D6",
  dark: "#20303C"
};

const MILESTONES = [
  { name: "CAD received", plan: ["TP/CAD Handover (Plan)"], actual: ["TP/ CAD received (Actual)"] },
  { name: "Material arrived", plan: ["Material  Arrive (Plan)"], actual: ["Latest  Material Arrive (Actual)"] },
  { name: "Pattern lock", plan: ["Pattern lock (Plan)"], actual: ["Pattern locked (Actual)"] },
  { name: "Tooling lock", plan: ["Tooling lock (Plan)"], actual: ["Tooling locked (Actual)"] },
  { name: "Mold finished", plan: ["Mold ETC (Plan)"], actual: ["Mold finished (Actual)"] }
];

const excelFile = document.getElementById("excelFile");
const sheetSelect = document.getElementById("sheetSelect");
const generateBtn = document.getElementById("generateBtn");
const dashboard = document.getElementById("dashboard");
const loadBadge = document.getElementById("loadBadge");

excelFile.addEventListener("change", loadWorkbook);
generateBtn.addEventListener("click", generateDashboard);
document.getElementById("resetBtn").addEventListener("click", resetFilters);
document.getElementById("healthStage").addEventListener("change", updateDashboard);
document.getElementById("tableSearch").addEventListener("input", e => {
  searchTerm = e.target.value.trim().toLowerCase();
  renderTable();
});

function loadWorkbook(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  document.getElementById("fileName").textContent = file.name;
  document.getElementById("fileMeta").textContent = `${(file.size / 1024 / 1024).toFixed(2)} MB`;

  const reader = new FileReader();
  reader.onload = e => {
    workbook = XLSX.read(new Uint8Array(e.target.result), {
      type: "array",
      cellDates: true
    });

    sheetSelect.innerHTML = "";
    workbook.SheetNames.forEach(name => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      sheetSelect.appendChild(option);
    });

    const defaultSheet = workbook.SheetNames.find(name =>
      normalize(name).includes("development tracking")
    );
    if (defaultSheet) sheetSelect.value = defaultSheet;

    sheetSelect.disabled = false;
    generateBtn.disabled = false;
    loadBadge.textContent = `${workbook.SheetNames.length} sheets loaded`;
    loadBadge.classList.add("loaded");
  };
  reader.readAsArrayBuffer(file);
}

function generateDashboard() {
  const sheet = workbook.Sheets[sheetSelect.value];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: true
  });

  const headerRowIndex = findHeaderRow(rows);
  if (headerRowIndex < 0) {
    alert("Could not find the header row containing Season, DEV. Stage and DEV. Code.");
    return;
  }

  headers = rows[headerRowIndex].map(value => String(value ?? "").trim());
  rawData = rows.slice(headerRowIndex + 1)
    .filter(row => row.some(value => value !== "" && value !== null))
    .map(row => {
      const record = {};
      headers.forEach((header, index) => {
        if (header) record[header] = row[index] ?? "";
      });
      return record;
    })
    .filter(row => text(getCol(row, ["DEV. Code"])) !== "");

  buildFilters();
  dashboard.classList.remove("hidden");
  applyFilters();
  dashboard.scrollIntoView({ behavior: "smooth", block: "start" });
}

function findHeaderRow(rows) {
  return rows.findIndex(row => {
    const normalized = row.map(normalize);
    return normalized.includes("season")
      && normalized.includes("dev. stage")
      && normalized.includes("dev. code");
  });
}

function normalize(value) {
  return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

function text(value) {
  return String(value ?? "").trim();
}

function getCol(row, aliases) {
  const keys = Object.keys(row);
  for (const alias of aliases) {
    const target = normalize(alias);
    const exact = keys.find(key => normalize(key) === target);
    if (exact) return row[exact];
  }
  return "";
}

function distinctCount(data, columnAliases) {
  return new Set(
    data.map(row => text(getCol(row, columnAliases))).filter(Boolean)
  ).size;
}

function uniqueValues(data, aliases) {
  return [...new Set(
    data.map(row => text(getCol(row, aliases))).filter(Boolean)
  )].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function buildFilters() {
  Object.values(choiceControls).forEach(control => control.destroy());
  choiceControls = {};

  const configs = [
    ["seasonFilter", ["Season"]],
    ["stageFilter", ["DEV. Stage"]],
    ["factoryFilter", ["Factory"]],
    ["materialFilter", ["Material Indicator"]],
    ["modelFilter", ["Model"]]
  ];

  configs.forEach(([id, aliases]) => {
    const element = document.getElementById(id);
    element.innerHTML = "";
    uniqueValues(rawData, aliases).forEach(value => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      option.selected = true;
      element.appendChild(option);
    });

    const control = new Choices(element, {
      removeItemButton: true,
      shouldSort: false,
      searchEnabled: true,
      placeholder: true,
      placeholderValue: "Select values",
      itemSelectText: ""
    });
    element.addEventListener("change", applyFilters);
    choiceControls[id] = control;
  });
}

function selectedValues(id) {
  return choiceControls[id]?.getValue(true) ?? [];
}

function applyFilters() {
  const selections = {
    season: selectedValues("seasonFilter"),
    stage: selectedValues("stageFilter"),
    factory: selectedValues("factoryFilter"),
    material: selectedValues("materialFilter"),
    model: selectedValues("modelFilter")
  };

  filteredData = rawData.filter(row => {
    const values = {
      season: text(getCol(row, ["Season"])),
      stage: text(getCol(row, ["DEV. Stage"])),
      factory: text(getCol(row, ["Factory"])),
      material: text(getCol(row, ["Material Indicator"])),
      model: text(getCol(row, ["Model"]))
    };
    return Object.keys(selections).every(key =>
      selections[key].length === 0 || selections[key].includes(values[key])
    );
  });

  updateDashboard();
}

function resetFilters() {
  Object.entries(choiceControls).forEach(([id, control]) => {
    const aliases = {
      seasonFilter: ["Season"],
      stageFilter: ["DEV. Stage"],
      factoryFilter: ["Factory"],
      materialFilter: ["Material Indicator"],
      modelFilter: ["Model"]
    }[id];
    control.removeActiveItems();
    control.setChoiceByValue(uniqueValues(rawData, aliases));
  });
  searchTerm = "";
  document.getElementById("tableSearch").value = "";
  applyFilters();
}

function isDropped(row) {
  return normalize(getCol(row, ["Active/Drop"])).includes("drop");
}

function isActive(row) {
  const status = normalize(getCol(row, ["Active/Drop"]));
  return status.includes("active") || (!status.includes("drop") && status !== "");
}

function stageRows(stage, data = filteredData) {
  const target = normalize(stage);
  if (target === "cfm") {
    return data.filter(row =>
      normalize(getCol(row, ["DEV. Stage"])) === "sms" && isActive(row)
    );
  }
  return data.filter(row => normalize(getCol(row, ["DEV. Stage"])) === target);
}

function stageCount(stage, status = "all", data = filteredData) {
  let rows = stageRows(stage, data);
  if (status === "drop") rows = rows.filter(isDropped);
  if (status === "active") rows = rows.filter(isActive);
  return distinctCount(rows, ["DEV. Code"]);
}

function getWaterfallData() {
  const lr2 = stageCount("LR2");
  const lr2Drop = stageCount("LR2", "drop");
  const flc = stageCount("FLC");
  const flcDrop = stageCount("FLC", "drop");
  const sms = stageCount("SMS");
  const smsDrop = stageCount("SMS", "drop");
  const cfm = Math.max(sms - smsDrop, 0);

  return {
    labels: ["LR2 total", "LR2 dropped", "FLC total", "FLC dropped", "SMS total", "SMS dropped", "CFM"],
    values: [lr2, -lr2Drop, flc, -flcDrop, sms, -smsDrop, cfm],
    measures: ["absolute", "relative", "absolute", "relative", "absolute", "relative", "absolute"],
    cfm
  };
}

function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const s = text(value);
  if (!s) return null;
  const timestamp = Date.parse(s);
  return Number.isNaN(timestamp) ? null : new Date(timestamp);
}

function semanticStatus(value) {
  const s = normalize(value);
  if (!s) return "blank";
  if (["-", "n/a", "na", "no need", "none", "not required"].includes(s)) return "exempt";
  if (/(existing|locked|arrived|finished|completed|complete|approved|pass|sent)/.test(s)) return "complete";
  return "other";
}

function milestoneOutcome(row, milestone) {
  const planRaw = getCol(row, milestone.plan);
  const actualRaw = getCol(row, milestone.actual);
  const planState = semanticStatus(planRaw);
  const actualState = semanticStatus(actualRaw);

  if (planState === "exempt" || actualState === "exempt") return { score: 1, status: "exempt" };
  if (actualState === "complete") return { score: 1, status: "ontime" };

  const planDate = parseExcelDate(planRaw);
  const actualDate = parseExcelDate(actualRaw);

  if (planDate && actualDate) {
    const delayDays = Math.ceil((actualDate - planDate) / 86400000);
    return delayDays <= 0
      ? { score: 1, status: "ontime", delayDays }
      : { score: 0, status: "late", delayDays };
  }

  if (!planDate && actualDate) return { score: 1, status: "ontime" };

  if (planDate && !actualDate) {
    const overdue = Date.now() > planDate.getTime();
    return overdue
      ? { score: 0, status: "overdue" }
      : { score: null, status: "not-due" };
  }

  if (actualState === "other") return { score: 1, status: "ontime" };
  return { score: null, status: "unknown" };
}

function milestoneMetrics(stage) {
  const rows = stageRows(stage);
  const metrics = MILESTONES.map(milestone => {
    const devOutcomes = new Map();

    rows.forEach(row => {
      const dev = text(getCol(row, ["DEV. Code"]));
      if (!dev) return;
      const outcome = milestoneOutcome(row, milestone);
      const previous = devOutcomes.get(dev);

      // Conservative deduplication: a late/overdue record wins over an on-time duplicate.
      if (!previous || outcome.score === 0 || previous.score === null) {
        devOutcomes.set(dev, outcome);
      }
    });

    const eligible = [...devOutcomes.values()].filter(x => x.score !== null);
    const successes = eligible.filter(x => x.score === 1).length;
    return {
      name: milestone.name,
      rate: eligible.length ? Math.round(successes / eligible.length * 100) : 0,
      eligible: eligible.length
    };
  });

  const totalEligible = metrics.reduce((sum, item) => sum + item.eligible, 0);
  const weightedRate = totalEligible
    ? Math.round(metrics.reduce((sum, item) => sum + item.rate * item.eligible, 0) / totalEligible)
    : 0;

  return { metrics, weightedRate };
}

function overallMilestoneRate() {
  const stages = ["LR2", "FLC", "SMS"];
  let numerator = 0;
  let denominator = 0;
  stages.forEach(stage => {
    const { metrics } = milestoneMetrics(stage);
    metrics.forEach(item => {
      numerator += item.rate * item.eligible;
      denominator += item.eligible;
    });
  });
  return denominator ? Math.round(numerator / denominator) : 0;
}

function updateDashboard() {
  const waterfall = getWaterfallData();
  document.getElementById("kpiDev").textContent = distinctCount(filteredData, ["DEV. Code"]);
  document.getElementById("kpiHealth").textContent = `${overallMilestoneRate()}%`;
  document.getElementById("kpiCfm").textContent = waterfall.cfm;
  document.getElementById("kpiModels").textContent = distinctCount(filteredData, ["Model"]);

  drawWaterfall(waterfall);
  drawHealth();
  drawSankey();
  drawHeatmap();
  renderTable();
}

function baseLayout(extra = {}) {
  return {
    paper_bgcolor: "rgba(0,0,0,0)",
    plot_bgcolor: "rgba(0,0,0,0)",
    font: { family: "Inter, sans-serif", color: COLORS.dark, size: 12 },
    margin: { l: 48, r: 20, t: 20, b: 55 },
    ...extra
  };
}

const plotConfig = { responsive: true, displayModeBar: false };

function drawWaterfall(data) {
  Plotly.react("waterfallChart", [{
    type: "waterfall",
    orientation: "v",
    measure: data.measures,
    x: data.labels,
    y: data.values,
    text: data.values.map(v => `${v > 0 && !data.labels[data.values.indexOf(v)].includes("dropped") ? "" : ""}${v}`),
    textposition: "inside",
    textfont: { color: "#ffffff", size: 12 },
    connector: { line: { color: "#AAB4BB", width: 1 } },
    increasing: { marker: { color: COLORS.teal, line: { color: "#ffffff", width: 1 } } },
    decreasing: { marker: { color: COLORS.coral, line: { color: "#ffffff", width: 1 } } },
    totals: { marker: { color: COLORS.navy, line: { color: "#ffffff", width: 1 } } },
    hovertemplate: "%{x}<br><b>%{y}</b> DEV codes<extra></extra>"
  }], baseLayout({
    showlegend: false,
    xaxis: { fixedrange: true, tickangle: -20, showgrid: false, zeroline: false },
    yaxis: { fixedrange: true, rangemode: "tozero", gridcolor: "#EFE7E1", zerolinecolor: "#B9C1C6" }
  }), plotConfig);
}

function scoreColor(rate) {
  if (rate >= 90) return COLORS.teal;
  if (rate >= 75) return COLORS.cyan;
  if (rate >= 60) return COLORS.orange;
  return COLORS.coral;
}

function drawHealth() {
  const stage = document.getElementById("healthStage").value;
  const { metrics, weightedRate } = milestoneMetrics(stage);

  Plotly.react("healthChart", [{
    type: "pie",
    labels: metrics.map(x => `${x.name}<br>${x.rate}%`),
    values: metrics.map(() => 1),
    hole: .58,
    sort: false,
    direction: "clockwise",
    marker: {
      colors: metrics.map(x => scoreColor(x.rate)),
      line: { color: "#ffffff", width: 4 }
    },
    textinfo: "label",
    textposition: "inside",
    textfont: { color: "#ffffff", size: 11 },
    hovertemplate: "%{label}<extra></extra>"
  }], baseLayout({
    margin: { l: 10, r: 10, t: 10, b: 10 },
    showlegend: false,
    annotations: [{
      text: `<b>${stage}</b><br><span style="font-size:28px">${weightedRate}%</span><br><span style="font-size:11px;color:#71808D">on time</span>`,
      x: .5, y: .5, showarrow: false, align: "center"
    }]
  }), plotConfig);
}

function sankeyRows() {
  const stages = ["LR2", "FLC", "SMS", "CFM"];
  const materials = uniqueValues(filteredData, ["Material Indicator"])
    .filter(value => ["TN", "NU", "CONC"].includes(value.toUpperCase()));

  const links = [];
  stages.forEach(stage => {
    const rows = stageRows(stage);
    materials.forEach(material => {
      const count = distinctCount(
        rows.filter(row => text(getCol(row, ["Material Indicator"])) === material),
        ["DEV. Code"]
      );
      if (count > 0) links.push({ stage, material, count });
    });
  });
  return { stages, materials, links };
}

function drawSankey() {
  const { stages, materials, links } = sankeyRows();
  const labels = [...stages, ...materials];
  const stageColors = [COLORS.orange, COLORS.coral, COLORS.cyan, COLORS.navy];
  const materialColors = materials.map((_, i) => [COLORS.teal, "#91C9C0", "#E8C988"][i % 3]);

  Plotly.react("sankeyChart", [{
    type: "sankey",
    arrangement: "snap",
    node: {
      label: labels,
      color: [...stageColors, ...materialColors],
      pad: 24,
      thickness: 20,
      line: { color: "#ffffff", width: 1 }
    },
    link: {
      source: links.map(x => labels.indexOf(x.stage)),
      target: links.map(x => labels.indexOf(x.material)),
      value: links.map(x => x.count),
      color: links.map(x => {
        const base = stageColors[stages.indexOf(x.stage)];
        return hexToRgba(base, .45);
      }),
      customdata: links.map(x => x.count),
      hovertemplate: "%{source.label} → %{target.label}<br><b>%{value}</b> DEV codes<extra></extra>"
    }
  }], baseLayout({
    margin: { l: 10, r: 10, t: 18, b: 18 }
  }), plotConfig);
}

function hexToRgba(hex, alpha) {
  const value = hex.replace("#", "");
  const bigint = parseInt(value, 16);
  return `rgba(${(bigint >> 16) & 255},${(bigint >> 8) & 255},${bigint & 255},${alpha})`;
}

function drawHeatmap() {
  const models = uniqueValues(filteredData, ["Model"])
    .map(model => ({
      model,
      count: distinctCount(filteredData.filter(row => text(getCol(row, ["Model"])) === model), ["DEV. Code"])
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 20)
    .map(x => x.model);

  const stages = ["LR2", "FLC", "SMS", "CFM"];
  const z = models.map(model => {
    const modelRows = filteredData.filter(row => text(getCol(row, ["Model"])) === model);
    const counts = stages.map(stage => distinctCount(
      stageRows(stage, modelRows), ["DEV. Code"]
    ));
    const denominator = counts.reduce((sum, value, index) =>
      sum + (index === 3 ? 0 : value), 0
    ) || 1;

    // CFM is derived from SMS active, so percentage uses unique base stages as denominator.
    return counts.map(value => Math.round(value / denominator * 100));
  });

  Plotly.react("heatmapChart", [{
    type: "heatmap",
    x: stages,
    y: models,
    z,
    zmin: 0,
    zmax: 100,
    colorscale: [
      [0, "#FFF4E7"],
      [.35, "#F6C27A"],
      [.65, "#59C3A5"],
      [1, "#405B73"]
    ],
    text: z.map(row => row.map(value => `${value}%`)),
    texttemplate: "%{text}",
    textfont: { size: 11 },
    hovertemplate: "%{y}<br>%{x}: <b>%{z}%</b><extra></extra>",
    colorbar: { title: "Share", ticksuffix: "%", thickness: 12 }
  }], baseLayout({
    margin: { l: 120, r: 55, t: 20, b: 45 },
    xaxis: { side: "top", fixedrange: true, showgrid: false },
    yaxis: { fixedrange: true, autorange: "reversed", tickfont: { size: 10 } }
  }), plotConfig);
}

function renderTable() {
  const table = document.getElementById("detailTable");
  const displayHeaders = [
    "Season", "Active/Drop", "DEV. Stage", "DEV. Code", "Factory",
    "Model", "Gender / Model", "Material Indicator",
    "TP/CAD Handover (Plan)", "TP/ CAD received (Actual)",
    "Material  Arrive (Plan)", "Latest  Material Arrive (Actual)",
    "Pattern lock (Plan)", "Pattern locked (Actual)",
    "Tooling lock (Plan)", "Tooling locked (Actual)",
    "Mold ETC (Plan)", "Mold finished (Actual)", "Shipment ontime"
  ];

  const rows = filteredData.filter(row => {
    if (!searchTerm) return true;
    return [
      getCol(row, ["DEV. Code"]),
      getCol(row, ["Model"]),
      getCol(row, ["Factory"]),
      getCol(row, ["Season"])
    ].some(value => normalize(value).includes(searchTerm));
  }).slice(0, 500);

  table.innerHTML = `
    <thead><tr>${displayHeaders.map(header => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
    <tbody>
      ${rows.map(row => `<tr>${
        displayHeaders.map(header => `<td>${formatCell(getCol(row, [header]))}</td>`).join("")
      }</tr>`).join("")}
    </tbody>
  `;
}

function formatCell(value) {
  const date = parseExcelDate(value);
  if (date && (value instanceof Date || typeof value === "number")) {
    return date.toLocaleDateString("en-GB");
  }
  return escapeHtml(text(value));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
