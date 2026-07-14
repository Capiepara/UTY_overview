const APP_VERSION = "v5.0.0";
document.getElementById("appVersion").textContent = APP_VERSION;

let workbook = null;
let rawData = [];
let filteredData = [];
let headers = [];
let choiceControls = {};
let searchTerm = "";
let riskRecords = [];

const COLORS = {
  orange: "#F6A623",
  coral: "#EF5B5B",
  cyan: "#35A9CF",
  teal: "#59C3A5",
  navy: "#405B73",
  cream: "#FFF8F3",
  gray: "#C7D0D6",
  dark: "#263746"
};

const MILESTONES = [
  { name: "CAD received", shortName: "CAD", weight: 10,
    plan: ["TP/CAD Handover (Plan)"], actual: ["TP/ CAD received (Actual)"] },
  { name: "Material arrived", shortName: "Material", weight: 20,
    plan: ["Material  Arrive (Plan)", "Material Arrive (Plan)"],
    actual: ["Latest  Material Arrive (Actual)", "Latest Material Arrive (Actual)"] },
  { name: "Pattern locked", shortName: "Pattern", weight: 20,
    plan: ["Pattern lock (Plan)"], actual: ["Pattern locked (Actual)"] },
  { name: "Tooling locked", shortName: "Tooling", weight: 25,
    plan: ["Tooling lock (Plan)"], actual: ["Tooling locked (Actual)"] },
  { name: "Mold finished", shortName: "Mold", weight: 25,
    plan: ["Mold ETC (Plan)"], actual: ["Mold finished (Actual)"] }
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
document.getElementById("detailStage").addEventListener("change", updateMaterialDetails);
document.getElementById("detailMaterial").addEventListener("change", updateMaterialDetails);
document.getElementById("timelineDev").addEventListener("change", drawTimeline);
document.getElementById("tableSearch").addEventListener("input", event => {
  searchTerm = event.target.value.trim().toLowerCase();
  renderTable();
});

function normalize(value) {
  return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}
function text(value) { return String(value ?? "").trim(); }
function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;").replaceAll("'", "&#039;");
}

function getCol(row, aliases) {
  const keys = Object.keys(row);
  for (const alias of aliases) {
    const wanted = normalize(alias);
    const exact = keys.find(key => normalize(key) === wanted);
    if (exact) return row[exact];
  }
  return "";
}

function distinctCount(data, aliases) {
  return new Set(data.map(row => text(getCol(row, aliases))).filter(Boolean)).size;
}

function uniqueValues(data, aliases) {
  return [...new Set(data.map(row => text(getCol(row, aliases))).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function loadWorkbook(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  document.getElementById("fileName").textContent = file.name;
  document.getElementById("fileMeta").textContent =
    `${(file.size / 1024 / 1024).toFixed(2)} MB`;

  const reader = new FileReader();
  reader.onload = readEvent => {
    workbook = XLSX.read(new Uint8Array(readEvent.target.result), {
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
    loadBadge.textContent = `${workbook.SheetNames.length} sheets loaded • ${APP_VERSION}`;
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

  const headerRowIndex = rows.findIndex(row => {
    const normalized = row.map(normalize);
    return normalized.includes("season")
      && normalized.includes("dev. stage")
      && normalized.includes("dev. code");
  });

  if (headerRowIndex < 0) {
    alert("The header row could not be found. Please choose the Development Tracking sheet.");
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
  const aliasMap = {
    seasonFilter: ["Season"],
    stageFilter: ["DEV. Stage"],
    factoryFilter: ["Factory"],
    materialFilter: ["Material Indicator"],
    modelFilter: ["Model"]
  };

  Object.entries(choiceControls).forEach(([id, control]) => {
    control.removeActiveItems();
    control.setChoiceByValue(uniqueValues(rawData, aliasMap[id]));
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
    labels: ["LR2 total", "LR2 dropped", "FLC total", "FLC dropped",
      "SMS total", "SMS dropped", "CFM"],
    values: [lr2, -lr2Drop, flc, -flcDrop, sms, -smsDrop, cfm],
    measures: ["absolute", "relative", "absolute", "relative",
      "absolute", "relative", "absolute"],
    cfm
  };
}

function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const valueText = text(value);
  if (!valueText) return null;
  const timestamp = Date.parse(valueText);
  return Number.isNaN(timestamp) ? null : new Date(timestamp);
}

function semanticStatus(value) {
  const status = normalize(value);
  if (!status) return "blank";
  if (["-", "n/a", "na", "no need", "none", "not required"].includes(status)) return "exempt";
  if (/(existing|locked|arrived|finished|completed|complete|approved|pass|sent)/.test(status)) {
    return "complete";
  }
  return "other";
}

function milestoneOutcome(row, milestone) {
  const planRaw = getCol(row, milestone.plan);
  const actualRaw = getCol(row, milestone.actual);
  const planState = semanticStatus(planRaw);
  const actualState = semanticStatus(actualRaw);

  if (planState === "exempt" || actualState === "exempt") {
    return { score: 1, status: "not-needed", delayDays: 0 };
  }
  if (actualState === "complete") {
    return { score: 1, status: "on-time", delayDays: 0 };
  }

  const planDate = parseExcelDate(planRaw);
  const actualDate = parseExcelDate(actualRaw);

  if (planDate && actualDate) {
    const delayDays = Math.ceil((actualDate - planDate) / 86400000);
    return delayDays <= 0
      ? { score: 1, status: "on-time", delayDays: 0 }
      : { score: 0, status: "late", delayDays };
  }

  if (!planDate && actualDate) return { score: 1, status: "on-time", delayDays: 0 };

  if (planDate && !actualDate) {
    const overdueDays = Math.ceil((Date.now() - planDate.getTime()) / 86400000);
    return overdueDays > 0
      ? { score: 0, status: "overdue", delayDays: overdueDays }
      : { score: null, status: "not-due", delayDays: 0 };
  }

  if (actualState === "other") return { score: 1, status: "on-time", delayDays: 0 };
  return { score: null, status: "unknown", delayDays: null };
}

function dedupedDevRows(data = filteredData) {
  const map = new Map();
  data.forEach(row => {
    const dev = text(getCol(row, ["DEV. Code"]));
    if (!dev) return;
    if (!map.has(dev)) map.set(dev, []);
    map.get(dev).push(row);
  });
  return map;
}

function milestoneMetrics(stage) {
  const rows = stageRows(stage);
  const metrics = MILESTONES.map(milestone => {
    const outcomesByDev = new Map();

    rows.forEach(row => {
      const devCode = text(getCol(row, ["DEV. Code"]));
      if (!devCode) return;
      const outcome = milestoneOutcome(row, milestone);
      const previous = outcomesByDev.get(devCode);

      if (!previous || outcome.score === 0 || previous.score === null) {
        outcomesByDev.set(devCode, outcome);
      }
    });

    const eligible = [...outcomesByDev.values()].filter(item => item.score !== null);
    const onTime = eligible.filter(item => item.score === 1).length;

    return {
      name: milestone.name,
      shortName: milestone.shortName,
      rate: eligible.length ? Math.round(onTime / eligible.length * 100) : 0,
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
  let weightedPoints = 0;
  let totalEligible = 0;

  stages.forEach(stage => {
    milestoneMetrics(stage).metrics.forEach(item => {
      weightedPoints += item.rate * item.eligible;
      totalEligible += item.eligible;
    });
  });

  return totalEligible ? Math.round(weightedPoints / totalEligible) : 0;
}

function calculateRiskRecords() {
  const records = [];

  dedupedDevRows().forEach((rows, devCode) => {
    let risk = 0;
    const reasons = [];

    MILESTONES.forEach(milestone => {
      let worst = null;
      rows.forEach(row => {
        const outcome = milestoneOutcome(row, milestone);
        if (!worst || outcome.score === 0 || worst.score === null) worst = outcome;
      });

      if (worst?.score === 0) {
        const delayFactor = Math.min(Math.max(worst.delayDays || 1, 1) / 14, 1);
        const points = Math.round(milestone.weight * (0.55 + 0.45 * delayFactor));
        risk += points;
        reasons.push(`${milestone.shortName}: ${worst.delayDays || 0} days late`);
      }
    });

    const shipmentLate = rows.some(row => {
      const value = getCol(row, ["Shipment ontime", "Shipment Ontime"]);
      return Number(value) === 0 || normalize(value) === "no";
    });
    if (shipmentLate) {
      risk += 15;
      reasons.push("Shipment is late");
    }

    if (rows.some(isDropped)) {
      risk = Math.max(risk, 70);
      reasons.push("Dropped");
    }

    risk = Math.min(Math.round(risk), 100);
    const first = rows[0];

    records.push({
      devCode,
      risk,
      reasons,
      season: text(getCol(first, ["Season"])),
      stage: text(getCol(first, ["DEV. Stage"])),
      factory: text(getCol(first, ["Factory"])),
      model: text(getCol(first, ["Model"])),
      genderModel: text(getCol(first, ["Gender / Model", "Gender_Model"])),
      material: text(getCol(first, ["Material Indicator"]))
    });
  });

  return records.sort((a, b) => b.risk - a.risk);
}

function updateDashboard() {
  const waterfall = getWaterfallData();
  riskRecords = calculateRiskRecords();

  document.getElementById("kpiDev").textContent = distinctCount(filteredData, ["DEV. Code"]);
  document.getElementById("kpiHealth").textContent = `${overallMilestoneRate()}%`;
  document.getElementById("kpiRisk").textContent = riskRecords.filter(item => item.risk >= 60).length;
  document.getElementById("kpiCfm").textContent = waterfall.cfm;
  document.getElementById("kpiModels").textContent = distinctCount(filteredData, ["Model"]);

  renderInsights();
  drawWaterfall(waterfall);
  drawHealth();
  drawFactoryRanking();
  drawDelayHeatmap();
  drawSankey();
  buildMaterialDetailSelectors();
  updateMaterialDetails();
  renderCriticalTable();
  buildTimelineSelector();
  drawTimeline();
  renderTable();
}

function baseLayout(extra = {}) {
  return {
    paper_bgcolor: "rgba(0,0,0,0)",
    plot_bgcolor: "rgba(0,0,0,0)",
    font: { family: "Nunito, Tahoma, sans-serif", color: COLORS.dark, size: 12 },
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
    connector: { line: { color: "#A7AEB2", width: 1 } },
    increasing: { marker: { color: COLORS.teal, line: { color: "#fff", width: 1 } } },
    decreasing: { marker: { color: COLORS.coral, line: { color: "#fff", width: 1 } } },
    totals: { marker: { color: COLORS.navy, line: { color: "#fff", width: 1 } } },
    hovertemplate: "%{x}<br><b>%{y}</b> DEV codes<extra></extra>"
  }], baseLayout({
    margin: { l: 16, r: 12, t: 15, b: 58 },
    showlegend: false,
    xaxis: { fixedrange: true, tickangle: -18, showgrid: false, zeroline: false, showline: false },
    yaxis: { fixedrange: true, showgrid: false, zeroline: false, showline: false,
      showticklabels: false, ticks: "" }
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
  const values = metrics.map(() => 1);
  const labels = metrics.map(item => `${item.shortName}<br>${item.rate}%`);

  Plotly.react("healthChart", [{
    type: "pie",
    labels: [...labels, ""],
    values: [...values, values.reduce((sum, value) => sum + value, 0)],
    hole: .57,
    sort: false,
    direction: "clockwise",
    rotation: 270,
    marker: {
      colors: [...metrics.map(item => scoreColor(item.rate)), "rgba(0,0,0,0)"],
      line: { color: "#fff", width: 4 }
    },
    text: [...labels, ""],
    textinfo: "text",
    textposition: "outside",
    textfont: { color: COLORS.dark, size: 12, family: "Nunito, Tahoma, sans-serif" },
    pull: [.025, .025, .025, .025, .025, 0],
    hovertemplate: "<b>%{label}</b><extra></extra>",
    showlegend: false
  }], baseLayout({
    margin: { l: 45, r: 45, t: 10, b: 0 },
    showlegend: false,
    annotations: [{
      text: `<b>${stage}</b><br><span style="font-size:31px">${weightedRate}%</span><br>` +
        `<span style="font-size:11px;color:#70808E">on time</span>`,
      x: .5, y: .53, showarrow: false, align: "center"
    }]
  }), plotConfig);

  document.getElementById("healthSummary").innerHTML = metrics.map(item => `
    <div class="health-item" style="border-left-color:${scoreColor(item.rate)}">
      <strong>${item.rate}%</strong><span>${escapeHtml(item.name)}</span>
    </div>
  `).join("");
}

function factoryMetrics() {
  return uniqueValues(filteredData, ["Factory"]).map(factory => {
    const rows = filteredData.filter(row => text(getCol(row, ["Factory"])) === factory);
    let points = 0, eligible = 0;

    ["LR2", "FLC", "SMS"].forEach(stage => {
      const stageData = rows.filter(row => normalize(getCol(row, ["DEV. Stage"])) === normalize(stage));
      MILESTONES.forEach(milestone => {
        const seen = new Map();
        stageData.forEach(row => {
          const dev = text(getCol(row, ["DEV. Code"]));
          const outcome = milestoneOutcome(row, milestone);
          if (!dev) return;
          const previous = seen.get(dev);
          if (!previous || outcome.score === 0 || previous.score === null) seen.set(dev, outcome);
        });
        [...seen.values()].filter(item => item.score !== null).forEach(item => {
          eligible += 1;
          points += item.score;
        });
      });
    });

    return {
      factory,
      rate: eligible ? Math.round(points / eligible * 100) : 0,
      devCount: distinctCount(rows, ["DEV. Code"])
    };
  }).sort((a, b) => b.rate - a.rate);
}

function drawFactoryRanking() {
  const data = factoryMetrics().slice(0, 12).reverse();

  Plotly.react("factoryChart", [{
    type: "bar",
    orientation: "h",
    y: data.map(item => item.factory),
    x: data.map(item => item.rate),
    text: data.map(item => `${item.rate}%`),
    textposition: "outside",
    marker: { color: data.map(item => scoreColor(item.rate)) },
    customdata: data.map(item => item.devCount),
    hovertemplate: "%{y}<br>On time: <b>%{x}%</b><br>DEV codes: %{customdata}<extra></extra>"
  }], baseLayout({
    margin: { l: 70, r: 38, t: 15, b: 35 },
    xaxis: { range: [0, 105], ticksuffix: "%", showgrid: false, zeroline: false },
    yaxis: { showgrid: false }
  }), plotConfig);
}

function averageDelayForRows(rows, milestone) {
  const byDev = new Map();

  rows.forEach(row => {
    const dev = text(getCol(row, ["DEV. Code"]));
    if (!dev) return;
    const outcome = milestoneOutcome(row, milestone);
    if (outcome.delayDays === null || outcome.score === null) return;
    const previous = byDev.get(dev);
    if (previous === undefined || outcome.delayDays > previous) {
      byDev.set(dev, Math.max(outcome.delayDays, 0));
    }
  });

  const values = [...byDev.values()];
  return values.length ? Math.round(values.reduce((a, b) => a + b, 0) / values.length) : null;
}

function drawDelayHeatmap() {
  const models = uniqueValues(filteredData, ["Model"])
    .map(model => ({
      model,
      count: distinctCount(filteredData.filter(row => text(getCol(row, ["Model"])) === model), ["DEV. Code"])
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 16)
    .map(item => item.model);

  const z = models.map(model => {
    const rows = filteredData.filter(row => text(getCol(row, ["Model"])) === model);
    return MILESTONES.map(milestone => averageDelayForRows(rows, milestone));
  });

  Plotly.react("delayHeatmap", [{
    type: "heatmap",
    x: MILESTONES.map(item => item.shortName),
    y: models,
    z,
    zmin: 0,
    zmax: Math.max(14, ...z.flat().filter(value => value !== null)),
    colorscale: [
      [0, "#E3F4EF"],
      [.35, "#F6E7A8"],
      [.65, "#F6B35D"],
      [1, "#EF5B5B"]
    ],
    text: z.map(row => row.map(value => value === null ? "" : `${value}d`)),
    texttemplate: "%{text}",
    hovertemplate: "%{y}<br>%{x}: <b>%{z} days late</b><extra></extra>",
    colorbar: { title: "Days", thickness: 12 }
  }], baseLayout({
    margin: { l: 120, r: 55, t: 18, b: 45 },
    xaxis: { side: "top", showgrid: false },
    yaxis: { autorange: "reversed", tickfont: { size: 10 } }
  }), plotConfig);
}

function buildMaterialLinks() {
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

function hexToRgba(hex, alpha) {
  const value = hex.replace("#", "");
  const number = parseInt(value, 16);
  return `rgba(${(number >> 16) & 255},${(number >> 8) & 255},${number & 255},${alpha})`;
}

function drawSankey() {
  const { stages, materials, links } = buildMaterialLinks();
  const labels = [...stages, ...materials];
  const stageColors = [COLORS.orange, COLORS.coral, COLORS.cyan, COLORS.navy];
  const materialColors = materials.map((_, index) =>
    [COLORS.teal, "#88C8BF", "#D8B767"][index % 3]
  );

  Plotly.react("sankeyChart", [{
    type: "sankey",
    arrangement: "snap",
    node: {
      label: labels,
      color: [...stageColors, ...materialColors],
      pad: 24,
      thickness: 20,
      line: { color: "#fff", width: 1 }
    },
    link: {
      source: links.map(item => labels.indexOf(item.stage)),
      target: links.map(item => labels.indexOf(item.material)),
      value: links.map(item => item.count),
      color: links.map(item => hexToRgba(stageColors[stages.indexOf(item.stage)], .42)),
      customdata: links.map(item => [item.stage, item.material]),
      hovertemplate: "%{source.label} → %{target.label}<br><b>%{value}</b> DEV codes<extra></extra>"
    }
  }], baseLayout({ margin: { l: 10, r: 10, t: 18, b: 18 } }), plotConfig)
  .then(() => {
    const chart = document.getElementById("sankeyChart");
    if (typeof chart.removeAllListeners === "function") chart.removeAllListeners("plotly_click");

    chart.on("plotly_click", event => {
      const selected = event.points?.[0]?.customdata;
      if (!selected || selected.length < 2) return;
      document.getElementById("detailStage").value = selected[0];
      document.getElementById("detailMaterial").value = selected[1];
      updateMaterialDetails();
    });
  });
}

function buildMaterialDetailSelectors() {
  const stageSelect = document.getElementById("detailStage");
  const materialSelect = document.getElementById("detailMaterial");
  const oldStage = stageSelect.value;
  const oldMaterial = materialSelect.value;
  const stages = ["LR2", "FLC", "SMS", "CFM"];
  const materials = uniqueValues(filteredData, ["Material Indicator"])
    .filter(value => ["TN", "NU", "CONC"].includes(value.toUpperCase()));

  stageSelect.innerHTML = stages.map(stage =>
    `<option value="${escapeHtml(stage)}">${escapeHtml(stage)}</option>`).join("");
  materialSelect.innerHTML = materials.map(material =>
    `<option value="${escapeHtml(material)}">${escapeHtml(material)}</option>`).join("");

  if (stages.includes(oldStage)) stageSelect.value = oldStage;
  if (materials.includes(oldMaterial)) materialSelect.value = oldMaterial;
}

function updateMaterialDetails() {
  const stage = document.getElementById("detailStage").value;
  const material = document.getElementById("detailMaterial").value;
  const list = document.getElementById("materialDetailList");

  if (!stage || !material) {
    document.getElementById("materialDetailCount").textContent = "0";
    list.innerHTML = '<div class="material-detail-empty">No matching group is available.</div>';
    return;
  }

  const rows = stageRows(stage).filter(row =>
    text(getCol(row, ["Material Indicator"])) === material
  );
  const devMap = new Map();

  rows.forEach(row => {
    const devCode = text(getCol(row, ["DEV. Code"]));
    if (!devCode || devMap.has(devCode)) return;

    const genderModel = text(getCol(row, ["Gender / Model", "Gender_Model"]));
    const model = text(getCol(row, ["Model"]));

    devMap.set(devCode, {
      devCode,
      style: genderModel || model || "No style name",
      model
    });
  });

  const items = [...devMap.values()].sort((a, b) =>
    a.devCode.localeCompare(b.devCode, undefined, { numeric: true })
  );

  document.getElementById("materialDetailCount").textContent = items.length;
  list.innerHTML = items.length
    ? items.map(item => `
      <div class="material-detail-item">
        <div class="material-detail-code">${escapeHtml(item.devCode)}</div>
        <div class="material-detail-style">
          ${escapeHtml(item.style)}
          ${item.model && item.model !== item.style
            ? `<br><span style="color:#70808E">${escapeHtml(item.model)}</span>` : ""}
        </div>
      </div>`).join("")
    : '<div class="material-detail-empty">No DEV codes match this group.</div>';
}

function renderInsights() {
  const milestoneAll = MILESTONES.map(milestone => {
    let onTime = 0, eligible = 0;
    ["LR2", "FLC", "SMS"].forEach(stage => {
      const rows = stageRows(stage);
      const outcomes = new Map();

      rows.forEach(row => {
        const dev = text(getCol(row, ["DEV. Code"]));
        if (!dev) return;
        const outcome = milestoneOutcome(row, milestone);
        const previous = outcomes.get(dev);
        if (!previous || outcome.score === 0 || previous.score === null) outcomes.set(dev, outcome);
      });

      [...outcomes.values()].filter(item => item.score !== null).forEach(item => {
        eligible += 1;
        onTime += item.score;
      });
    });

    return {
      name: milestone.name,
      rate: eligible ? Math.round(onTime / eligible * 100) : 0
    };
  }).sort((a, b) => a.rate - b.rate);

  const factories = factoryMetrics().sort((a, b) => a.rate - b.rate);
  const materialCounts = uniqueValues(filteredData, ["Material Indicator"]).map(material => ({
    material,
    count: distinctCount(
      filteredData.filter(row => text(getCol(row, ["Material Indicator"])) === material),
      ["DEV. Code"]
    )
  })).sort((a, b) => b.count - a.count);

  const highRisk = riskRecords.filter(item => item.risk >= 60);
  const cards = [
    {
      color: COLORS.coral,
      title: "Weakest checkpoint",
      body: milestoneAll.length
        ? `${milestoneAll[0].name} has the lowest on-time rate at ${milestoneAll[0].rate}%.`
        : "No checkpoint data is available."
    },
    {
      color: COLORS.orange,
      title: "Factory to review",
      body: factories.length
        ? `${factories[0].factory} has the lowest on-time rate at ${factories[0].rate}%.`
        : "No factory data is available."
    },
    {
      color: COLORS.navy,
      title: "High-risk workload",
      body: `${highRisk.length} DEV codes have a risk score of 60 or above.`
    },
    {
      color: COLORS.teal,
      title: "Largest material group",
      body: materialCounts.length
        ? `${materialCounts[0].material} is the largest group with ${materialCounts[0].count} DEV codes.`
        : "No material data is available."
    }
  ];

  document.getElementById("insightGrid").innerHTML = cards.map(card => `
    <article class="insight-card" style="border-left-color:${card.color}">
      <strong>${escapeHtml(card.title)}</strong>
      <p>${escapeHtml(card.body)}</p>
    </article>
  `).join("");
}

function riskClass(score) {
  if (score >= 60) return "risk-high";
  if (score >= 30) return "risk-medium";
  return "risk-low";
}

function renderCriticalTable() {
  const rows = riskRecords.slice(0, 20);

  document.getElementById("criticalTable").innerHTML = `
    <thead><tr>
      <th>Risk</th><th>DEV Code</th><th>Season</th><th>Stage</th>
      <th>Factory</th><th>Model</th><th>Material</th><th>Main reasons</th>
    </tr></thead>
    <tbody>
      ${rows.map(item => `
        <tr>
          <td class="${riskClass(item.risk)}">${item.risk}</td>
          <td>${escapeHtml(item.devCode)}</td>
          <td>${escapeHtml(item.season)}</td>
          <td>${escapeHtml(item.stage)}</td>
          <td>${escapeHtml(item.factory)}</td>
          <td>${escapeHtml(item.model)}</td>
          <td>${escapeHtml(item.material)}</td>
          <td>${escapeHtml(item.reasons.slice(0, 3).join("; ") || "No current warning")}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function buildTimelineSelector() {
  const select = document.getElementById("timelineDev");
  const oldValue = select.value;
  const records = riskRecords.length ? riskRecords : calculateRiskRecords();

  select.innerHTML = records.map(item =>
    `<option value="${escapeHtml(item.devCode)}">${escapeHtml(item.devCode)} — ` +
    `${escapeHtml(item.model || item.genderModel || "")} — risk ${item.risk}</option>`
  ).join("");

  if (records.some(item => item.devCode === oldValue)) select.value = oldValue;
}

function drawTimeline() {
  const devCode = document.getElementById("timelineDev").value;
  const rows = filteredData.filter(row => text(getCol(row, ["DEV. Code"])) === devCode);

  if (!rows.length) {
    Plotly.purge("timelineChart");
    return;
  }

  const row = rows[0];
  const checkpointNames = MILESTONES.map(item => item.shortName);
  const planDates = MILESTONES.map(item => parseExcelDate(getCol(row, item.plan)));
  const actualDates = MILESTONES.map(item => parseExcelDate(getCol(row, item.actual)));

  const traces = [
    {
      type: "scatter",
      mode: "lines+markers+text",
      name: "Plan",
      x: planDates,
      y: checkpointNames,
      text: planDates.map(date => date ? date.toLocaleDateString("en-GB") : ""),
      textposition: "top center",
      marker: { color: COLORS.navy, size: 10 },
      line: { color: COLORS.navy, width: 2 },
      connectgaps: false
    },
    {
      type: "scatter",
      mode: "lines+markers+text",
      name: "Actual",
      x: actualDates,
      y: checkpointNames,
      text: actualDates.map(date => date ? date.toLocaleDateString("en-GB") : ""),
      textposition: "bottom center",
      marker: { color: COLORS.coral, size: 10, symbol: "diamond" },
      line: { color: COLORS.coral, width: 2, dash: "dot" },
      connectgaps: false
    }
  ];

  Plotly.react("timelineChart", traces, baseLayout({
    margin: { l: 85, r: 35, t: 30, b: 55 },
    xaxis: { type: "date", showgrid: false, zeroline: false },
    yaxis: { autorange: "reversed", showgrid: false },
    legend: { orientation: "h", x: 0, y: 1.12 }
  }), plotConfig);
}

function formatCell(value) {
  const date = parseExcelDate(value);
  if (date && (value instanceof Date || typeof value === "number")) {
    return date.toLocaleDateString("en-GB");
  }
  return escapeHtml(text(value));
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
