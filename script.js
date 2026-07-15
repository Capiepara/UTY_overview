
function formatUTCDate(date){
  if(!date) return "—";
  return date.toLocaleDateString("en-GB",{timeZone:"UTC"});
}

const APP_VERSION = "v4.2.2";
document.getElementById("appVersion").textContent = APP_VERSION;

let workbook = null;
let rawData = [];
let filteredData = [];
let headers = [];
let choiceControls = {};
let searchTerm = "";
let k1Records = [];

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
  {
    name: "CAD received",
    shortName: "CAD",
    plan: ["TP/CAD Handover (Plan)"],
    actual: ["TP/ CAD received (Actual)"]
  },
  {
    name: "Material arrived",
    shortName: "Material",
    plan: ["Material  Arrive (Plan)", "Material Arrive (Plan)"],
    actual: ["Latest  Material Arrive (Actual)", "Latest Material Arrive (Actual)"]
  },
  {
    name: "Pattern locked",
    shortName: "Pattern",
    plan: ["Pattern lock (Plan)"],
    actual: ["Pattern locked (Actual)"]
  },
  {
    name: "Tooling locked",
    shortName: "Tooling",
    plan: ["Tooling lock (Plan)"],
    actual: ["Tooling locked (Actual)"]
  },
  {
    name: "Mold finished",
    shortName: "Mold",
    plan: ["Mold ETC (Plan)"],
    actual: ["Mold finished (Actual)"]
  }
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
document.getElementById("tableSearch").addEventListener("input", event => {
  searchTerm = event.target.value.trim().toLowerCase();
  renderTable();
});

document.querySelectorAll(".main-tab").forEach(button => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".main-tab").forEach(tab => tab.classList.remove("active"));
    document.querySelectorAll(".dashboard-page").forEach(page => page.classList.remove("active-page"));

    button.classList.add("active");
    document.getElementById(button.dataset.page).classList.add("active-page");

    if (button.dataset.page === "qualityPage") {
      updateK1Dashboard();
      setTimeout(() => {
        ["k1SankeyChart", "k1RetryChart", "k1FactoryChart"].forEach(id => {
          const chart = document.getElementById(id);
          if (chart && chart.data) Plotly.Plots.resize(chart);
        });
      }, 30);
    }
  });
});

document.getElementById("k1DetailFilter").addEventListener("change", renderK1DetailTable);

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

  const headerRowIndex = findHeaderRow(rows);

  if (headerRowIndex < 0) {
    alert("The header row could not be found. Please choose the Development Tracking sheet.");
    return;
  }

  headers = rows[headerRowIndex].map(value => String(value ?? "").trim());

  rawData = rows
    .slice(headerRowIndex + 1)
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
  document.getElementById("mainTabs").classList.remove("hidden");
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
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function text(value) {
  return String(value ?? "").trim();
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
  return new Set(
    data.map(row => text(getCol(row, aliases))).filter(Boolean)
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

  const filterConfigs = [
    ["seasonFilter", ["Season"]],
    ["stageFilter", ["DEV. Stage"]],
    ["factoryFilter", ["Factory"]],
    ["materialFilter", ["Material Indicator"]],
    ["modelFilter", ["Model"]]
  ];

  filterConfigs.forEach(([id, aliases]) => {
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

  return data.filter(row =>
    normalize(getCol(row, ["DEV. Stage"])) === target
  );
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
    labels: [
      "LR2 total",
      "LR2 dropped",
      "FLC total",
      "FLC dropped",
      "SMS total",
      "SMS dropped",
      "CFM"
    ],
    values: [lr2, -lr2Drop, flc, -flcDrop, sms, -smsDrop, cfm],
    measures: [
      "absolute",
      "relative",
      "absolute",
      "relative",
      "absolute",
      "relative",
      "absolute"
    ],
    cfm
  };
}

function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const stringValue = text(value);
  if (!stringValue) return null;

  const timestamp = Date.parse(stringValue);
  return Number.isNaN(timestamp) ? null : new Date(timestamp);
}

function semanticStatus(value) {
  const status = normalize(value);

  if (!status) return "blank";

  if (["-", "n/a", "na", "no need", "none", "not required"].includes(status)) {
    return "exempt";
  }

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
    return { score: 1, status: "not-needed" };
  }

  if (actualState === "complete") {
    return { score: 1, status: "on-time" };
  }

  const planDate = parseExcelDate(planRaw);
  const actualDate = parseExcelDate(actualRaw);

  if (planDate && actualDate) {
    const delayDays = Math.ceil((actualDate - planDate) / 86400000);

    return delayDays <= 0
      ? { score: 1, status: "on-time", delayDays }
      : { score: 0, status: "late", delayDays };
  }

  if (!planDate && actualDate) {
    return { score: 1, status: "on-time" };
  }

  if (planDate && !actualDate) {
    return Date.now() > planDate.getTime()
      ? { score: 0, status: "overdue" }
      : { score: null, status: "not-due" };
  }

  if (actualState === "other") {
    return { score: 1, status: "on-time" };
  }

  return { score: null, status: "unknown" };
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
    ? Math.round(
        metrics.reduce((sum, item) => sum + item.rate * item.eligible, 0) /
        totalEligible
      )
    : 0;

  return { metrics, weightedRate };
}

function overallMilestoneRate() {
  const stages = ["LR2", "FLC", "SMS"];
  let weightedPoints = 0;
  let totalEligible = 0;

  stages.forEach(stage => {
    const result = milestoneMetrics(stage);

    result.metrics.forEach(item => {
      weightedPoints += item.rate * item.eligible;
      totalEligible += item.eligible;
    });
  });

  return totalEligible ? Math.round(weightedPoints / totalEligible) : 0;
}

function updateDashboard() {
  const waterfall = getWaterfallData();

  document.getElementById("kpiDev").textContent =
    distinctCount(filteredData, ["DEV. Code"]);

  document.getElementById("kpiHealth").textContent =
    `${overallMilestoneRate()}%`;

  document.getElementById("kpiCfm").textContent = waterfall.cfm;

  document.getElementById("kpiModels").textContent =
    distinctCount(filteredData, ["Model"]);

  drawWaterfall(waterfall);
  drawHealth();
  drawSankey();
  buildMaterialDetailSelectors();
  updateMaterialDetails();
  drawHeatmap();
  renderTable();
  updateK1Dashboard();
}

function baseLayout(extra = {}) {
  return {
    paper_bgcolor: "rgba(0,0,0,0)",
    plot_bgcolor: "rgba(0,0,0,0)",
    font: {
      family: "Nunito, Tahoma, sans-serif",
      color: COLORS.dark,
      size: 12
    },
    margin: { l: 48, r: 20, t: 20, b: 55 },
    ...extra
  };
}

const plotConfig = {
  responsive: true,
  displayModeBar: false
};

function drawWaterfall(data) {
  Plotly.react(
    "waterfallChart",
    [{
      type: "waterfall",
      orientation: "v",
      measure: data.measures,
      x: data.labels,
      y: data.values,
      text: data.values.map(value => String(value)),
      textposition: "outside",
      textfont: {
        color: COLORS.dark,
        size: 12,
        family: "Nunito, Tahoma, sans-serif"
      },
      cliponaxis: false,
      connector: {
        line: { color: "#9DA8B1", width: 1 }
      },
      increasing: {
        marker: {
          color: "#4A6FA5",
          line: { color: "#ffffff", width: 1 }
        }
      },
      decreasing: {
        marker: {
          color: "#F28482",
          line: { color: "#ffffff", width: 1 }
        }
      },
      totals: {
        marker: {
          color: "#43AA8B",
          line: { color: "#ffffff", width: 1 }
        }
      },
      hovertemplate: "%{x}<br><b>%{y}</b> DEV codes<extra></extra>"
    }],
    baseLayout({
      margin: { l: 18, r: 18, t: 32, b: 58 },
      showlegend: false,
      xaxis: {
        fixedrange: true,
        tickangle: -18,
        showgrid: false,
        zeroline: false,
        showline: false
      },
      yaxis: {
        fixedrange: true,
        showgrid: false,
        zeroline: false,
        showline: false,
        showticklabels: false,
        ticks: ""
      }
    }),
    plotConfig
  );
}

function scoreColor(rate) {
  if (rate >= 90) return COLORS.teal;
  if (rate >= 75) return COLORS.cyan;
  if (rate >= 60) return COLORS.orange;
  return COLORS.coral;
}

function drawHealth() {
  const stage = document.getElementById("healthStage").value;

  if (stage === "CFM") {
    drawCFMHealth();
    return;
  }

  document.getElementById("healthTitle").textContent = "Plan compared with actual";
  document.getElementById("healthSummary").style.display = "";

  const { metrics, weightedRate } = milestoneMetrics(stage);

  const width = 820;
  const height = 500;
  const cx = 245;
  const cy = 250;
  const outerR = 190;
  const innerR = 62;

  const startAngle = -90;
  const totalSweep = 180;
  const gap = 2;
  const segmentSweep = (totalSweep - gap * 4) / 5;

  const calloutY = [55, 145, 240, 335, 425];
  const calloutX = 525;

  function pointOnCircle(radius, angleDeg) {
    const radians = angleDeg * Math.PI / 180;
    return {
      x: cx + radius * Math.cos(radians),
      y: cy + radius * Math.sin(radians)
    };
  }

  function arcPath(innerRadius, outerRadius, angle1, angle2) {
    const p1 = pointOnCircle(outerRadius, angle1);
    const p2 = pointOnCircle(outerRadius, angle2);
    const p3 = pointOnCircle(innerRadius, angle2);
    const p4 = pointOnCircle(innerRadius, angle1);
    const largeArc = Math.abs(angle2 - angle1) > 180 ? 1 : 0;

    return [
      `M ${p1.x.toFixed(2)} ${p1.y.toFixed(2)}`,
      `A ${outerRadius} ${outerRadius} 0 ${largeArc} 1 ${p2.x.toFixed(2)} ${p2.y.toFixed(2)}`,
      `L ${p3.x.toFixed(2)} ${p3.y.toFixed(2)}`,
      `A ${innerRadius} ${innerRadius} 0 ${largeArc} 0 ${p4.x.toFixed(2)} ${p4.y.toFixed(2)}`,
      "Z"
    ].join(" ");
  }

  const segmentSvg = metrics.map((item, index) => {
    const a1 = startAngle + index * (segmentSweep + gap);
    const a2 = a1 + segmentSweep;
    const mid = (a1 + a2) / 2;

    const path = arcPath(innerR, outerR, a1, a2);
    const labelPoint = pointOnCircle((innerR + outerR) / 2, mid);
    const connectorStart = pointOnCircle(outerR + 5, mid);
    const elbowX = 455;
    const targetY = calloutY[index];
    const color = scoreColor(item.rate);

    const shortLabel = item.shortName === "Material" ? "MATERIAL"
      : item.shortName === "Pattern" ? "PATTERN"
      : item.shortName === "Tooling" ? "TOOLING"
      : item.shortName === "Mold" ? "MOLD"
      : "CAD";

    return `
      <path class="health-segment"
            d="${path}"
            fill="${color}">
        <title>${escapeHtml(item.name)} — ${item.rate}% on time — ${item.eligible} DEV codes checked</title>
      </path>

      <text class="health-segment-label"
            x="${labelPoint.x.toFixed(1)}"
            y="${(labelPoint.y - 3).toFixed(1)}">${shortLabel}</text>
      <text class="health-segment-value"
            x="${labelPoint.x.toFixed(1)}"
            y="${(labelPoint.y + 14).toFixed(1)}">${item.rate}%</text>

      <polyline class="health-connector"
                points="${connectorStart.x.toFixed(1)},${connectorStart.y.toFixed(1)}
                        ${elbowX},${connectorStart.y.toFixed(1)}
                        ${elbowX + 18},${targetY}" />

      <circle class="health-dot-outer" cx="${connectorStart.x.toFixed(1)}" cy="${connectorStart.y.toFixed(1)}" r="6"></circle>
      <circle class="health-dot-inner" cx="${connectorStart.x.toFixed(1)}" cy="${connectorStart.y.toFixed(1)}" r="3"></circle>

      <circle class="health-dot-outer" cx="${elbowX + 18}" cy="${targetY}" r="6"></circle>
      <circle class="health-dot-inner" cx="${elbowX + 18}" cy="${targetY}" r="3"></circle>

      <line class="health-callout-line"
            x1="${calloutX}"
            y1="${targetY + 17}"
            x2="785"
            y2="${targetY + 17}" />

      <text class="health-callout-title" x="${calloutX}" y="${targetY}">
        ${escapeHtml(item.name.toUpperCase())}
      </text>
      <text class="health-callout-value" x="${calloutX}" y="${targetY + 34}">
        ${item.rate}% ON TIME
      </text>
      <text class="health-callout-desc" x="${calloutX}" y="${targetY + 51}">
        ${item.eligible} DEV codes checked
      </text>
    `;
  }).join("");

  document.getElementById("healthChart").innerHTML = `
    <div class="health-svg-wrap">
      <svg class="health-svg"
           viewBox="0 0 ${width} ${height}"
           role="img"
           aria-label="${escapeHtml(stage)} on-time infographic">

        ${segmentSvg}

        <circle class="health-center-disc" cx="${cx}" cy="${cy}" r="84"></circle>
        <circle class="health-center-inner" cx="${cx}" cy="${cy}" r="63"></circle>

        <text class="health-center-stage" x="${cx}" y="${cy - 31}">
          ${escapeHtml(stage)}
        </text>
        <text class="health-center-value" x="${cx}" y="${cy + 10}">
          ${weightedRate}%
        </text>
        <text class="health-center-caption" x="${cx}" y="${cy + 37}">
          OVERALL ON TIME
        </text>
      </svg>
    </div>
  `;

  renderHealthSummary(metrics);
}

function getCFMOutcome(row) {
  const planRaw = getCol(row, ["CFM SMPL ETD (Plan)"]);
  const actualRaw = getCol(row, ["CFM SMPL Sent (Actual)"]);

  const planStatus = semanticStatus(planRaw);
  const actualStatus = semanticStatus(actualRaw);

  if (planStatus === "exempt" || actualStatus === "exempt") {
    return {
      score: 1,
      status: "on-time",
      delayDays: 0,
      planDate: null,
      actualDate: null
    };
  }

  const planDate = parseExcelDate(planRaw);
  const actualDate = parseExcelDate(actualRaw);

  if (planDate && actualDate) {
    const delayDays = Math.ceil((actualDate - planDate) / 86400000);

    return {
      score: delayDays <= 0 ? 1 : 0,
      status: delayDays <= 0 ? "on-time" : "late",
      delayDays: Math.max(delayDays, 0),
      planDate,
      actualDate
    };
  }

  if (planDate && !actualDate) {
    const overdueDays = Math.ceil((Date.now() - planDate.getTime()) / 86400000);

    return {
      score: overdueDays > 0 ? 0 : null,
      status: overdueDays > 0 ? "overdue" : "pending",
      delayDays: Math.max(overdueDays, 0),
      planDate,
      actualDate: null
    };
  }

  if (!planDate && actualDate) {
    return {
      score: 1,
      status: "on-time",
      delayDays: 0,
      planDate: null,
      actualDate
    };
  }

  return {
    score: null,
    status: "pending",
    delayDays: null,
    planDate: null,
    actualDate: null
  };
}

function drawCFMHealth() {
  document.getElementById("healthTitle").textContent =
    "CFM sample shipment performance";
  document.getElementById("healthSummary").style.display = "none";

  const rows = stageRows("CFM");
  const outcomesByDev = new Map();

  rows.forEach(row => {
    const devCode = text(getCol(row, ["DEV. Code"]));
    if (!devCode) return;

    const outcome = getCFMOutcome(row);
    const previous = outcomesByDev.get(devCode);

    // Keep the most critical outcome when duplicated.
    if (
      !previous ||
      outcome.score === 0 ||
      (previous.score === null && outcome.score !== null)
    ) {
      outcomesByDev.set(devCode, outcome);
    }
  });

  const outcomes = [...outcomesByDev.values()];
  const eligible = outcomes.filter(item => item.score !== null);
  const onTime = eligible.filter(item => item.score === 1).length;
  const late = eligible.filter(item => item.score === 0).length;
  const pending = outcomes.filter(item => item.score === null).length;

  const rate = eligible.length
    ? Math.round(onTime / eligible.length * 100)
    : 0;

  const lateDays = eligible
    .filter(item => item.score === 0 && Number.isFinite(item.delayDays))
    .map(item => item.delayDays);

  const averageDelay = lateDays.length
    ? Math.round(
        lateDays.reduce((sum, value) => sum + value, 0) /
        lateDays.length
      )
    : 0;

  const planDates = outcomes
    .map(item => item.planDate)
    .filter(Boolean)
    .sort((a, b) => a - b);

  const actualDates = outcomes
    .map(item => item.actualDate)
    .filter(Boolean)
    .sort((a, b) => a - b);

  const middleDate = dates => {
    if (!dates.length) return null;
    return dates[Math.floor(dates.length / 2)];
  };

  const medianPlan = middleDate(planDates);
  const medianActual = middleDate(actualDates);

  const color = scoreColor(rate);
  const angle = Math.max(0, Math.min(rate, 100)) * 3.6;

  if (!outcomes.length) {
    document.getElementById("healthChart").innerHTML = `
      <div class="cfm-empty">
        <strong>No CFM sample data is available</strong>
        <span>Expected columns: CFM SMPL ETD (Plan) and CFM SMPL Sent (Actual)</span>
      </div>
    `;
    return;
  }

  document.getElementById("healthChart").innerHTML = `
    <div class="cfm-health-view">
      <div class="cfm-gauge"
           style="--cfm-color:${color};--cfm-angle:${angle}deg">
        <div class="cfm-gauge-center">
          <strong>${rate}%</strong>
          <span>CFM sample on time</span>
        </div>
      </div>

      <div class="cfm-summary">
        <div class="cfm-summary-card" style="border-left-color:${COLORS.navy}">
          <span>DEV codes checked</span>
          <strong>${outcomes.length}</strong>
        </div>

        <div class="cfm-summary-card" style="border-left-color:${COLORS.cyan}">
          <span>Median planned ETD</span>
          <strong>${medianPlan ? medianPlan.toLocaleDateString("en-GB",{timeZone:"UTC"}) : "—"}</strong>
        </div>

        <div class="cfm-summary-card" style="border-left-color:${COLORS.teal}">
          <span>Median actual sent</span>
          <strong>${medianActual ? medianActual.toLocaleDateString("en-GB",{timeZone:"UTC"}) : "—"}</strong>
        </div>

        <div class="cfm-summary-card" style="border-left-color:${COLORS.coral}">
          <span>Average late days</span>
          <strong>${averageDelay} days</strong>
        </div>

        <div class="cfm-status-row">
          <div class="cfm-status-box" style="border-top-color:${COLORS.teal}">
            <strong>${onTime}</strong>
            <span>On time</span>
          </div>
          <div class="cfm-status-box" style="border-top-color:${COLORS.coral}">
            <strong>${late}</strong>
            <span>Late</span>
          </div>
          <div class="cfm-status-box" style="border-top-color:${COLORS.orange}">
            <strong>${pending}</strong>
            <span>Pending</span>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderHealthSummary(metrics) {
  const container = document.getElementById("healthSummary");

  container.innerHTML = metrics.map(item => `
    <div class="health-item" style="border-left-color:${scoreColor(item.rate)}">
      <strong>${item.rate}%</strong>
      <span>${escapeHtml(item.name)}</span>
    </div>
  `).join("");
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
        rows.filter(row =>
          text(getCol(row, ["Material Indicator"])) === material
        ),
        ["DEV. Code"]
      );

      if (count > 0) links.push({ stage, material, count });
    });
  });

  return { stages, materials, links };
}

function drawSankey() {
  const { stages, materials, links } = buildMaterialLinks();
  const labels = [...stages, ...materials];
  const stageColors = [
    COLORS.orange,
    COLORS.coral,
    COLORS.cyan,
    COLORS.navy
  ];
  const materialColors = materials.map((_, index) =>
    [COLORS.teal, "#88C8BF", "#D8B767"][index % 3]
  );

  Plotly.react(
    "sankeyChart",
    [{
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
        source: links.map(item => labels.indexOf(item.stage)),
        target: links.map(item => labels.indexOf(item.material)),
        value: links.map(item => item.count),
        color: links.map(item =>
          hexToRgba(stageColors[stages.indexOf(item.stage)], 0.42)
        ),
        customdata: links.map(item => [item.stage, item.material]),
        hovertemplate:
          "%{source.label} → %{target.label}<br>" +
          "<b>%{value}</b> DEV codes<extra></extra>"
      }
    }],
    baseLayout({
      margin: { l: 10, r: 10, t: 18, b: 18 }
    }),
    plotConfig
  ).then(() => {
    const chart = document.getElementById("sankeyChart");

    if (typeof chart.removeAllListeners === "function") {
      chart.removeAllListeners("plotly_click");
    }

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

  stageSelect.innerHTML = stages
    .map(stage => `<option value="${escapeHtml(stage)}">${escapeHtml(stage)}</option>`)
    .join("");

  materialSelect.innerHTML = materials
    .map(material => `<option value="${escapeHtml(material)}">${escapeHtml(material)}</option>`)
    .join("");

  if (stages.includes(oldStage)) stageSelect.value = oldStage;
  if (materials.includes(oldMaterial)) materialSelect.value = oldMaterial;
}

function updateMaterialDetails() {
  const stage = document.getElementById("detailStage").value;
  const material = document.getElementById("detailMaterial").value;
  const list = document.getElementById("materialDetailList");

  if (!stage || !material) {
    document.getElementById("materialDetailCount").textContent = "0";
    list.innerHTML =
      '<div class="material-detail-empty">No matching group is available.</div>';
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
              ? `<br><span style="color:#70808E">${escapeHtml(item.model)}</span>`
              : ""}
          </div>
        </div>
      `).join("")
    : '<div class="material-detail-empty">No DEV codes match this stage and material indicator.</div>';
}

function hexToRgba(hex, alpha) {
  const value = hex.replace("#", "");
  const number = parseInt(value, 16);

  return `rgba(${(number >> 16) & 255},${(number >> 8) & 255},${number & 255},${alpha})`;
}

function drawHeatmap() {
  const models = uniqueValues(filteredData, ["Model"])
    .map(model => ({
      model,
      count: distinctCount(
        filteredData.filter(row =>
          text(getCol(row, ["Model"])) === model
        ),
        ["DEV. Code"]
      )
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 20)
    .map(item => item.model);

  const stages = ["LR2", "FLC", "SMS", "CFM"];

  const percentages = models.map(model => {
    const modelRows = filteredData.filter(row =>
      text(getCol(row, ["Model"])) === model
    );

    const counts = stages.map(stage =>
      distinctCount(stageRows(stage, modelRows), ["DEV. Code"])
    );

    const baseTotal =
      counts[0] + counts[1] + counts[2] || 1;

    return counts.map(value =>
      Math.round(value / baseTotal * 100)
    );
  });

  Plotly.react(
    "heatmapChart",
    [{
      type: "heatmap",
      x: stages,
      y: models,
      z: percentages,
      zmin: 0,
      zmax: 100,
      colorscale: [
        [0, "#FFF3E4"],
        [0.35, "#F2C377"],
        [0.65, "#66BBA2"],
        [1, "#405B73"]
      ],
      text: percentages.map(row =>
        row.map(value => `${value}%`)
      ),
      texttemplate: "%{text}",
      textfont: { size: 11 },
      hovertemplate:
        "%{y}<br>%{x}: <b>%{z}%</b><extra></extra>",
      colorbar: {
        title: "Share",
        ticksuffix: "%",
        thickness: 12
      }
    }],
    baseLayout({
      margin: { l: 120, r: 55, t: 20, b: 45 },
      xaxis: {
        side: "top",
        fixedrange: true,
        showgrid: false
      },
      yaxis: {
        fixedrange: true,
        autorange: "reversed",
        tickfont: { size: 10 }
      }
    }),
    plotConfig
  );
}

function renderTable() {
  const table = document.getElementById("detailTable");

  const displayHeaders = [
    "Season",
    "Active/Drop",
    "DEV. Stage",
    "DEV. Code",
    "Factory",
    "Model",
    "Gender / Model",
    "Material Indicator",
    "TP/CAD Handover (Plan)",
    "TP/ CAD received (Actual)",
    "Material  Arrive (Plan)",
    "Latest  Material Arrive (Actual)",
    "Pattern lock (Plan)",
    "Pattern locked (Actual)",
    "Tooling lock (Plan)",
    "Tooling locked (Actual)",
    "Mold ETC (Plan)",
    "Mold finished (Actual)",
    "Shipment ontime"
  ];

  const rows = filteredData
    .filter(row => {
      if (!searchTerm) return true;

      return [
        getCol(row, ["DEV. Code"]),
        getCol(row, ["Model"]),
        getCol(row, ["Factory"]),
        getCol(row, ["Season"])
      ].some(value =>
        normalize(value).includes(searchTerm)
      );
    })
    .slice(0, 500);

  table.innerHTML = `
    <thead>
      <tr>
        ${displayHeaders.map(header =>
          `<th>${escapeHtml(header)}</th>`
        ).join("")}
      </tr>
    </thead>
    <tbody>
      ${rows.map(row => `
        <tr>
          ${displayHeaders.map(header =>
            `<td>${formatCell(getCol(row, [header]))}</td>`
          ).join("")}
        </tr>
      `).join("")}
    </tbody>
  `;
}

function normalizeK1Result(value) {
  const result = normalize(value);

  if (!result || result === "-" || result === "n/a" || result === "no need") {
    return "";
  }

  if (result.includes("pass")) return "Pass";
  if (result.includes("fail")) return "Fail";
  if (result.includes("pending") || result.includes("waiting")) return "Pending";

  return text(value);
}

function k1FieldHasValue(value) {
  const result = normalize(value);
  return Boolean(result && result !== "-" && result !== "n/a" && result !== "no need");
}

function buildK1Records() {
  const byDevCode = new Map();

  filteredData.forEach(row => {
    const devCode = text(getCol(row, ["DEV. Code"]));
    if (!devCode) return;

    const plan = getCol(row, ["K1 (Plan)"]);
    const ship1 = getCol(row, ["K1 Ship (Actual)"]);
    const result1 = normalizeK1Result(getCol(row, ["K1 result"]));

    const ship2 = getCol(row, ["K1 re-fit Ship"]);
    const result2 = normalizeK1Result(
      getCol(row, ["K1 resut (2nd)", "K1 result (2nd)"])
    );

    const ship3 = getCol(row, ["K1 re-fit Ship2"]);
    const result3 = normalizeK1Result(
      getCol(row, ["K1 resut (3rd)", "K1 result (3rd)"])
    );

    const hasActivity = [
      plan, ship1, result1, ship2, result2, ship3, result3
    ].some(k1FieldHasValue);

    if (!hasActivity) return;

    let attempts = 0;
    if (k1FieldHasValue(ship1) || result1) attempts = 1;
    if (k1FieldHasValue(ship2) || result2) attempts = 2;
    if (k1FieldHasValue(ship3) || result3) attempts = 3;

    let passRound = null;
    if (result1 === "Pass") passRound = 1;
    else if (result2 === "Pass") passRound = 2;
    else if (result3 === "Pass") passRound = 3;

    const record = {
      devCode,
      season: text(getCol(row, ["Season"])),
      stage: text(getCol(row, ["DEV. Stage"])),
      factory: text(getCol(row, ["Factory"])),
      model: text(getCol(row, ["Model"])),
      genderModel: text(getCol(row, ["Gender / Model"])),
      material: text(getCol(row, ["Material Indicator"])),
      plan,
      ship1,
      result1,
      ship2,
      result2,
      ship3,
      result3,
      attempts: Math.max(attempts, passRound || 0, 1),
      passRound,
      firstFailed: result1 === "Fail",
      stillOpen: !passRound
    };

    const previous = byDevCode.get(devCode);

    // Keep the row containing the furthest K1 attempt.
    if (!previous || record.attempts >= previous.attempts) {
      byDevCode.set(devCode, record);
    }
  });

  return [...byDevCode.values()].sort((a, b) =>
    a.devCode.localeCompare(b.devCode, undefined, { numeric: true })
  );
}

function updateK1Dashboard() {
  k1Records = buildK1Records();

  const total = k1Records.length;
  const firstPass = k1Records.filter(item => item.passRound === 1).length;
  const needRefit = k1Records.filter(item => item.firstFailed).length;
  const stillOpen = k1Records.filter(item => item.stillOpen).length;
  const averageAttempts = total
    ? k1Records.reduce((sum, item) => sum + item.attempts, 0) / total
    : 0;

  document.getElementById("k1FirstPass").textContent =
    total ? `${Math.round(firstPass / total * 100)}%` : "0%";

  document.getElementById("k1AverageAttempts").textContent =
    averageAttempts.toFixed(1);

  document.getElementById("k1NeedRefit").textContent = needRefit;
  document.getElementById("k1StillOpen").textContent = stillOpen;

  drawK1Sankey();
  drawK1RetryChart();
  drawK1FactoryChart();
  renderK1DetailTable();
}

function drawK1Sankey() {
  const count = predicate => k1Records.filter(predicate).length;

  const pass1 = count(item => item.passRound === 1);
  const fail1 = count(item => item.result1 === "Fail");
  const pending1 = count(item => !item.result1 || item.result1 === "Pending");

  const pass2 = count(item => item.passRound === 2);
  const fail2 = count(item => item.result1 === "Fail" && item.result2 === "Fail");
  const waiting2 = count(item =>
    item.result1 === "Fail" && !item.result2 && !item.passRound
  );

  const pass3 = count(item => item.passRound === 3);
  const fail3 = count(item =>
    item.result1 === "Fail" &&
    item.result2 === "Fail" &&
    item.result3 === "Fail"
  );
  const waiting3 = count(item =>
    item.result1 === "Fail" &&
    item.result2 === "Fail" &&
    !item.result3 &&
    !item.passRound
  );

  const labels = [
    "K1 started",
    "Pass round 1",
    "Fail round 1",
    "Waiting round 1",
    "Pass round 2",
    "Fail round 2",
    "Waiting round 2",
    "Pass round 3",
    "Fail round 3",
    "Waiting round 3"
  ];

  const links = [];

  function add(source, target, value, color) {
    if (value > 0) links.push({ source, target, value, color });
  }

  add(0, 1, pass1, hexToRgba(COLORS.teal, 0.56));
  add(0, 2, fail1, hexToRgba(COLORS.coral, 0.48));
  add(0, 3, pending1, hexToRgba(COLORS.orange, 0.44));

  add(2, 4, pass2, hexToRgba(COLORS.teal, 0.56));
  add(2, 5, fail2, hexToRgba(COLORS.coral, 0.48));
  add(2, 6, waiting2, hexToRgba(COLORS.orange, 0.44));

  add(5, 7, pass3, hexToRgba(COLORS.teal, 0.56));
  add(5, 8, fail3, hexToRgba(COLORS.coral, 0.48));
  add(5, 9, waiting3, hexToRgba(COLORS.orange, 0.44));

  Plotly.react(
    "k1SankeyChart",
    [{
      type: "sankey",
      arrangement: "snap",
      node: {
        label: labels,
        color: [
          COLORS.navy,
          COLORS.teal,
          COLORS.coral,
          COLORS.orange,
          COLORS.teal,
          COLORS.coral,
          COLORS.orange,
          COLORS.teal,
          COLORS.coral,
          COLORS.orange
        ],
        pad: 22,
        thickness: 20,
        line: { color: "#ffffff", width: 1 }
      },
      link: {
        source: links.map(item => item.source),
        target: links.map(item => item.target),
        value: links.map(item => item.value),
        color: links.map(item => item.color),
        hovertemplate:
          "%{source.label} → %{target.label}<br>" +
          "<b>%{value}</b> DEV codes<extra></extra>"
      }
    }],
    baseLayout({
      margin: { l: 10, r: 10, t: 18, b: 18 }
    }),
    plotConfig
  );
}

function drawK1RetryChart() {
  const data = [
    {
      label: "Pass round 1",
      value: k1Records.filter(item => item.passRound === 1).length,
      color: COLORS.teal
    },
    {
      label: "Pass round 2",
      value: k1Records.filter(item => item.passRound === 2).length,
      color: COLORS.cyan
    },
    {
      label: "Pass round 3",
      value: k1Records.filter(item => item.passRound === 3).length,
      color: COLORS.navy
    },
    {
      label: "Still open",
      value: k1Records.filter(item => item.stillOpen).length,
      color: COLORS.coral
    }
  ].reverse();

  Plotly.react(
    "k1RetryChart",
    [{
      type: "bar",
      orientation: "h",
      y: data.map(item => item.label),
      x: data.map(item => item.value),
      text: data.map(item => String(item.value)),
      textposition: "outside",
      marker: { color: data.map(item => item.color) },
      hovertemplate: "%{y}<br><b>%{x}</b> DEV codes<extra></extra>"
    }],
    baseLayout({
      margin: { l: 105, r: 35, t: 20, b: 35 },
      xaxis: { showgrid: false, zeroline: false, rangemode: "tozero" },
      yaxis: { showgrid: false }
    }),
    plotConfig
  );
}

function drawK1FactoryChart() {
  const factoryMap = new Map();

  k1Records.forEach(item => {
    if (!item.factory) return;
    if (!factoryMap.has(item.factory)) factoryMap.set(item.factory, []);
    factoryMap.get(item.factory).push(item);
  });

  const data = [...factoryMap.entries()]
    .map(([factory, records]) => ({
      factory,
      firstPassRate: Math.round(
        records.filter(item => item.passRound === 1).length /
        records.length *
        100
      ),
      averageAttempts:
        records.reduce((sum, item) => sum + item.attempts, 0) /
        records.length
    }))
    .sort((a, b) => b.firstPassRate - a.firstPassRate);

  Plotly.react(
    "k1FactoryChart",
    [
      {
        type: "bar",
        x: data.map(item => item.factory),
        y: data.map(item => item.firstPassRate),
        name: "First-pass rate",
        marker: { color: COLORS.teal },
        text: data.map(item => `${item.firstPassRate}%`),
        textposition: "outside",
        hovertemplate:
          "%{x}<br>First-pass rate: <b>%{y}%</b><extra></extra>"
      },
      {
        type: "scatter",
        mode: "lines+markers",
        x: data.map(item => item.factory),
        y: data.map(item => Number(item.averageAttempts.toFixed(2))),
        name: "Average attempts",
        yaxis: "y2",
        marker: { color: COLORS.coral, size: 8 },
        line: { color: COLORS.coral, width: 2 },
        hovertemplate:
          "%{x}<br>Average attempts: <b>%{y}</b><extra></extra>"
      }
    ],
    baseLayout({
      margin: { l: 45, r: 50, t: 28, b: 60 },
      xaxis: { showgrid: false, tickangle: -25 },
      yaxis: {
        title: "First-pass rate",
        range: [0, 110],
        ticksuffix: "%",
        showgrid: false,
        zeroline: false
      },
      yaxis2: {
        title: "Attempts",
        overlaying: "y",
        side: "right",
        range: [
          0,
          Math.max(
            3.2,
            ...data.map(item => item.averageAttempts + 0.5)
          )
        ],
        showgrid: false,
        zeroline: false
      },
      legend: { orientation: "h", x: 0, y: 1.16 }
    }),
    plotConfig
  );
}

function renderK1DetailTable() {
  const filter = document.getElementById("k1DetailFilter").value;

  const rows = k1Records.filter(item => {
    if (filter === "pass1") return item.passRound === 1;
    if (filter === "pass2") return item.passRound === 2;
    if (filter === "pass3") return item.passRound === 3;
    if (filter === "open") return item.stillOpen;
    return true;
  });

  document.getElementById("k1DetailTable").innerHTML = `
    <thead>
      <tr>
        <th>DEV Code</th>
        <th>Factory</th>
        <th>Model</th>
        <th>Attempts</th>
        <th>Round 1</th>
        <th>Round 2</th>
        <th>Round 3</th>
        <th>Latest status</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map(item => {
        const statusText = item.passRound
          ? `Pass round ${item.passRound}`
          : "Still open";

        const statusClass = item.passRound
          ? "k1-pass"
          : item.result3 === "Fail"
            ? "k1-fail"
            : "k1-open";

        return `
          <tr>
            <td>${escapeHtml(item.devCode)}</td>
            <td>${escapeHtml(item.factory)}</td>
            <td>${escapeHtml(item.model || item.genderModel)}</td>
            <td>${item.attempts}</td>
            <td>${escapeHtml(item.result1 || "-")}</td>
            <td>${escapeHtml(item.result2 || "-")}</td>
            <td>${escapeHtml(item.result3 || "-")}</td>
            <td class="${statusClass}">${statusText}</td>
          </tr>
        `;
      }).join("")}
    </tbody>
  `;
}

function formatCell(value) {
  const date = parseExcelDate(value);

  if (
    date &&
    (value instanceof Date || typeof value === "number")
  ) {
    return date.toLocaleDateString("en-GB",{timeZone:"UTC"});
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
