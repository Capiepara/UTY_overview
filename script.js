const APP_VERSION = "v4.1.0";
document.getElementById("appVersion").textContent = APP_VERSION;

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
  const { metrics, weightedRate } = milestoneMetrics(stage);

  const angles = [90, 18, -54, -126, -198];
  const labels = metrics.map(item => `${item.shortName}<br>${item.rate}%`);

  Plotly.react(
    "healthChart",
    [{
      type: "barpolar",
      r: metrics.map(item => Math.max(item.rate, 5)),
      theta: angles,
      width: metrics.map(() => 58),
      marker: {
        color: metrics.map(item => scoreColor(item.rate)),
        line: { color: "#ffffff", width: 5 }
      },
      customdata: metrics.map(item => [item.name, item.rate, item.eligible]),
      hovertemplate:
        "<b>%{customdata[0]}</b><br>" +
        "On time: %{customdata[1]}%<br>" +
        "DEV codes checked: %{customdata[2]}<extra></extra>"
    }],
    baseLayout({
      margin: { l: 12, r: 12, t: 28, b: 12 },
      polar: {
        bgcolor: "rgba(0,0,0,0)",
        radialaxis: {
          visible: false,
          range: [0, 110]
        },
        angularaxis: {
          visible: false,
          rotation: 90,
          direction: "clockwise"
        }
      },
      showlegend: false,
      annotations: [
        {
          text:
            `<b>${stage}</b><br>` +
            `<span style="font-size:34px">${weightedRate}%</span><br>` +
            `<span style="font-size:11px;color:#70808E">on time</span>`,
          x: 0.5,
          y: 0.5,
          showarrow: false,
          align: "center"
        },
        ...metrics.map((item, index) => {
          const rad = angles[index] * Math.PI / 180;
          const radius = 0.41;
          return {
            text: `<b>${item.shortName}</b><br>${item.rate}%`,
            x: 0.5 + radius * Math.cos(rad),
            y: 0.5 + radius * Math.sin(rad),
            showarrow: false,
            font: {
              color: "#ffffff",
              size: 11,
              family: "Nunito, Tahoma, sans-serif"
            },
            align: "center"
          };
        })
      ]
    }),
    plotConfig
  );

  renderHealthSummary(metrics);
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

function formatCell(value) {
  const date = parseExcelDate(value);

  if (
    date &&
    (value instanceof Date || typeof value === "number")
  ) {
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
