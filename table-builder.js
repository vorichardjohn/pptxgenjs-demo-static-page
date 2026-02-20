const fileInput = document.getElementById("data-file");
const columnConfigEl = document.getElementById("column-config");
const generateButton = document.getElementById("generate-table");
const previewButton = document.getElementById("generate-preview");
const selectAllButton = document.getElementById("select-all-columns");
const deselectAllButton = document.getElementById("deselect-all-columns");
const statusEl = document.getElementById("table-status");
const previewHeadEl = document.getElementById("preview-head");
const previewBodyEl = document.getElementById("preview-body");
const previewTableEl = document.getElementById("preview-table");
const progressEl = document.getElementById("generation-progress");
const progressLabelEl = document.getElementById("progress-label");

const optionIds = [
  "slide-title",
  "file-name",
  "rows-per-slide",
  "preview-slide-count",
  "header-fill",
  "header-text",
  "body-fill",
  "body-text",
  "border-color",
  "header-font-size",
  "body-font-size",
  "header-font-weight",
  "body-font-weight",
  "body-italic",
  "row-height",
  "cell-margin",
  "valign",
  "autofit",
  "include-notes",
  "split-export",
  "max-file-size-mb",
];

let rows = [];
let columns = [];
let draggedColumn = null;

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? "#b91c1c" : "#14532d";
}

function sanitizeHex(value) {
  return value.replace(/[^0-9a-fA-F]/g, "").slice(0, 6).toUpperCase() || "000000";
}

function toNumber(value, fallback) {
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function toPositiveWholeNumber(value, fallback) {
  const normalized = String(value).trim();
  if (!/^\d+$/.test(normalized)) {
    return fallback;
  }
  const num = Number(normalized);
  return Number.isInteger(num) && num > 0 ? num : fallback;
}

function parseCsv(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (lines.length < 2) {
    throw new Error("CSV must include a header row and at least one data row.");
  }

  const splitRow = (line) =>
    line
      .split(",")
      .map((cell) => cell.trim().replace(/^"|"$/g, ""));

  const headers = splitRow(lines[0]);
  const parsedRows = lines.slice(1).map((line) => splitRow(line));

  return {
    columns: headers,
    rows: parsedRows.map((row) => {
      const record = {};
      headers.forEach((header, index) => {
        record[header] = row[index] ?? "";
      });
      return record;
    }),
  };
}

function parseJson(text) {
  const parsed = JSON.parse(text);

  if (Array.isArray(parsed) && parsed.length > 0 && Array.isArray(parsed[0])) {
    const [headerRow, ...dataRows] = parsed;
    return {
      columns: headerRow.map((cell) => String(cell)),
      rows: dataRows.map((row) => {
        const record = {};
        headerRow.forEach((header, index) => {
          record[String(header)] = row[index] ?? "";
        });
        return record;
      }),
    };
  }

  const sourceRows = Array.isArray(parsed) ? parsed : parsed.rows;
  if (!Array.isArray(sourceRows) || sourceRows.length === 0) {
    throw new Error("JSON must be an array of rows or include a non-empty rows array.");
  }

  if (Array.isArray(sourceRows[0])) {
    const [headerRow, ...dataRows] = sourceRows;
    return {
      columns: headerRow.map((cell) => String(cell)),
      rows: dataRows.map((row) => {
        const record = {};
        headerRow.forEach((header, index) => {
          record[String(header)] = row[index] ?? "";
        });
        return record;
      }),
    };
  }

  const cols = Object.keys(sourceRows[0]);
  return {
    columns: cols,
    rows: sourceRows,
  };
}

function getIncludedColumns() {
  return columns.filter((column) => column.included);
}

function getOptions() {
  return {
    title: document.getElementById("slide-title").value.trim() || "Uploaded Data Table",
    filePrefix: document.getElementById("file-name").value.trim() || "pptxgenjs-uploaded-table",
    rowsPerSlide: Math.max(1, toNumber(document.getElementById("rows-per-slide").value, 20)),
    previewSlideCount: toPositiveWholeNumber(document.getElementById("preview-slide-count").value, 2),
    headerFill: sanitizeHex(document.getElementById("header-fill").value),
    headerText: sanitizeHex(document.getElementById("header-text").value),
    bodyFill: sanitizeHex(document.getElementById("body-fill").value),
    bodyText: sanitizeHex(document.getElementById("body-text").value),
    borderColor: sanitizeHex(document.getElementById("border-color").value),
    headerFontSize: toNumber(document.getElementById("header-font-size").value, 13),
    bodyFontSize: toNumber(document.getElementById("body-font-size").value, 12),
    headerBold: document.getElementById("header-font-weight").value === "bold",
    bodyBold: document.getElementById("body-font-weight").value === "bold",
    bodyItalic: document.getElementById("body-italic").value === "true",
    rowH: toNumber(document.getElementById("row-height").value, 0.45),
    margin: toNumber(document.getElementById("cell-margin").value, 0.04),
    valign: document.getElementById("valign").value,
    autoFit: document.getElementById("autofit").value === "true",
    includeNotes: document.getElementById("include-notes").value === "true",
    splitExport: document.getElementById("split-export").value === "true",
    maxFileSizeMb: Math.max(1, toNumber(document.getElementById("max-file-size-mb").value, 8)),
  };
}

function applyPreviewStyles() {
  const options = getOptions();
  previewTableEl.style.borderColor = `#${options.borderColor}`;

  previewHeadEl.querySelectorAll("th").forEach((th) => {
    th.style.backgroundColor = `#${options.headerFill}`;
    th.style.color = `#${options.headerText}`;
    th.style.fontSize = `${options.headerFontSize}px`;
    th.style.fontWeight = options.headerBold ? "700" : "400";
  });

  previewBodyEl.querySelectorAll("td").forEach((td) => {
    td.style.backgroundColor = `#${options.bodyFill}`;
    td.style.color = `#${options.bodyText}`;
    td.style.fontSize = `${options.bodyFontSize}px`;
    td.style.fontWeight = options.bodyBold ? "700" : "400";
    td.style.fontStyle = options.bodyItalic ? "italic" : "normal";
  });
}

function renderPreview() {
  previewHeadEl.innerHTML = "";
  previewBodyEl.innerHTML = "";

  const activeColumns = getIncludedColumns();
  if (!activeColumns.length || !rows.length) {
    return;
  }

  const orderedCols = activeColumns.map((col) => col.name);
  const headRow = document.createElement("tr");

  orderedCols.forEach((name) => {
    const th = document.createElement("th");
    const col = activeColumns.find((item) => item.name === name);
    th.textContent = name;
    th.style.textAlign = col?.align || "left";
    headRow.appendChild(th);
  });
  previewHeadEl.appendChild(headRow);

  rows.slice(0, 10).forEach((row) => {
    const tr = document.createElement("tr");
    orderedCols.forEach((name) => {
      const td = document.createElement("td");
      const col = activeColumns.find((item) => item.name === name);
      td.textContent = row[name] ?? "";
      td.style.textAlign = col?.align || "left";
      tr.appendChild(td);
    });
    previewBodyEl.appendChild(tr);
  });

  applyPreviewStyles();
}

function onDragStart(event) {
  draggedColumn = event.currentTarget.dataset.column;
  event.dataTransfer.effectAllowed = "move";
}

function onDragOver(event) {
  event.preventDefault();
  event.dataTransfer.dropEffect = "move";
}

function onDrop(event) {
  event.preventDefault();
  const targetColumn = event.currentTarget.dataset.column;

  if (!draggedColumn || draggedColumn === targetColumn) {
    return;
  }

  const fromIndex = columns.findIndex((col) => col.name === draggedColumn);
  const toIndex = columns.findIndex((col) => col.name === targetColumn);
  const [moved] = columns.splice(fromIndex, 1);
  columns.splice(toIndex, 0, moved);

  renderColumnConfig();
  renderPreview();
}

function renderColumnConfig() {
  columnConfigEl.innerHTML = "";

  columns.forEach((column) => {
    const item = document.createElement("li");
    item.className = "column-item";
    item.draggable = true;
    item.dataset.column = column.name;
    item.addEventListener("dragstart", onDragStart);
    item.addEventListener("dragover", onDragOver);
    item.addEventListener("drop", onDrop);

    const name = document.createElement("strong");
    name.textContent = column.name;

    const includeLabel = document.createElement("label");
    includeLabel.className = "include-toggle";
    includeLabel.textContent = "Include";

    const includeInput = document.createElement("input");
    includeInput.type = "checkbox";
    includeInput.checked = column.included;
    includeInput.setAttribute("aria-label", `Include ${column.name}`);
    includeInput.addEventListener("change", () => {
      column.included = includeInput.checked;
      renderPreview();
    });
    includeLabel.prepend(includeInput);

    const widthInput = document.createElement("input");
    widthInput.type = "number";
    widthInput.min = "0.5";
    widthInput.max = "4";
    widthInput.step = "0.1";
    widthInput.value = String(column.width);
    widthInput.setAttribute("aria-label", `${column.name} width`);
    widthInput.addEventListener("input", () => {
      column.width = toNumber(widthInput.value, 1.5);
    });

    const alignSelect = document.createElement("select");
    ["left", "center", "right"].forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      if (value === column.align) {
        option.selected = true;
      }
      alignSelect.appendChild(option);
    });
    alignSelect.addEventListener("change", () => {
      column.align = alignSelect.value;
      renderPreview();
    });

    item.append(name, includeLabel, widthInput, alignSelect);
    columnConfigEl.appendChild(item);
  });
}

function initializeColumns(names) {
  columns = names.map((name) => ({
    name,
    included: true,
    width: 13.0 / Math.max(names.length, 1),
    align: "left",
  }));

  renderColumnConfig();
  renderPreview();

  generateButton.disabled = false;
  previewButton.disabled = false;
  setColumnBulkLinksEnabled(true);
}

async function handleFileUpload() {
  const [file] = fileInput.files;
  if (!file) {
    return;
  }

  try {
    const text = await file.text();
    const extension = file.name.split(".").pop()?.toLowerCase();
    const result = extension === "csv" ? parseCsv(text) : parseJson(text);

    rows = result.rows;
    initializeColumns(result.columns);
    setStatus(`Loaded ${rows.length} rows and ${columns.length} columns from ${file.name}.`);
  } catch (error) {
    console.error(error);
    rows = [];
    columns = [];
    renderColumnConfig();
    renderPreview();
    generateButton.disabled = true;
    previewButton.disabled = true;
    setColumnBulkLinksEnabled(false);
    setStatus(error.message || "Failed to parse file.", true);
  }
}

function buildTableRows(dataRows, activeColumns, options) {
  const header = activeColumns.map((column) => ({
    text: column.name,
    options: {
      bold: options.headerBold,
      color: options.headerText,
      align: column.align,
      fill: options.headerFill,
      fontSize: options.headerFontSize,
    },
  }));

  const body = dataRows.map((row) =>
    activeColumns.map((column) => ({
      text: String(row[column.name] ?? ""),
      options: {
        align: column.align,
        bold: options.bodyBold,
        italic: options.bodyItalic,
        color: options.bodyText,
        fill: options.bodyFill,
        fontSize: options.bodyFontSize,
      },
    }))
  );

  return [header, ...body];
}

function chunkRows(dataRows, chunkSize) {
  const chunks = [];
  for (let i = 0; i < dataRows.length; i += chunkSize) {
    chunks.push(dataRows.slice(i, i + chunkSize));
  }
  return chunks;
}

function estimateDeckSizeMb(totalRows, activeColumnsCount, slidesCount) {
  const roughBytes = totalRows * activeColumnsCount * 35 + slidesCount * 18000 + 75000;
  return roughBytes / (1024 * 1024);
}

function updateProgress(value, label, animated = false) {
  progressEl.value = Math.max(0, Math.min(100, value));
  progressEl.classList.toggle("progress-animated", animated);
  progressLabelEl.textContent = label;
}

function setColumnBulkLinksEnabled(enabled) {
  const links = [selectAllButton, deselectAllButton];
  links.forEach((link) => {
    link.setAttribute("aria-disabled", String(!enabled));
    link.classList.toggle("disabled-link", !enabled);
  });
}

async function generateDeckFile({ fileName, title, activeColumns, slideRowChunks, options, deckIndex, deckCount }) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "PptxGenJS Table Builder";
  pptx.subject = "Uploaded data to PowerPoint table";

  const columnWidths = activeColumns.map((column) => toNumber(column.width, 1.2));

  slideRowChunks.forEach((slideRows, index) => {
    const slide = pptx.addSlide();
    slide.addText(`${title}${deckCount > 1 ? ` (Part ${deckIndex} of ${deckCount})` : ""}`, {
      x: 0.5,
      y: 0.2,
      w: 12.2,
      h: 0.45,
      bold: true,
      fontSize: 20,
      color: "0F172A",
    });

    slide.addText(`Slide ${index + 1} of ${slideRowChunks.length}`, {
      x: 10.4,
      y: 0.22,
      w: 2,
      h: 0.3,
      fontSize: 10,
      color: "64748B",
      align: "right",
    });

    const tableRows = buildTableRows(slideRows, activeColumns, options);

    slide.addTable(tableRows, {
      x: 0.5,
      y: 0.85,
      w: 12.3,
      colW: columnWidths,
      rowH: options.rowH,
      fontSize: options.bodyFontSize,
      color: options.bodyText,
      fill: options.bodyFill,
      border: { pt: 1, color: options.borderColor },
      valign: options.valign,
      autoFit: options.autoFit,
      margin: options.margin,
    });

    if (options.includeNotes) {
      slide.addNotes(`
[Notes]
- Generated from uploaded data.
- Column order, inclusion, widths, and styles were customized in the table builder.
[/Notes]
`);
    }
  });

  await pptx.writeFile({ fileName });
}

async function generateTableDeck({ previewOnly }) {
  const activeColumns = getIncludedColumns();
  if (!activeColumns.length) {
    throw new Error("Select at least one column to include.");
  }

  const options = getOptions();
  const title = options.title;
  const baseName = options.filePrefix.endsWith(".pptx")
    ? options.filePrefix.slice(0, -5)
    : options.filePrefix;

  const allChunks = chunkRows(rows, options.rowsPerSlide);
  if (!allChunks.length) {
    throw new Error("No rows available to export.");
  }

  const limitedChunks = previewOnly
    ? allChunks.slice(0, options.previewSlideCount)
    : allChunks;

  const estimatedMb = estimateDeckSizeMb(rows.length, activeColumns.length, limitedChunks.length);
  const shouldSplit = !previewOnly && options.splitExport && estimatedMb > options.maxFileSizeMb;

  let deckChunks = [limitedChunks];
  if (shouldSplit) {
    const targetDeckCount = Math.ceil(estimatedMb / options.maxFileSizeMb);
    const slidesPerDeck = Math.max(1, Math.ceil(limitedChunks.length / targetDeckCount));
    deckChunks = chunkRows(limitedChunks, slidesPerDeck);
  }

  updateProgress(5, "Preparing export...", true);

  for (let i = 0; i < deckChunks.length; i += 1) {
    const part = i + 1;
    const pct = 10 + Math.round((part / deckChunks.length) * 80);
    const fileName = deckChunks.length > 1 ? `${baseName}-part-${part}.pptx` : `${baseName}.pptx`;

    updateProgress(pct, `Generating file ${part} of ${deckChunks.length}...`, true);

    await generateDeckFile({
      fileName,
      title,
      activeColumns,
      slideRowChunks: deckChunks[i],
      options,
      deckIndex: part,
      deckCount: deckChunks.length,
    });
  }

  updateProgress(100, "Export complete.", false);
  return { fileCount: deckChunks.length, previewOnly };
}

fileInput.addEventListener("change", handleFileUpload);

selectAllButton.addEventListener("click", (event) => {
  event.preventDefault();
  if (selectAllButton.getAttribute("aria-disabled") === "true") {
    return;
  }
  columns.forEach((column) => {
    column.included = true;
  });
  renderColumnConfig();
  renderPreview();
});

deselectAllButton.addEventListener("click", (event) => {
  event.preventDefault();
  if (deselectAllButton.getAttribute("aria-disabled") === "true") {
    return;
  }
  columns.forEach((column) => {
    column.included = false;
  });
  renderColumnConfig();
  renderPreview();
});

optionIds.forEach((id) => {
  const el = document.getElementById(id);
  el?.addEventListener("input", renderPreview);
  el?.addEventListener("change", renderPreview);
});

previewButton.addEventListener("click", async () => {
  if (!rows.length || !columns.length) {
    setStatus("Upload a valid file first.", true);
    return;
  }

  const previewInput = document.getElementById("preview-slide-count");
  const previewSlideCount = toPositiveWholeNumber(previewInput.value, Number.NaN);
  if (!Number.isFinite(previewSlideCount)) {
    setStatus("Preview slide count must be a whole positive number.", true);
    previewInput.focus();
    return;
  }

  try {
    setStatus("Generating preview export...");
    await generateTableDeck({ previewOnly: true });
    setStatus("Preview download started.");
  } catch (error) {
    console.error(error);
    updateProgress(0, "Idle", false);
    setStatus(error.message || "Failed to generate preview export.", true);
  }
});

generateButton.addEventListener("click", async () => {
  if (!rows.length || !columns.length) {
    setStatus("Upload a valid file first.", true);
    return;
  }

  try {
    setStatus("Generating full export...");
    const result = await generateTableDeck({ previewOnly: false });
    setStatus(
      result.fileCount > 1
        ? `Download started for ${result.fileCount} files (split export).`
        : "Download started for full export."
    );
  } catch (error) {
    console.error(error);
    updateProgress(0, "Idle", false);
    setStatus(error.message || "Failed to generate full export.", true);
  }
});

setColumnBulkLinksEnabled(false);
updateProgress(0, "Idle", false);
