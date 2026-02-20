const fileInput = document.getElementById("data-file");
const columnConfigEl = document.getElementById("column-config");
const generateButton = document.getElementById("generate-table");
const statusEl = document.getElementById("table-status");
const previewHeadEl = document.getElementById("preview-head");
const previewBodyEl = document.getElementById("preview-body");

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

function renderPreview() {
  previewHeadEl.innerHTML = "";
  previewBodyEl.innerHTML = "";

  if (!columns.length || !rows.length) {
    return;
  }

  const orderedCols = columns.map((col) => col.name);

  const headRow = document.createElement("tr");
  orderedCols.forEach((name) => {
    const th = document.createElement("th");
    th.textContent = name;
    headRow.appendChild(th);
  });
  previewHeadEl.appendChild(headRow);

  rows.slice(0, 8).forEach((row) => {
    const tr = document.createElement("tr");
    orderedCols.forEach((name) => {
      const td = document.createElement("td");
      td.textContent = row[name] ?? "";
      tr.appendChild(td);
    });
    previewBodyEl.appendChild(tr);
  });
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

    item.append(name, widthInput, alignSelect);
    columnConfigEl.appendChild(item);
  });
}

function initializeColumns(names) {
  columns = names.map((name) => ({
    name,
    width: 13.0 / Math.max(names.length, 1),
    align: "left",
  }));
  renderColumnConfig();
  renderPreview();
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
    generateButton.disabled = false;
    setStatus(`Loaded ${rows.length} rows and ${columns.length} columns from ${file.name}.`);
  } catch (error) {
    console.error(error);
    rows = [];
    columns = [];
    renderColumnConfig();
    renderPreview();
    generateButton.disabled = true;
    setStatus(error.message || "Failed to parse file.", true);
  }
}

function buildTableRows() {
  const headerFill = sanitizeHex(document.getElementById("header-fill").value);
  const headerText = sanitizeHex(document.getElementById("header-text").value);

  const header = columns.map((column) => ({
    text: column.name,
    options: {
      bold: true,
      color: headerText,
      align: column.align,
      fill: headerFill,
    },
  }));

  const body = rows.map((row) =>
    columns.map((column) => ({
      text: String(row[column.name] ?? ""),
      options: {
        align: column.align,
      },
    }))
  );

  return [header, ...body];
}

async function generateTableDeck() {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "PptxGenJS Table Builder";
  pptx.subject = "Uploaded data to PowerPoint table";

  const title = document.getElementById("slide-title").value.trim() || "Uploaded Data Table";
  const fileName = document.getElementById("file-name").value.trim() || "pptxgenjs-uploaded-table.pptx";

  const fontSize = toNumber(document.getElementById("font-size").value, 12);
  const rowH = toNumber(document.getElementById("row-height").value, 0.45);
  const margin = toNumber(document.getElementById("cell-margin").value, 0.04);
  const valign = document.getElementById("valign").value;
  const autoFit = document.getElementById("autofit").value === "true";
  const includeNotes = document.getElementById("include-notes").value === "true";

  const bodyFill = sanitizeHex(document.getElementById("body-fill").value);
  const borderColor = sanitizeHex(document.getElementById("border-color").value);

  const slide = pptx.addSlide();
  slide.addText(title, {
    x: 0.5,
    y: 0.2,
    w: 12.2,
    h: 0.45,
    bold: true,
    fontSize: 20,
  });

  const tableRows = buildTableRows();
  const columnWidths = columns.map((column) => toNumber(column.width, 1.2));

  slide.addTable(tableRows, {
    x: 0.5,
    y: 0.85,
    w: 12.3,
    colW: columnWidths,
    rowH,
    fontSize,
    color: "0F172A",
    fill: bodyFill,
    border: { pt: 1, color: borderColor },
    valign,
    autoFit,
    margin,
  });

  if (includeNotes) {
    slide.addNotes(`
[Notes]
- Generated from uploaded data.
- Column order and widths were customized in the table builder.
[/Notes]
`);
  }

  await pptx.writeFile({ fileName });
}

fileInput.addEventListener("change", handleFileUpload);

generateButton.addEventListener("click", async () => {
  if (!rows.length || !columns.length) {
    setStatus("Upload a valid file first.", true);
    return;
  }

  try {
    setStatus("Generating presentation...");
    await generateTableDeck();
    setStatus("Download started.");
  } catch (error) {
    console.error(error);
    setStatus("Failed to generate presentation.", true);
  }
});
