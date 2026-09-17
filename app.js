const sheetContainer = document.querySelector("#sheet");
const tabsContainer = document.querySelector("#sheet-tabs");
const filenameElement = document.querySelector("#filename");
const downloadLink = document.querySelector("#download");

function showError(message) {
  sheetContainer.replaceChildren();
  const paragraph = document.createElement("p");
  paragraph.className = "status error";
  paragraph.textContent = message;
  sheetContainer.append(paragraph);
}

function repositoryFromPagesUrl() {
  if (!location.hostname.endsWith(".github.io")) return null;
  const owner = location.hostname.slice(0, -".github.io".length);
  const repository = location.pathname.split("/").filter(Boolean)[0];
  return owner && repository ? { owner, repository } : null;
}

async function discoverWorkbook() {
  const repository = repositoryFromPagesUrl();
  if (!repository) {
    throw new Error("Falta workbook.json y no se pudo identificar el repositorio.");
  }

  const apiUrl = `https://api.github.com/repos/${encodeURIComponent(repository.owner)}/${encodeURIComponent(repository.repository)}/contents/`;
  const response = await fetch(apiUrl, {
    headers: { Accept: "application/vnd.github+json" },
  });
  if (!response.ok) throw new Error("No se pudo consultar el contenido del repositorio.");

  const entries = await response.json();
  const workbooks = entries.filter((entry) =>
    entry.type === "file" && /\.(xlsx|xls|xlsm|xlsb)$/i.test(entry.name)
  );
  if (workbooks.length !== 1) {
    throw new Error(`Se esperaba un archivo Excel y se encontraron ${workbooks.length}.`);
  }

  return {
    filename: workbooks[0].name,
    url: new URL(encodeURIComponent(workbooks[0].name), location.href).href,
  };
}

async function workbookSource() {
  const response = await fetch("workbook.json", { cache: "no-store" });
  if (response.ok) {
    const manifest = await response.json();
    return { filename: manifest.filename, url: "workbook.xlsx" };
  }
  return discoverWorkbook();
}

const THEME_COLORS = [
  "#ffffff", "#000000", "#e7e6e6", "#44546a", "#5b9bd5", "#ed7d31",
  "#a5a5a5", "#ffc000", "#4472c4", "#70ad47", "#0563c1", "#954f72",
];

const INDEXED_COLORS = {
  0: "#000000", 1: "#ffffff", 2: "#ff0000", 3: "#00ff00",
  4: "#0000ff", 5: "#ffff00", 6: "#ff00ff", 7: "#00ffff",
  8: "#000000", 9: "#ffffff", 10: "#ff0000", 11: "#00ff00",
  12: "#0000ff", 13: "#ffff00", 14: "#ff00ff", 15: "#00ffff",
  22: "#c0c0c0", 23: "#808080",
};

const BORDER_STYLES = {
  hair: "1px solid", thin: "1px solid", medium: "2px solid",
  thick: "3px solid", double: "3px double", dotted: "1px dotted",
  dashed: "1px dashed", dashDot: "1px dashed", dashDotDot: "1px dashed",
  mediumDashed: "2px dashed", mediumDashDot: "2px dashed",
  mediumDashDotDot: "2px dashed", slantDashDot: "2px dashed",
};

function colorToCss(color) {
  if (!color) return null;
  if (color.argb) {
    const argb = color.argb.padStart(8, "F");
    const alpha = Number.parseInt(argb.slice(0, 2), 16) / 255;
    const rgb = argb.slice(2);
    return alpha < 1
      ? `rgb(${Number.parseInt(rgb.slice(0, 2), 16)} ${Number.parseInt(rgb.slice(2, 4), 16)} ${Number.parseInt(rgb.slice(4), 16)} / ${alpha})`
      : `#${rgb}`;
  }
  if (color.indexed !== undefined) return INDEXED_COLORS[color.indexed] || null;
  if (color.theme !== undefined) return THEME_COLORS[color.theme] || null;
  return null;
}

function applyFontStyle(element, font = {}) {
  if (font.name) element.style.fontFamily = `"${font.name}", Arial, sans-serif`;
  if (font.size) element.style.fontSize = `${font.size}pt`;
  if (font.bold) element.style.fontWeight = "700";
  if (font.italic) element.style.fontStyle = "italic";
  const color = colorToCss(font.color);
  if (color) element.style.color = color;
  const decorations = [];
  if (font.underline) decorations.push("underline");
  if (font.strike) decorations.push("line-through");
  if (decorations.length) element.style.textDecoration = decorations.join(" ");
  if (font.vertAlign === "superscript") element.style.verticalAlign = "super";
  if (font.vertAlign === "subscript") element.style.verticalAlign = "sub";
}

function applyBorderStyle(element, border = {}) {
  for (const side of ["top", "right", "bottom", "left"]) {
    const edge = border[side];
    const line = edge?.style ? BORDER_STYLES[edge.style] : null;
    element.style[`border${side[0].toUpperCase()}${side.slice(1)}`] = line
      ? `${line} ${colorToCss(edge.color) || "#000"}`
      : "none";
  }
}

function applyCellStyle(element, cell) {
  applyFontStyle(element, cell.font);
  applyBorderStyle(element, cell.border);

  if (cell.fill?.type === "pattern" && cell.fill.pattern === "solid") {
    const color = colorToCss(cell.fill.fgColor);
    if (color) element.style.backgroundColor = color;
  }

  const alignment = cell.alignment || {};
  const horizontal = { centerContinuous: "center", distributed: "justify" };
  if (alignment.horizontal) {
    element.style.textAlign = horizontal[alignment.horizontal] || alignment.horizontal;
  }
  if (alignment.vertical) element.style.verticalAlign = alignment.vertical;
  if (alignment.wrapText === false) element.style.whiteSpace = "nowrap";
  if (alignment.wrapText) element.style.whiteSpace = "pre-wrap";
  if (alignment.indent) element.style.paddingLeft = `${7 + alignment.indent * 12}px`;
  if (alignment.textRotation === "vertical") {
    element.style.writingMode = "vertical-rl";
  }
}

function columnNumber(reference) {
  return [...reference.toUpperCase()].reduce((number, letter) =>
    number * 26 + letter.charCodeAt(0) - 64, 0);
}

function parseRange(reference) {
  const [start, end] = reference.split(":");
  const decode = (address) => {
    const match = address.match(/^([A-Z]+)(\d+)$/i);
    return { column: columnNumber(match[1]), row: Number(match[2]) };
  };
  return { start: decode(start), end: decode(end || start) };
}

function mergeMap(worksheet) {
  const starts = new Map();
  const covered = new Set();
  for (const reference of worksheet.model.merges || []) {
    const range = parseRange(reference);
    starts.set(`${range.start.row}:${range.start.column}`, range);
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let column = range.start.column; column <= range.end.column; column += 1) {
        if (row !== range.start.row || column !== range.start.column) {
          covered.add(`${row}:${column}`);
        }
      }
    }
  }
  return { starts, covered };
}

function safeHyperlink(url) {
  try {
    const parsed = new URL(url, location.href);
    return ["http:", "https:", "mailto:"].includes(parsed.protocol) ? parsed.href : null;
  } catch {
    return null;
  }
}

function appendCellContent(element, cell, formattedCell) {
  const richText = cell.value?.richText;
  if (richText) {
    for (const part of richText) {
      const span = document.createElement("span");
      span.textContent = part.text;
      applyFontStyle(span, part.font);
      element.append(span);
    }
    return;
  }

  const text = formattedCell?.w ?? cell.text ?? "";
  const hyperlink = typeof cell.value === "object" && cell.value?.hyperlink
    ? safeHyperlink(cell.value.hyperlink)
    : null;
  if (hyperlink) {
    const anchor = document.createElement("a");
    anchor.href = hyperlink;
    anchor.textContent = text;
    anchor.target = "_blank";
    anchor.rel = "noopener noreferrer";
    element.append(anchor);
  } else {
    element.textContent = text;
  }
}

function renderSheet(workbook, formattedWorkbook, sheetName) {
  const worksheet = workbook.getWorksheet(sheetName);
  const formattedWorksheet = formattedWorkbook.Sheets[sheetName];
  if (!worksheet || !worksheet.rowCount || !worksheet.columnCount) {
    sheetContainer.replaceChildren();
    const message = document.createElement("p");
    message.className = "status";
    message.textContent = "Esta hoja está vacía.";
    sheetContainer.append(message);
    return;
  }
  const table = document.createElement("table");
  table.setAttribute("aria-label", sheetName);
  const columns = document.createElement("colgroup");
  for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
    const column = document.createElement("col");
    const definition = worksheet.getColumn(columnNumber);
    if (definition.width) column.style.width = `${Math.max(24, definition.width * 7 + 5)}px`;
    if (definition.hidden) column.style.visibility = "collapse";
    columns.append(column);
  }
  table.append(columns);

  const body = document.createElement("tbody");
  const merges = mergeMap(worksheet);
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const excelRow = worksheet.getRow(rowNumber);
    if (excelRow.hidden) continue;
    const row = document.createElement("tr");
    if (excelRow.height) row.style.height = `${excelRow.height * 96 / 72}px`;

    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      const key = `${rowNumber}:${columnNumber}`;
      if (merges.covered.has(key)) continue;
      const cell = excelRow.getCell(columnNumber);
      const element = document.createElement("td");
      const mergedRange = merges.starts.get(key);
      if (mergedRange) {
        element.rowSpan = mergedRange.end.row - mergedRange.start.row + 1;
        element.colSpan = mergedRange.end.column - mergedRange.start.column + 1;
      }
      applyCellStyle(element, cell);
      const address = cell.address;
      appendCellContent(element, cell, formattedWorksheet?.[address]);
      if (cell.note) element.title = typeof cell.note === "string" ? cell.note : "Esta celda contiene una nota";
      row.append(element);
    }
    body.append(row);
  }
  table.append(body);
  sheetContainer.replaceChildren(table);
}

function renderBasicSheet(workbook, sheetName) {
  const worksheet = workbook.Sheets[sheetName];
  const markup = XLSX.utils.sheet_to_html(worksheet, {
    editable: false,
    header: "",
    footer: "",
  });
  const template = document.createElement("template");
  template.innerHTML = markup.trim();
  const table = template.content.querySelector("table");
  if (!table) {
    showError("Esta hoja está vacía.");
    return;
  }
  table.removeAttribute("contenteditable");
  table.setAttribute("aria-label", sheetName);
  sheetContainer.replaceChildren(table);
}

function createTabs(workbook, formattedWorkbook) {
  tabsContainer.replaceChildren();
  workbook.worksheets.forEach((worksheet, index) => {
    const sheetName = worksheet.name;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "sheet-tab";
    button.textContent = sheetName;
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", String(index === 0));
    button.addEventListener("click", () => {
      tabsContainer.querySelectorAll(".sheet-tab").forEach((tab) => {
        tab.setAttribute("aria-selected", String(tab === button));
      });
      renderSheet(workbook, formattedWorkbook, sheetName);
    });
    tabsContainer.append(button);
  });
  tabsContainer.hidden = workbook.worksheets.length < 2;
}

function createBasicTabs(workbook) {
  tabsContainer.replaceChildren();
  workbook.SheetNames.forEach((sheetName, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "sheet-tab";
    button.textContent = sheetName;
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", String(index === 0));
    button.addEventListener("click", () => {
      tabsContainer.querySelectorAll(".sheet-tab").forEach((tab) => {
        tab.setAttribute("aria-selected", String(tab === button));
      });
      renderBasicSheet(workbook, sheetName);
    });
    tabsContainer.append(button);
  });
  tabsContainer.hidden = workbook.SheetNames.length < 2;
}

async function loadWorkbook() {
  try {
    const source = await workbookSource();
    filenameElement.textContent = source.filename;
    document.title = source.filename;
    downloadLink.href = source.url;
    downloadLink.download = source.filename;

    const workbookResponse = await fetch(source.url, { cache: "no-store" });
    if (!workbookResponse.ok) throw new Error("No se pudo descargar el libro.");
    const data = await workbookResponse.arrayBuffer();
    const formattedWorkbook = XLSX.read(data, { type: "array", cellDates: true });
    if (!formattedWorkbook.SheetNames.length) throw new Error("El libro no contiene hojas.");

    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(data);
      if (!workbook.worksheets.length) throw new Error("El libro no contiene hojas.");
      createTabs(workbook, formattedWorkbook);
      renderSheet(workbook, formattedWorkbook, workbook.worksheets[0].name);
    } catch (styleError) {
      console.warn("No se pudo cargar el formato; se usará la vista básica.", styleError);
      createBasicTabs(formattedWorkbook);
      renderBasicSheet(formattedWorkbook, formattedWorkbook.SheetNames[0]);
    }
  } catch (error) {
    console.error(error);
    showError(`No se pudo mostrar el archivo Excel. ${error.message}`);
  }
}

loadWorkbook();
