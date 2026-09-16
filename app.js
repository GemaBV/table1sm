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

function renderSheet(workbook, sheetName) {
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
    sheetContainer.replaceChildren();
    const message = document.createElement("p");
    message.className = "status";
    message.textContent = "Esta hoja está vacía.";
    sheetContainer.append(message);
    return;
  }
  table.removeAttribute("contenteditable");
  table.setAttribute("aria-label", sheetName);
  sheetContainer.replaceChildren(table);
}

function createTabs(workbook) {
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
      renderSheet(workbook, sheetName);
    });
    tabsContainer.append(button);
  });
  tabsContainer.hidden = workbook.SheetNames.length < 2;
}

async function loadWorkbook() {
  try {
    const manifestResponse = await fetch("workbook.json", { cache: "no-store" });
    if (!manifestResponse.ok) throw new Error("No se encontró la información del libro.");
    const manifest = await manifestResponse.json();

    filenameElement.textContent = manifest.filename;
    document.title = manifest.filename;
    downloadLink.download = manifest.filename;

    const workbookResponse = await fetch("workbook.xlsx", { cache: "no-store" });
    if (!workbookResponse.ok) throw new Error("No se pudo descargar el libro.");
    const data = await workbookResponse.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: true });

    if (!workbook.SheetNames.length) throw new Error("El libro no contiene hojas.");
    createTabs(workbook);
    renderSheet(workbook, workbook.SheetNames[0]);
  } catch (error) {
    console.error(error);
    showError("No se pudo mostrar el archivo Excel. Inténtalo de nuevo más tarde.");
  }
}

loadWorkbook();
