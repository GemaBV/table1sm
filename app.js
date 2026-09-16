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
    const source = await workbookSource();
    filenameElement.textContent = source.filename;
    document.title = source.filename;
    downloadLink.href = source.url;
    downloadLink.download = source.filename;

    const workbookResponse = await fetch(source.url, { cache: "no-store" });
    if (!workbookResponse.ok) throw new Error("No se pudo descargar el libro.");
    const data = await workbookResponse.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: true });

    if (!workbook.SheetNames.length) throw new Error("El libro no contiene hojas.");
    createTabs(workbook);
    renderSheet(workbook, workbook.SheetNames[0]);
  } catch (error) {
    console.error(error);
    showError(`No se pudo mostrar el archivo Excel. ${error.message}`);
  }
}

loadWorkbook();
