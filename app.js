const STORAGE_DB = "inventario-prototipo-db";
const STORAGE_STORE = "workbooks";
const STORAGE_KEY = "active-workbook";

const state = {
  workbook: null,
  workbookName: "",
  activeSheetName: "",
  selectedCellAddress: "",
  selectedRowNumber: null,
  pageSize: 50,
  page: 1,
  searchTerm: "",
  dirty: false,
  saving: false,
  saveTimeoutId: null
};

const dom = {
  excelFile: document.getElementById("excelFile"),
  fileSummary: document.getElementById("fileSummary"),
  saveLocalBtn: document.getElementById("saveLocalBtn"),
  exportBtn: document.getElementById("exportBtn"),
  deleteWorkbookBtn: document.getElementById("deleteWorkbookBtn"),
  pageSizeSelect: document.getElementById("pageSizeSelect"),
  sheetSearch: document.getElementById("sheetSearch"),
  selectedCellLabel: document.getElementById("selectedCellLabel"),
  selectedCellMeta: document.getElementById("selectedCellMeta"),
  formulaBar: document.getElementById("formulaBar"),
  applyCellBtn: document.getElementById("applyCellBtn"),
  clearCellBtn: document.getElementById("clearCellBtn"),
  addRowBtn: document.getElementById("addRowBtn"),
  deleteRowBtn: document.getElementById("deleteRowBtn"),
  workbookTitle: document.getElementById("workbookTitle"),
  sheetCount: document.getElementById("sheetCount"),
  visibleRowsCount: document.getElementById("visibleRowsCount"),
  changeState: document.getElementById("changeState"),
  sheetTabs: document.getElementById("sheetTabs"),
  prevPageBtn: document.getElementById("prevPageBtn"),
  nextPageBtn: document.getElementById("nextPageBtn"),
  pageIndicator: document.getElementById("pageIndicator"),
  sheetMeta: document.getElementById("sheetMeta"),
  tableWrapper: document.getElementById("tableWrapper"),
  confirmModal: document.getElementById("confirmModal"),
  modalTitle: document.getElementById("modalTitle"),
  modalMessage: document.getElementById("modalMessage"),
  modalCancelBtn: document.getElementById("modalCancelBtn"),
  modalConfirmBtn: document.getElementById("modalConfirmBtn")
};

let confirmResolver = null;

initialize();

async function initialize() {
  bindEvents();
  await restoreSavedWorkbook();
  render();
}

function bindEvents() {
  dom.excelFile.addEventListener("change", handleFileUpload);
  dom.saveLocalBtn.addEventListener("click", saveWorkbookToBrowser);
  dom.exportBtn.addEventListener("click", exportWorkbook);
  dom.deleteWorkbookBtn.addEventListener("click", handleDeleteWorkbook);
  dom.pageSizeSelect.addEventListener("change", handlePageSizeChange);
  dom.sheetSearch.addEventListener("input", handleSearchInput);
  dom.applyCellBtn.addEventListener("click", applyCellChange);
  dom.clearCellBtn.addEventListener("click", clearSelectedCell);
  dom.addRowBtn.addEventListener("click", addRowBelowSelection);
  dom.deleteRowBtn.addEventListener("click", deleteSelectedRow);
  dom.prevPageBtn.addEventListener("click", () => changePage(-1));
  dom.nextPageBtn.addEventListener("click", () => changePage(1));
  dom.formulaBar.addEventListener("keydown", handleFormulaBarKeyboard);
  dom.modalCancelBtn.addEventListener("click", () => closeConfirmModal(false));
  dom.modalConfirmBtn.addEventListener("click", () => closeConfirmModal(true));
}

async function handleFileUpload(event) {
  const [file] = event.target.files || [];
  if (!file) {
    return;
  }

  const buffer = await file.arrayBuffer();
  loadWorkbookFromArrayBuffer(buffer, file.name);
  await saveWorkbookRecord(file.name, buffer);
  state.dirty = false;
  render();
  event.target.value = "";
}

function loadWorkbookFromArrayBuffer(buffer, workbookName) {
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    cellText: true
  });

  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.CalcPr = {
    calcMode: "auto",
    fullCalcOnLoad: "1",
    forceFullCalc: "1"
  };

  state.workbook = workbook;
  state.workbookName = workbookName;
  state.activeSheetName = workbook.SheetNames[0] || "";
  state.selectedCellAddress = "";
  state.selectedRowNumber = null;
  state.page = 1;
  dom.sheetSearch.value = "";
  state.searchTerm = "";
}

async function restoreSavedWorkbook() {
  const savedRecord = await readWorkbookRecord();
  if (!savedRecord) {
    return;
  }

  loadWorkbookFromArrayBuffer(savedRecord.data, savedRecord.name);
  state.dirty = false;
}

function render() {
  renderFileSummary();
  renderTopbar();
  renderSheetTabs();
  renderGrid();
  renderSelectionCard();
  renderButtons();
}

function renderFileSummary() {
  const dotClass = state.workbook ? (state.dirty ? "status-dot dirty" : "status-dot ready") : "status-dot";
  const statusTitle = state.workbook ? state.workbookName : "No hay archivo cargado";
  const statusText = state.workbook
    ? `${state.workbook.SheetNames.length} hoja(s) disponibles. ${state.dirty ? "Tienes cambios pendientes." : "Guardado local listo."}`
    : "Cuando subas un Excel quedar\u00e1 guardado localmente en este navegador.";

  dom.fileSummary.innerHTML = `
    <span class="${dotClass}"></span>
    <div>
      <strong>${escapeHtml(statusTitle)}</strong>
      <p>${escapeHtml(statusText)}</p>
    </div>
  `;
}

function renderTopbar() {
  dom.workbookTitle.textContent = state.workbookName || "Sin archivo";
  dom.sheetCount.textContent = state.workbook ? String(state.workbook.SheetNames.length) : "0";
  dom.visibleRowsCount.textContent = state.workbook ? String(getCurrentPageRows().length) : "0";
  dom.changeState.textContent = state.saving ? "Guardando..." : state.dirty ? "Cambios pendientes" : "Sin cambios";
}

function renderSheetTabs() {
  if (!state.workbook) {
    dom.sheetTabs.innerHTML = "";
    return;
  }

  dom.sheetTabs.innerHTML = state.workbook.SheetNames.map((sheetName) => {
    const activeClass = sheetName === state.activeSheetName ? "sheet-tab active" : "sheet-tab";
    return `<button type="button" class="${activeClass}" data-sheet-name="${escapeAttribute(sheetName)}">${escapeHtml(sheetName)}</button>`;
  }).join("");

  dom.sheetTabs.querySelectorAll("[data-sheet-name]").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeSheetName = button.dataset.sheetName;
      state.selectedCellAddress = "";
      state.selectedRowNumber = null;
      state.page = 1;
      render();
    });
  });
}

function renderGrid() {
  if (!state.workbook || !state.activeSheetName) {
    dom.sheetMeta.textContent = "Carga un archivo Excel para comenzar.";
    dom.pageIndicator.textContent = "P\u00e1gina 0 de 0";
    dom.tableWrapper.innerHTML = `
      <div class="empty-state">
        <h3>Sube tu archivo</h3>
        
      </div>
    `;
    return;
  }

  const sheet = state.workbook.Sheets[state.activeSheetName];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const totalRows = range.e.r - range.s.r + 1;
  const totalPages = Math.max(1, Math.ceil(totalRows / state.pageSize));
  if (state.page > totalPages) {
    state.page = totalPages;
  }

  const sheetRows = getCurrentPageRows();
  const visibleCols = getVisibleColumns(sheet);
  const filteredRows = filterRows(sheetRows, visibleCols);

  dom.pageIndicator.textContent = `P\u00e1gina ${state.page} de ${totalPages}`;
  dom.sheetMeta.textContent = `Hoja activa: ${state.activeSheetName} | Filas totales: ${totalRows} | Columnas visibles: ${visibleCols.length}`;

  if (!filteredRows.length) {
    dom.tableWrapper.innerHTML = `
      <div class="empty-state">
        <h3>Sin resultados</h3>
        <p>No encontramos coincidencias en la hoja activa con el texto buscado.</p>
      </div>
    `;
    return;
  }

  const headerCells = visibleCols.map((columnIndex) => `<th>${columnIndexToName(columnIndex)}</th>`).join("");
  const bodyRows = filteredRows.map((rowNumber) => {
    const rowCells = visibleCols.map((columnIndex) => {
      const address = XLSX.utils.encode_cell({ r: rowNumber - 1, c: columnIndex });
      const cell = sheet[address];
      const displayValue = getCellDisplayValue(cell);
      const buttonClass = [
        "cell-button",
        cell && cell.f ? "formula-cell" : "",
        address === state.selectedCellAddress ? "selected" : ""
      ].filter(Boolean).join(" ");

      return `
        <td>
          <button
            type="button"
            class="${buttonClass}"
            data-address="${escapeAttribute(address)}"
            data-row-number="${rowNumber}"
            title="${escapeAttribute(displayValue || "Celda vac\u00eda")}"
          >
            <span class="cell-text">${escapeHtml(displayValue || " ")}</span>
          </button>
        </td>
      `;
    }).join("");

    return `
      <tr>
        <th class="row-header">${rowNumber}</th>
        ${rowCells}
      </tr>
    `;
  }).join("");

  dom.tableWrapper.innerHTML = `
    <table class="grid-table">
      <thead>
        <tr>
          <th class="corner-cell">Fila</th>
          ${headerCells}
        </tr>
      </thead>
      <tbody>
        ${bodyRows}
      </tbody>
    </table>
  `;

  dom.tableWrapper.querySelectorAll("[data-address]").forEach((button) => {
    button.addEventListener("click", () => selectCell(button.dataset.address, Number(button.dataset.rowNumber)));
  });
}

function renderSelectionCard() {
  if (!state.workbook || !state.activeSheetName || !state.selectedCellAddress) {
    dom.selectedCellLabel.textContent = "Ninguna";
    dom.selectedCellMeta.textContent = "Selecciona una celda para editarla.";
    dom.formulaBar.value = "";
    return;
  }

  const cell = getActiveSheet()[state.selectedCellAddress];
  const rawValue = getCellRawValue(cell);
  const cellType = cell?.f ? "F\u00f3rmula" : cell?.t === "n" ? "N\u00famero" : cell?.t === "b" ? "Booleano" : "Texto";

  dom.selectedCellLabel.textContent = `${state.activeSheetName} - ${state.selectedCellAddress}`;
  dom.selectedCellMeta.textContent = `${cellType} | Fila ${state.selectedRowNumber ?? "-"}`;
  dom.formulaBar.value = rawValue;
}

function renderButtons() {
  const hasWorkbook = Boolean(state.workbook);
  const hasSelection = hasWorkbook && Boolean(state.selectedCellAddress);
  const totalPages = hasWorkbook ? Math.max(1, Math.ceil(getSheetRowCount(getActiveSheet()) / state.pageSize)) : 0;

  dom.saveLocalBtn.disabled = !hasWorkbook;
  dom.exportBtn.disabled = !hasWorkbook;
  dom.deleteWorkbookBtn.disabled = !hasWorkbook;
  dom.applyCellBtn.disabled = !hasSelection;
  dom.clearCellBtn.disabled = !hasSelection;
  dom.addRowBtn.disabled = !hasWorkbook;
  dom.deleteRowBtn.disabled = !hasWorkbook || !state.selectedRowNumber;
  dom.prevPageBtn.disabled = !hasWorkbook || state.page <= 1;
  dom.nextPageBtn.disabled = !hasWorkbook || state.page >= totalPages;
}

function selectCell(address, rowNumber) {
  state.selectedCellAddress = address;
  state.selectedRowNumber = rowNumber;
  render();
}

function applyCellChange() {
  if (!state.workbook || !state.selectedCellAddress) {
    return;
  }

  const sheet = getActiveSheet();
  const existingCell = sheet[state.selectedCellAddress] || {};
  const nextValue = dom.formulaBar.value.trim();

  if (!nextValue) {
    delete sheet[state.selectedCellAddress];
  } else if (nextValue.startsWith("=")) {
    sheet[state.selectedCellAddress] = {
      ...existingCell,
      f: nextValue.slice(1),
      t: existingCell.t || "n"
    };
    delete sheet[state.selectedCellAddress].v;
    delete sheet[state.selectedCellAddress].w;
  } else {
    sheet[state.selectedCellAddress] = createTypedCell(nextValue, existingCell);
  }

  ensureSheetRefIncludesAddress(sheet, state.selectedCellAddress);
  markDirty();
  render();
}

function clearSelectedCell() {
  if (!state.workbook || !state.selectedCellAddress) {
    return;
  }

  delete getActiveSheet()[state.selectedCellAddress];
  markDirty();
  render();
}

function addRowBelowSelection() {
  if (!state.workbook) {
    return;
  }

  const sheet = getActiveSheet();
  const sheetRange = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const insertRowNumber = state.selectedRowNumber ? state.selectedRowNumber + 1 : sheetRange.e.r + 2;

  shiftRows(sheet, insertRowNumber - 1, 1);
  state.selectedRowNumber = insertRowNumber;
  state.selectedCellAddress = "";
  markDirty();
  render();
}

async function deleteSelectedRow() {
  if (!state.workbook || !state.selectedRowNumber) {
    return;
  }

  const confirmed = await openConfirmModal(
    "Eliminar fila",
    `Se eliminar\u00e1 la fila ${state.selectedRowNumber} de la hoja ${state.activeSheetName}.`
  );

  if (!confirmed) {
    return;
  }

  const sheet = getActiveSheet();
  shiftRows(sheet, state.selectedRowNumber - 1, -1);
  state.selectedCellAddress = "";
  state.selectedRowNumber = null;
  markDirty();
  render();
}

function changePage(direction) {
  if (!state.workbook) {
    return;
  }

  const totalPages = Math.max(1, Math.ceil(getSheetRowCount(getActiveSheet()) / state.pageSize));
  state.page = Math.min(totalPages, Math.max(1, state.page + direction));
  state.selectedCellAddress = "";
  state.selectedRowNumber = null;
  render();
}

function handlePageSizeChange(event) {
  state.pageSize = Number(event.target.value) || 50;
  state.page = 1;
  render();
}

function handleSearchInput(event) {
  state.searchTerm = event.target.value.trim().toLowerCase();
  render();
}

function handleFormulaBarKeyboard(event) {
  if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
    event.preventDefault();
    applyCellChange();
  }
}

async function saveWorkbookToBrowser() {
  if (!state.workbook) {
    return;
  }

  state.saving = true;
  renderTopbar();

  const data = workbookToArrayBuffer();
  await saveWorkbookRecord(state.workbookName || "inventario.xlsx", data);
  state.saving = false;
  state.dirty = false;
  render();
}

async function handleDeleteWorkbook() {
  if (!state.workbook) {
    return;
  }

  const confirmed = await openConfirmModal(
    "Eliminar archivo guardado",
    "Se borrar\u00e1 del navegador el archivo actualmente cargado."
  );

  if (!confirmed) {
    return;
  }

  await deleteWorkbookRecord();
  state.workbook = null;
  state.workbookName = "";
  state.activeSheetName = "";
  state.selectedCellAddress = "";
  state.selectedRowNumber = null;
  state.page = 1;
  state.searchTerm = "";
  dom.sheetSearch.value = "";
  dom.formulaBar.value = "";
  state.dirty = false;
  render();
}

function exportWorkbook() {
  if (!state.workbook) {
    return;
  }

  const data = workbookToArrayBuffer();
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = buildExportName(state.workbookName);
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}

function workbookToArrayBuffer() {
  state.workbook.Workbook = state.workbook.Workbook || {};
  state.workbook.Workbook.CalcPr = {
    calcMode: "auto",
    fullCalcOnLoad: "1",
    forceFullCalc: "1"
  };

  return XLSX.write(state.workbook, {
    bookType: "xlsx",
    type: "array",
    cellStyles: true,
    compression: true
  });
}

function buildExportName(fileName) {
  if (!fileName) {
    return "inventario-editado.xlsx";
  }

  const dotIndex = fileName.lastIndexOf(".");
  if (dotIndex === -1) {
    return `${fileName}-editado.xlsx`;
  }

  return `${fileName.slice(0, dotIndex)}-editado.xlsx`;
}

function getActiveSheet() {
  return state.workbook.Sheets[state.activeSheetName];
}

function getSheetRowCount(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  return range.e.r - range.s.r + 1;
}

function getCurrentPageRows() {
  if (!state.workbook || !state.activeSheetName) {
    return [];
  }

  const sheet = getActiveSheet();
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const start = range.s.r + (state.page - 1) * state.pageSize;
  const end = Math.min(range.e.r, start + state.pageSize - 1);
  const rows = [];

  for (let rowIndex = start; rowIndex <= end; rowIndex += 1) {
    rows.push(rowIndex + 1);
  }

  return rows;
}

function getVisibleColumns(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const visibleColumns = [];

  for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
    visibleColumns.push(columnIndex);
  }

  return visibleColumns;
}

function filterRows(rowNumbers, columns) {
  if (!state.searchTerm) {
    return rowNumbers;
  }

  const sheet = getActiveSheet();
  return rowNumbers.filter((rowNumber) => {
    return columns.some((columnIndex) => {
      const address = XLSX.utils.encode_cell({ r: rowNumber - 1, c: columnIndex });
      const displayValue = getCellDisplayValue(sheet[address]).toLowerCase();
      return displayValue.includes(state.searchTerm);
    });
  });
}

function getCellDisplayValue(cell) {
  if (!cell) {
    return "";
  }

  if (cell.w !== undefined) {
    return String(cell.w);
  }

  if (cell.v !== undefined) {
    return String(cell.v);
  }

  if (cell.f) {
    return `=${cell.f}`;
  }

  return "";
}

function getCellRawValue(cell) {
  if (!cell) {
    return "";
  }

  if (cell.f) {
    return `=${cell.f}`;
  }

  if (cell.v === undefined || cell.v === null) {
    return "";
  }

  return String(cell.v);
}

function createTypedCell(value, existingCell = {}) {
  if (/^-?\d+([.,]\d+)?$/.test(value)) {
    const numericValue = Number(value.replace(",", "."));
    const nextCell = {
      ...existingCell,
      t: "n",
      v: numericValue,
      w: String(numericValue)
    };
    delete nextCell.f;
    return nextCell;
  }

  if (value.toLowerCase() === "true" || value.toLowerCase() === "false") {
    const booleanValue = value.toLowerCase() === "true";
    const nextCell = {
      ...existingCell,
      t: "b",
      v: booleanValue,
      w: booleanValue ? "TRUE" : "FALSE"
    };
    delete nextCell.f;
    return nextCell;
  }

  const nextCell = {
    ...existingCell,
    t: "s",
    v: value,
    w: value
  };
  delete nextCell.f;
  return nextCell;
}

function ensureSheetRefIncludesAddress(sheet, address) {
  const decoded = XLSX.utils.decode_cell(address);
  const range = XLSX.utils.decode_range(sheet["!ref"] || address + ":" + address);

  range.s.r = Math.min(range.s.r, decoded.r);
  range.s.c = Math.min(range.s.c, decoded.c);
  range.e.r = Math.max(range.e.r, decoded.r);
  range.e.c = Math.max(range.e.c, decoded.c);
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function shiftRows(sheet, baseRowIndex, direction) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const entries = Object.keys(sheet)
    .filter((key) => !key.startsWith("!"))
    .map((address) => ({
      address,
      cell: sheet[address],
      decoded: XLSX.utils.decode_cell(address)
    }));

  const orderedEntries = direction > 0
    ? entries.sort((left, right) => right.decoded.r - left.decoded.r)
    : entries.sort((left, right) => left.decoded.r - right.decoded.r);

  orderedEntries.forEach(({ address, cell, decoded }) => {
    if (decoded.r < baseRowIndex) {
      return;
    }

    if (direction < 0 && decoded.r === baseRowIndex) {
      delete sheet[address];
      return;
    }

    const nextRow = decoded.r + direction;
    const nextAddress = XLSX.utils.encode_cell({ r: nextRow, c: decoded.c });
    sheet[nextAddress] = cell;
    if (nextAddress !== address) {
      delete sheet[address];
    }
  });

  if (direction > 0) {
    range.e.r += 1;
  } else {
    range.e.r = Math.max(range.s.r, range.e.r - 1);
  }

  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function markDirty() {
  state.dirty = true;
  scheduleAutoSave();
}

function scheduleAutoSave() {
  window.clearTimeout(state.saveTimeoutId);
  state.saveTimeoutId = window.setTimeout(() => {
    saveWorkbookToBrowser();
  }, 900);
}

function columnIndexToName(index) {
  let current = index + 1;
  let columnName = "";

  while (current > 0) {
    const remainder = (current - 1) % 26;
    columnName = String.fromCharCode(65 + remainder) + columnName;
    current = Math.floor((current - 1) / 26);
  }

  return columnName;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value);
}

function openConfirmModal(title, message) {
  dom.modalTitle.textContent = title;
  dom.modalMessage.textContent = message;
  dom.confirmModal.classList.remove("hidden");

  return new Promise((resolve) => {
    confirmResolver = resolve;
  });
}

function closeConfirmModal(confirmed) {
  dom.confirmModal.classList.add("hidden");
  if (confirmResolver) {
    confirmResolver(confirmed);
    confirmResolver = null;
  }
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = window.indexedDB.open(STORAGE_DB, 1);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORAGE_STORE)) {
        db.createObjectStore(STORAGE_STORE, { keyPath: "id" });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveWorkbookRecord(name, data) {
  const db = await openDatabase();

  await new Promise((resolve, reject) => {
    const tx = db.transaction(STORAGE_STORE, "readwrite");
    tx.objectStore(STORAGE_STORE).put({
      id: STORAGE_KEY,
      name,
      data,
      savedAt: new Date().toISOString()
    });
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });

  db.close();
}

async function readWorkbookRecord() {
  const db = await openDatabase();

  const result = await new Promise((resolve, reject) => {
    const tx = db.transaction(STORAGE_STORE, "readonly");
    const request = tx.objectStore(STORAGE_STORE).get(STORAGE_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });

  db.close();
  return result;
}

async function deleteWorkbookRecord() {
  const db = await openDatabase();

  await new Promise((resolve, reject) => {
    const tx = db.transaction(STORAGE_STORE, "readwrite");
    tx.objectStore(STORAGE_STORE).delete(STORAGE_KEY);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });

  db.close();
}
