/* UI for the data-cleaning tool. Depends on window.DataCleaner and window.XLSX. */
(function () {
  "use strict";

  const cleaner = window.DataCleaner;
  const PREVIEW_ROW_LIMIT = 200;
  const SAMPLE_CSV = [
    "顧客ID,会社名,担当者,電話番号,メール,都道府県,契約プラン",
    "C001,株式会社サクラ商事,山田 太郎,03-1234-5678,yamada@example.com,東京都,Standard",
    "C002,（株）サクラ商事,山田　太郎,０３－１２３４－５６７８,yamada2@example.com,東京都,standard",
    "C003,ミドリ物産株式会社,佐藤 花子,06-9876-5432,sato@example.jp,大阪府,Premium",
    "C003,ミドリ物産株式会社, 佐藤　花子 ,06ー9876ー5432,sato@example.jp,大阪府,Premium",
    "C004,ｱｵｿﾞﾗ工業,鈴木 一郎,045-111-2222,N/A,神奈川県,Lite",
    "C005,アオゾラ工業,鈴木  一郎,045-111-2222,suzuki@example.com,神奈川県,LITE",
    "C006,,,,,,",
    ",,,,,,",
    "C007,株式会社サクラ商事,田中 次郎,03-1234-5679,tanaka@example.com,東京都,STANDARD",
    "C008,ﾐﾄﾞﾘ物産株式会社,高橋 優,06-9876-5433,-,大阪府,premium ",
    "C009,ヒカリ電機株式会社,伊藤 愛,052-333-4444,ito@example.com,愛知県,Standard",
  ].join("\n");

  const state = {
    fileName: "sample.csv",
    isSample: true,
    workbook: null,
    table: { headers: [], rows: [] },
    view: "after",
    requiredColumns: new Set(),
    duplicateColumns: new Set(),
    variantChoices: new Map(), // `${column}\u0001${key}` -> {enabled, representative}
    variantGroups: [],
    manualEdits: {}, // "origin:column" -> typed value
    pendingRows: new Set(), // rows added by hand, kept while still blank
    numericColumns: [],
    result: null,
  };
  const KIND_LABEL = { normalize: "表記の正規化", blank: "空欄扱い", variant: "表記揺れの統一", manual: "手入力" };

  const element = (id) => document.getElementById(id);
  const downloadsReady =
    window.claude && typeof window.claude.use === "function"
      ? window.claude.use("downloads").catch(() => null)
      : Promise.resolve(null);
  let downloadsApi = null;
  downloadsReady.then((api) => {
    downloadsApi = api;
  });

  function make(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text != null) node.textContent = text;
    return node;
  }

  function readSettings() {
    return {
      normalize: {
        widthNormalize: element("opt-width").checked,
        trim: element("opt-trim").checked,
        collapseSpaces: element("opt-spaces").checked,
        unifyDashes: element("opt-dashes").checked,
      },
      placeholdersAsBlank: element("opt-placeholders").checked,
      removeEmptyRows: element("opt-empty-rows").checked,
      removeDuplicates: element("opt-duplicates").checked,
      requiredColumns: [...state.requiredColumns].sort((a, b) => a - b),
      duplicateKeyColumns: [...state.duplicateColumns].sort((a, b) => a - b),
      manualEdits: state.manualEdits,
      pendingRows: [...state.pendingRows],
    };
  }

  // ---------- Loading ----------

  function setTable(table, fileName, detail, isSample) {
    state.table = table;
    state.fileName = fileName;
    state.isSample = isSample;
    state.requiredColumns = new Set();
    state.duplicateColumns = new Set();
    state.variantChoices = new Map();
    state.manualEdits = {};
    state.pendingRows = new Set();
    element("file-name").textContent = fileName;
    element("file-detail").textContent = detail;
    const badge = element("source-badge");
    badge.textContent = isSample ? "サンプル" : "読み込み済み";
    badge.className = `badge ${isSample ? "sample" : "user"}`;
    element("load-error").hidden = true;
    renderColumnChips();
    recompute();
  }

  function loadSample() {
    state.workbook = null;
    element("sheet-picker").hidden = true;
    const table = cleaner.buildTable(cleaner.parseCSV(SAMPLE_CSV));
    setTable(table, "sample.csv", `${table.rows.length}行 × ${table.headers.length}列`, true);
  }

  function showLoadError(message) {
    const node = element("load-error");
    node.textContent = message;
    node.hidden = false;
  }

  async function loadFile(file) {
    try {
      const buffer = await file.arrayBuffer();
      const lowerName = file.name.toLowerCase();
      if (/\.(xlsx|xls)$/.test(lowerName)) {
        if (!window.XLSX) {
          showLoadError("Excel読み込み用のライブラリを読み込めませんでした。ネット接続を確認して再読み込みしてください。");
          return;
        }
        state.workbook = window.XLSX.read(buffer, { type: "array" });
        const select = element("sheet-select");
        select.replaceChildren(...state.workbook.SheetNames.map((name) => new Option(name, name)));
        element("sheet-picker").hidden = state.workbook.SheetNames.length < 2;
        state.fileName = file.name;
        loadSheet(state.workbook.SheetNames[0]);
        return;
      }
      state.workbook = null;
      element("sheet-picker").hidden = true;
      const { text, encoding } = cleaner.decodeText(buffer);
      const delimiter = lowerName.endsWith(".tsv") ? "\t" : cleaner.detectDelimiter(text);
      const table = cleaner.buildTable(cleaner.parseCSV(text, delimiter));
      if (table.headers.length === 0) {
        showLoadError("データが見つかりませんでした。空のファイルでないか確認してください。");
        return;
      }
      const delimiterName = { ",": "カンマ", "\t": "タブ", ";": "セミコロン" }[delimiter];
      setTable(table, file.name, `${table.rows.length}行 × ${table.headers.length}列 ・ ${encoding} ・ ${delimiterName}区切り`, false);
    } catch (error) {
      showLoadError(`読み込めませんでした（${error.message}）。CSVかExcel形式のファイルを選んでください。`);
    }
  }

  function loadPastedText() {
    const text = element("paste-input").value;
    if (text.trim() === "") {
      showLoadError("貼り付け欄が空です。データを貼り付けるか入力してください。");
      return;
    }
    state.workbook = null;
    element("sheet-picker").hidden = true;
    const table = cleaner.parseText(text);
    setTable(table, "貼り付けデータ", `${table.rows.length}行 × ${table.headers.length}列`, false);
    state.view = "after";
    renderView();
  }

  function loadSheet(sheetName) {
    const sheet = state.workbook.Sheets[sheetName];
    const matrix = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "", blankrows: true });
    const table = cleaner.buildTable(matrix);
    setTable(table, state.fileName, `シート「${sheetName}」・ ${table.rows.length}行 × ${table.headers.length}列`, false);
  }

  // ---------- Pipeline ----------

  function computeVariantGroups(settings) {
    const normalized = cleaner.normalizeAll(state.table.rows, settings.normalize, settings.placeholdersAsBlank);
    const groups = [];
    state.table.headers.forEach((header, column) => {
      for (const group of cleaner.findVariantGroups(normalized, column, settings.placeholdersAsBlank)) {
        const choiceKey = `${column}\u0001${group.key}`;
        if (!state.variantChoices.has(choiceKey)) {
          state.variantChoices.set(choiceKey, { enabled: true, representative: group.representative });
        }
        const choice = state.variantChoices.get(choiceKey);
        if (!group.variants.some((variant) => variant.value === choice.representative)) {
          choice.representative = group.representative;
        }
        groups.push({ column, header, choiceKey, ...group });
      }
    });
    return groups;
  }

  function variantRulesFrom(groups) {
    const rules = {};
    for (const group of groups) {
      const choice = state.variantChoices.get(group.choiceKey);
      if (!choice.enabled) continue;
      rules[group.column] = rules[group.column] || {};
      for (const variant of group.variants) rules[group.column][variant.value] = choice.representative;
    }
    return rules;
  }

  function recompute() {
    const settings = readSettings();
    state.variantGroups = computeVariantGroups(settings);
    settings.variantRules = variantRulesFrom(state.variantGroups);
    state.result = cleaner.cleanTable(state.table, settings);
    state.blankCounts = cleaner.blankCounts(state.table, settings.placeholdersAsBlank);
    state.placeholdersAsBlank = settings.placeholdersAsBlank;
    state.numericColumns = state.table.headers.map((_, column) => cleaner.isNumericColumn(state.table.rows, column));
    const editCount = Object.keys(state.manualEdits).length;
    const clearButton = element("clear-edits");
    clearButton.hidden = editCount === 0;
    clearButton.textContent = `手入力を取り消す（${editCount}）`;
    renderVariantGroups();
    renderSummary();
    renderBlankBars();
    renderView();
  }

  // ---------- Rendering ----------

  function renderColumnChips() {
    const build = (containerId, selection) => {
      const container = element(containerId);
      container.replaceChildren(
        ...state.table.headers.map((header, column) => {
          const label = make("label", "chip");
          const input = document.createElement("input");
          input.type = "checkbox";
          input.id = `${containerId}-${column}`;
          input.checked = selection.has(column);
          input.addEventListener("change", () => {
            if (input.checked) selection.add(column);
            else selection.delete(column);
            recompute();
          });
          label.append(input, document.createTextNode(header));
          return label;
        }),
      );
    };
    build("required-columns", state.requiredColumns);
    build("duplicate-columns", state.duplicateColumns);
  }

  function renderVariantGroups() {
    const container = element("variant-groups");
    if (state.variantGroups.length === 0) {
      container.replaceChildren(make("p", "muted", "表記揺れの候補は見つかりませんでした。"));
      return;
    }
    const byColumn = new Map();
    for (const group of state.variantGroups) {
      if (!byColumn.has(group.column)) byColumn.set(group.column, []);
      byColumn.get(group.column).push(group);
    }
    const sections = [];
    for (const [column, groups] of byColumn) {
      const section = make("div", "variant-column");
      section.append(make("h3", null, `${state.table.headers[column]}（${groups.length}組）`));
      for (const group of groups) {
        const choice = state.variantChoices.get(group.choiceKey);
        const box = make("div", "variant-group");
        const toggle = make("label", "check");
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = choice.enabled;
        checkbox.addEventListener("change", () => {
          choice.enabled = checkbox.checked;
          recompute();
        });
        const variants = make("span", "variants");
        for (const variant of group.variants) {
          const item = make("span", null, variant.value);
          item.append(make("small", null, ` ×${variant.count}`));
          variants.append(item);
        }
        toggle.append(checkbox, variants);
        const pick = make("label", "field-note", "統一先 ");
        const select = document.createElement("select");
        for (const variant of group.variants) select.append(new Option(variant.value, variant.value));
        select.value = choice.representative;
        select.disabled = !choice.enabled;
        select.addEventListener("change", () => {
          choice.representative = select.value;
          recompute();
        });
        pick.append(select);
        box.append(toggle, pick);
        section.append(box);
      }
      sections.push(section);
    }
    container.replaceChildren(...sections);
  }

  function renderSummary() {
    const { result, table } = state;
    const count = (type) => result.removed.filter((row) => row.type === type).length;
    element("stat-before").textContent = table.rows.length.toLocaleString();
    element("stat-after").textContent = result.table.rows.length.toLocaleString();
    element("stat-duplicates").textContent = count("duplicate").toLocaleString();
    element("stat-blank-rows").textContent = (count("empty") + count("required")).toLocaleString();
    element("stat-changes").textContent = result.changes.length.toLocaleString();
    element("stat-blank-cells").textContent = state.blankCounts.reduce((a, b) => a + b, 0).toLocaleString();
    element("count-changes").textContent = result.changes.length.toLocaleString();
    element("count-removed").textContent = result.removed.length.toLocaleString();
  }

  function renderBlankBars() {
    const total = Math.max(state.table.rows.length, 1);
    const nodes = [];
    state.table.headers.forEach((header, column) => {
      const blanks = state.blankCounts[column];
      const name = make("span", "name", header);
      name.title = header;
      const track = make("div", "track");
      const fill = make("div", "fill");
      fill.style.width = `${(blanks / total) * 100}%`;
      track.append(fill);
      const percent = Math.round((blanks / total) * 100);
      nodes.push(name, track, make("span", "value", `${blanks}件（${percent}%）`));
    });
    element("blank-bars").replaceChildren(...nodes);
  }

  function buildGrid(headers, extraHeaders) {
    const table = document.createElement("table");
    const headRow = document.createElement("tr");
    headRow.append(make("th", "rownum", "行"));
    for (const header of [...extraHeaders, ...headers]) headRow.append(make("th", null, header));
    const thead = document.createElement("thead");
    thead.append(headRow);
    const tbody = document.createElement("tbody");
    table.append(thead, tbody);
    return { table, tbody };
  }

  function setLegend(items) {
    element("legend").replaceChildren(...items.map(([className, text]) => make("span", className, text)));
  }

  function limitNote(shown, total) {
    element("table-note").textContent =
      total > shown ? `先頭${shown.toLocaleString()}行を表示しています（全${total.toLocaleString()}行）。ダウンロードには全行が含まれます。` : "";
  }

  function renderView() {
    const { result, table } = state;
    const container = element("table-container");
    document.querySelectorAll(".tab").forEach((tab) => {
      tab.setAttribute("aria-selected", String(tab.dataset.view === state.view));
    });
    const changeByCell = new Map(result.changes.map((change) => [`${change.origin}:${change.column}`, change]));

    if (state.view === "after") {
      setLegend([
        ["l-changed", "自動で修正したセル（修正前は「変更履歴」で確認）"],
        ["l-edited", "手入力したセル"],
        ["l-blank", "空欄"],
      ]);
      const { table: grid, tbody } = buildGrid(result.table.headers, []);
      const shownIndexes = result.table.rows
        .map((_, index) => index)
        .filter((index) => index < PREVIEW_ROW_LIMIT || state.pendingRows.has(result.rowOrigins[index]));
      for (const index of shownIndexes) {
        const cells = result.table.rows[index];
        const origin = result.rowOrigins[index];
        const tr = document.createElement("tr");
        tr.append(make("td", "rownum", String(index + 2)));
        cells.forEach((cell, column) => {
          const td = editableCell(cell, origin, column);
          const change = changeByCell.get(`${origin}:${column}`);
          if (change && change.kind === "manual") {
            td.classList.add("edited");
            td.title = `元の値: ${change.before === "" ? "（空）" : change.before}`;
          } else if (change) {
            td.classList.add("changed");
            td.title = `修正前: ${change.before === "" ? "（空）" : change.before}`;
          } else if (cell.trim() === "") {
            td.classList.add("blank");
          }
          tr.append(td);
        });
        tbody.append(tr);
      }
      container.replaceChildren(
        shownIndexes.length ? grid : make("p", "empty-state", "整理後に残った行はありません。「行を追加」で入力できます。"),
      );
      limitNote(shownIndexes.length, result.table.rows.length);
      return;
    }

    if (state.view === "before") {
      setLegend([
        ["l-removed", "削除される行"],
        ["l-changed", "修正されるセル"],
        ["l-edited", "手入力したセル"],
        ["l-blank", "空欄"],
      ]);
      const removedByOrigin = new Map(result.removed.map((row) => [row.origin, row]));
      const { table: grid, tbody } = buildGrid(table.headers, ["判定"]);
      const rows = table.rows.slice(0, PREVIEW_ROW_LIMIT);
      rows.forEach((cells, origin) => {
        const removed = removedByOrigin.get(origin);
        const tr = document.createElement("tr");
        if (removed) tr.className = "removed";
        tr.append(make("td", "rownum", String(origin + 2)));
        tr.append(make("td", "reason", removed ? removed.reason : ""));
        cells.forEach((cell, column) => {
          const td = editableCell(cell, origin, column);
          const change = changeByCell.get(`${origin}:${column}`);
          if (change && change.kind === "manual") {
            td.classList.add("edited");
            td.title = `入力値: ${change.after === "" ? "（空）" : change.after}`;
          } else if (change) {
            td.classList.add("changed");
            td.title = `修正後: ${change.after === "" ? "（空）" : change.after}`;
          } else if (cleaner.isBlank(cell, state.placeholdersAsBlank)) {
            td.classList.add("blank");
          }
          tr.append(td);
        });
        tbody.append(tr);
      });
      container.replaceChildren(grid);
      limitNote(rows.length, table.rows.length);
      return;
    }

    if (state.view === "changes") {
      setLegend([]);
      const { table: grid, tbody } = buildGrid(["列", "修正前", "修正後", "種類"], []);
      const changes = result.changes.slice(0, PREVIEW_ROW_LIMIT * 5);
      for (const change of changes) {
        const tr = document.createElement("tr");
        tr.append(make("td", "rownum", String(change.origin + 2)));
        tr.append(make("td", null, table.headers[change.column]));
        tr.append(make("td", "before-value", change.before === "" ? "（空）" : change.before));
        tr.append(make("td", null, change.after === "" ? "（空）" : change.after));
        tr.append(make("td", "kind", KIND_LABEL[change.kind]));
        tbody.append(tr);
      }
      container.replaceChildren(changes.length ? grid : make("p", "empty-state", "修正したセルはありません。"));
      element("table-note").textContent = "行番号は元ファイルの行番号です（見出し行が1行目）。";
      return;
    }

    setLegend([]);
    const { table: grid, tbody } = buildGrid(table.headers, ["削除理由"]);
    const removedRows = result.removed.slice(0, PREVIEW_ROW_LIMIT);
    for (const row of removedRows) {
      const tr = document.createElement("tr");
      tr.append(make("td", "rownum", String(row.origin + 2)));
      tr.append(make("td", "reason", row.reason));
      for (const cell of table.rows[row.origin]) tr.append(make("td", null, cell));
      tbody.append(tr);
    }
    container.replaceChildren(removedRows.length ? grid : make("p", "empty-state", "削除した行はありません。"));
    limitNote(removedRows.length, result.removed.length);
  }

  // ---------- Editing ----------

  function editableCell(value, origin, column) {
    const td = make("td", "editable", value);
    td.dataset.origin = String(origin);
    td.dataset.column = String(column);
    return td;
  }

  function startEdit(td) {
    if (td.classList.contains("editing")) return;
    const origin = Number(td.dataset.origin);
    const column = Number(td.dataset.column);
    const key = `${origin}:${column}`;
    const input = document.createElement("input");
    input.type = "text";
    input.id = `edit-${origin}-${column}`;
    input.value = Object.prototype.hasOwnProperty.call(state.manualEdits, key) ? state.manualEdits[key] : td.textContent;
    input.setAttribute("aria-label", `${state.table.headers[column]}（${origin + 2}行目）`);
    input.enterKeyHint = "done";
    if (state.numericColumns[column]) input.inputMode = "decimal";
    let finished = false;
    const finish = (commit) => {
      if (finished) return;
      finished = true;
      if (commit) {
        if (input.value === state.table.rows[origin][column]) delete state.manualEdits[key];
        else state.manualEdits[key] = input.value;
        recompute();
      } else {
        renderView();
      }
    };
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter" && !event.isComposing) {
        event.preventDefault();
        finish(true);
      } else if (event.key === "Escape") {
        finish(false);
      }
    });
    input.addEventListener("blur", () => finish(true));
    td.classList.add("editing");
    td.replaceChildren(input);
    input.focus();
    input.select();
  }

  function addRow() {
    const origin = state.table.rows.length;
    state.table.rows.push(state.table.headers.map(() => ""));
    state.pendingRows.add(origin);
    state.view = "after";
    recompute();
    const firstCell = element("table-container").querySelector(`td[data-origin="${origin}"][data-column="0"]`);
    if (firstCell) {
      firstCell.scrollIntoView({ block: "nearest" });
      startEdit(firstCell);
    }
  }

  function clearEdits() {
    state.manualEdits = {};
    recompute();
  }

  // ---------- Downloads ----------

  function outputBaseName() {
    const base = state.fileName.replace(/\.[^.]+$/, "") || "data";
    return `${base}_整理済み`;
  }

  async function saveFile(filename, data, mimeType) {
    const status = element("download-status");
    if (downloadsApi) {
      try {
        await downloadsApi.save({ filename, data });
        status.textContent = `${filename} を保存しました。`;
      } catch (error) {
        const code = error && error.code;
        status.textContent =
          code === "declined"
            ? "保存をキャンセルしました。"
            : code === "host_unavailable"
              ? "このアプリでは保存できません。ブラウザかパソコンで開いて保存してください。"
              : `保存できませんでした（${code || "unknown"}）。`;
      }
      return;
    }
    const blob = data instanceof Blob ? data : new Blob([data], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    status.textContent = `${filename} をダウンロードしました。`;
  }

  function downloadCSV() {
    const { headers, rows } = state.result.table;
    // A UTF-8 BOM makes Excel open the CSV with the right encoding.
    const csv = "﻿" + cleaner.toCSV(headers, rows);
    saveFile(`${outputBaseName()}.csv`, csv, "text/csv;charset=utf-8");
  }

  function downloadExcel() {
    if (!window.XLSX) {
      element("download-status").textContent = "Excel出力用のライブラリを読み込めませんでした。CSVをご利用ください。";
      return;
    }
    const { result, table } = state;
    const utils = window.XLSX.utils;
    const book = utils.book_new();
    const cleaned = [result.table.headers, ...result.table.rows.map((row) => row.map(cleaner.toSheetValue))];
    utils.book_append_sheet(book, utils.aoa_to_sheet(cleaned), "整理後");
    const changeLog = [
      ["元の行番号", "列", "修正前", "修正後", "種類"],
      ...result.changes.map((change) => [
        change.origin + 2,
        table.headers[change.column],
        change.before,
        change.after,
        KIND_LABEL[change.kind],
      ]),
    ];
    utils.book_append_sheet(book, utils.aoa_to_sheet(changeLog), "変更履歴");
    const removedLog = [
      ["元の行番号", "削除理由", ...table.headers],
      ...result.removed.map((row) => [row.origin + 2, row.reason, ...table.rows[row.origin]]),
    ];
    utils.book_append_sheet(book, utils.aoa_to_sheet(removedLog), "削除した行");
    const bytes = window.XLSX.write(book, { bookType: "xlsx", type: "array" });
    saveFile(
      `${outputBaseName()}.xlsx`,
      new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
  }

  // ---------- Wiring ----------

  function init() {
    element("file-input").addEventListener("change", (event) => {
      const [file] = event.target.files;
      if (file) loadFile(file);
      event.target.value = "";
    });
    element("load-sample").addEventListener("click", loadSample);
    element("sheet-select").addEventListener("change", (event) => loadSheet(event.target.value));
    const dropZone = document.body;
    const hasFiles = (event) => [...event.dataTransfer.types].includes("Files");
    dropZone.addEventListener("dragover", (event) => {
      if (!hasFiles(event)) return;
      event.preventDefault();
      element("drop-zone").classList.add("dragging");
    });
    dropZone.addEventListener("dragleave", () => element("drop-zone").classList.remove("dragging"));
    dropZone.addEventListener("drop", (event) => {
      if (!hasFiles(event)) return;
      event.preventDefault();
      element("drop-zone").classList.remove("dragging");
      const [file] = event.dataTransfer.files;
      if (file) loadFile(file);
    });
    for (const id of ["opt-width", "opt-trim", "opt-spaces", "opt-dashes", "opt-placeholders", "opt-empty-rows", "opt-duplicates"]) {
      element(id).addEventListener("change", recompute);
    }
    document.querySelectorAll(".tab").forEach((tab) => {
      tab.addEventListener("click", () => {
        state.view = tab.dataset.view;
        renderView();
      });
    });
    element("table-container").addEventListener("click", (event) => {
      const td = event.target.closest("td.editable");
      if (td) startEdit(td);
    });
    element("add-row").addEventListener("click", addRow);
    element("clear-edits").addEventListener("click", clearEdits);
    element("paste-apply").addEventListener("click", loadPastedText);
    element("download-csv").addEventListener("click", downloadCSV);
    element("download-xlsx").addEventListener("click", downloadExcel);
    loadSample();
  }

  init();
})();
