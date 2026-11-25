/*
 * Core data-cleaning logic (no DOM access). Loaded as a classic browser script
 * (exposes window.DataCleaner) and as a CommonJS module for Node tests.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) {
    module.exports = api;
  } else {
    root.DataCleaner = api;
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const DEFAULT_NORMALIZE = Object.freeze({
    widthNormalize: true,
    trim: true,
    collapseSpaces: true,
    unifyDashes: true,
  });

  // Values that spreadsheets commonly use to mean "no data".
  const PLACEHOLDER_BLANKS = new Set(["-", "null", "nan", "n/a", "#n/a", "none"]);
  const DASH_LIKE = /[‐‑‒–—―−﹣－]/g;
  const LONG_VOWEL_BETWEEN_DIGITS = /([0-9])[ーｰ](?=[0-9])/g;
  const RUN_OF_SPACES = /[ \t　]+/g;
  const PLAIN_NUMBER = /^-?(0|[1-9]\d{0,13})(\.\d+)?$/;

  /** Decode file bytes as UTF-8, falling back to Shift_JIS (common for Excel CSVs in Japan). */
  function decodeText(buffer) {
    const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    let text;
    let encoding = "UTF-8";
    try {
      text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    } catch (_error) {
      text = new TextDecoder("shift_jis").decode(bytes);
      encoding = "Shift_JIS";
    }
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    return { text, encoding };
  }

  /** Pick the delimiter (comma, tab or semicolon) that appears most in the first lines. */
  function detectDelimiter(text) {
    const sample = text.split(/\r?\n/, 5).join("\n");
    const candidates = [",", "\t", ";"];
    let best = ",";
    let bestCount = 0;
    for (const candidate of candidates) {
      const count = sample.split(candidate).length - 1;
      if (count > bestCount) {
        best = candidate;
        bestCount = count;
      }
    }
    return best;
  }

  /** RFC 4180 style CSV parser: quoted fields, escaped quotes, CRLF/LF, newlines inside quotes. */
  function parseCSV(text, delimiter) {
    const separator = delimiter || detectDelimiter(text);
    const rows = [];
    let row = [];
    let field = "";
    let inQuotes = false;
    let index = 0;
    while (index < text.length) {
      const character = text[index];
      if (inQuotes) {
        if (character === '"') {
          if (text[index + 1] === '"') {
            field += '"';
            index += 2;
            continue;
          }
          inQuotes = false;
          index += 1;
          continue;
        }
        field += character;
        index += 1;
        continue;
      }
      if (character === '"' && field === "") {
        inQuotes = true;
      } else if (character === separator) {
        row.push(field);
        field = "";
      } else if (character === "\r" || character === "\n") {
        if (character === "\r" && text[index + 1] === "\n") index += 1;
        row.push(field);
        rows.push(row);
        row = [];
        field = "";
      } else {
        field += character;
      }
      index += 1;
    }
    if (field !== "" || row.length > 0) {
      row.push(field);
      rows.push(row);
    }
    return rows;
  }

  /** Turn a 2D array into {headers, rows}: first row is the header, rows padded to equal width. */
  function buildTable(matrix) {
    const width = matrix.reduce((max, row) => Math.max(max, row.length), 0);
    const headerRow = matrix[0] || [];
    const seen = new Map();
    const headers = [];
    for (let column = 0; column < width; column += 1) {
      let name = String(headerRow[column] ?? "").trim() || `列${column + 1}`;
      const count = seen.get(name) || 0;
      seen.set(name, count + 1);
      if (count > 0) name = `${name}(${count + 1})`;
      headers.push(name);
    }
    const rows = matrix.slice(1).map((row) => {
      const cells = [];
      for (let column = 0; column < width; column += 1) {
        cells.push(row[column] == null ? "" : String(row[column]));
      }
      return cells;
    });
    return { headers, rows };
  }

  /** Normalize notation: NFKC width folding, dash unification, trimming and space collapsing. */
  function normalizeText(value, options) {
    const settings = options || DEFAULT_NORMALIZE;
    let text = value == null ? "" : String(value);
    if (settings.widthNormalize) text = text.normalize("NFKC");
    if (settings.unifyDashes) {
      text = text.replace(DASH_LIKE, "-").replace(LONG_VOWEL_BETWEEN_DIGITS, "$1-");
    }
    if (settings.trim) text = text.trim();
    if (settings.collapseSpaces) text = text.replace(RUN_OF_SPACES, " ");
    return text;
  }

  function isPlaceholderBlank(value) {
    return PLACEHOLDER_BLANKS.has(String(value).trim().toLowerCase());
  }

  function isBlank(value, placeholdersAsBlank) {
    const text = value == null ? "" : String(value);
    if (text.trim() === "") return true;
    return Boolean(placeholdersAsBlank) && isPlaceholderBlank(text);
  }

  /** Loose comparison key used to group spelling variants of the same value. */
  function variantKey(value) {
    return normalizeText(value, DEFAULT_NORMALIZE)
      .toLowerCase()
      .replace(/\(株\)/g, "株式会社")
      .replace(/\(有\)/g, "有限会社")
      .replace(/\s+/g, "");
  }

  /**
   * Find groups of differently written values that share a variantKey in one column.
   * Returns [{key, variants: [{value, count}], representative}] sorted by impact.
   */
  function findVariantGroups(rows, column, placeholdersAsBlank) {
    const byKey = new Map();
    for (const row of rows) {
      const value = row[column];
      if (isBlank(value, placeholdersAsBlank)) continue;
      const key = variantKey(value);
      if (!byKey.has(key)) byKey.set(key, new Map());
      const counts = byKey.get(key);
      counts.set(value, (counts.get(value) || 0) + 1);
    }
    const groups = [];
    for (const [key, counts] of byKey) {
      if (counts.size < 2) continue;
      const variants = [...counts.entries()]
        .map(([value, count]) => ({ value, count }))
        .sort((left, right) => right.count - left.count);
      groups.push({ key, variants, representative: variants[0].value });
    }
    return groups.sort(
      (left, right) =>
        right.variants.reduce((sum, item) => sum + item.count, 0) -
        left.variants.reduce((sum, item) => sum + item.count, 0),
    );
  }

  function normalizeAll(rows, normalizeOptions, placeholdersAsBlank) {
    return rows.map((row) =>
      row.map((cell) => {
        const normalized = normalizeText(cell, normalizeOptions);
        return placeholdersAsBlank && isPlaceholderBlank(normalized) ? "" : normalized;
      }),
    );
  }

  function duplicateSignature(cells, keyColumns) {
    return keyColumns.map((column) => String(cells[column] ?? "").trim().toLowerCase()).join("\u0001");
  }

  /**
   * Run the cleaning pipeline.
   * settings: {normalize, placeholdersAsBlank, removeEmptyRows, requiredColumns,
   *            variantRules: {[column]: {[value]: representative}},
   *            removeDuplicates, duplicateKeyColumns}
   * Row numbers in the result are spreadsheet rows (header = row 1).
   */
  function cleanTable(table, settings) {
    const options = Object.assign(
      {
        normalize: DEFAULT_NORMALIZE,
        placeholdersAsBlank: true,
        removeEmptyRows: true,
        requiredColumns: [],
        variantRules: {},
        removeDuplicates: true,
        duplicateKeyColumns: [],
      },
      settings,
    );
    const columnCount = table.headers.length;
    const changes = [];
    const working = table.rows.map((original, originIndex) => {
      const cells = original.map((before, column) => {
        let after = options.normalize ? normalizeText(before, options.normalize) : before;
        let kind = "normalize";
        if (options.placeholdersAsBlank && after !== "" && isPlaceholderBlank(after)) {
          after = "";
          kind = "blank";
        }
        const rule = options.variantRules[column];
        if (rule && Object.prototype.hasOwnProperty.call(rule, after) && rule[after] !== after) {
          after = rule[after];
          kind = "variant";
        }
        if (after !== before) changes.push({ origin: originIndex, column, before, after, kind });
        return after;
      });
      return { origin: originIndex, cells };
    });

    const removed = [];
    const survivors = [];
    for (const row of working) {
      const blankFlags = row.cells.map((cell) => cell.trim() === "");
      if (options.removeEmptyRows && blankFlags.every(Boolean)) {
        removed.push({ origin: row.origin, cells: row.cells, reason: "空行", type: "empty" });
        continue;
      }
      const missing = options.requiredColumns.filter((column) => blankFlags[column]);
      if (missing.length > 0) {
        const names = missing.map((column) => `「${table.headers[column]}」`).join("");
        removed.push({ origin: row.origin, cells: row.cells, reason: `必須列${names}が空欄`, type: "required" });
        continue;
      }
      survivors.push(row);
    }

    const kept = [];
    if (options.removeDuplicates) {
      const keyColumns = options.duplicateKeyColumns.length
        ? options.duplicateKeyColumns
        : [...Array(columnCount).keys()];
      const firstSeen = new Map();
      for (const row of survivors) {
        const signature = duplicateSignature(row.cells, keyColumns);
        if (firstSeen.has(signature)) {
          const firstRowNumber = firstSeen.get(signature) + 2;
          removed.push({
            origin: row.origin,
            cells: row.cells,
            reason: `重複（${firstRowNumber}行目と同じ）`,
            type: "duplicate",
          });
          continue;
        }
        firstSeen.set(signature, row.origin);
        kept.push(row);
      }
    } else {
      kept.push(...survivors);
    }

    removed.sort((left, right) => left.origin - right.origin);
    const keptOrigins = new Set(kept.map((row) => row.origin));
    return {
      table: { headers: table.headers.slice(), rows: kept.map((row) => row.cells) },
      rowOrigins: kept.map((row) => row.origin),
      changes: changes.filter((change) => keptOrigins.has(change.origin)),
      removed,
    };
  }

  /** Count blank cells per column. */
  function blankCounts(table, placeholdersAsBlank) {
    const counts = table.headers.map(() => 0);
    for (const row of table.rows) {
      row.forEach((cell, column) => {
        if (isBlank(cell, placeholdersAsBlank)) counts[column] += 1;
      });
    }
    return counts;
  }

  function escapeCSVField(value) {
    const text = value == null ? "" : String(value);
    return /[",\r\n]|^\s|\s$/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  }

  /** Serialize to CSV with CRLF line endings (what Excel expects). */
  function toCSV(headers, rows) {
    return [headers, ...rows].map((row) => row.map(escapeCSVField).join(",")).join("\r\n") + "\r\n";
  }

  /** Convert plain numbers to numbers for Excel output; keep leading-zero codes as text. */
  function toSheetValue(value) {
    const text = value == null ? "" : String(value);
    return PLAIN_NUMBER.test(text) ? Number(text) : text;
  }

  return {
    DEFAULT_NORMALIZE,
    decodeText,
    detectDelimiter,
    parseCSV,
    buildTable,
    normalizeText,
    normalizeAll,
    isBlank,
    isPlaceholderBlank,
    variantKey,
    findVariantGroups,
    cleanTable,
    blankCounts,
    toCSV,
    toSheetValue,
  };
});
