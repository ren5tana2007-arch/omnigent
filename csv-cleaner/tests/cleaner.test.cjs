const test = require("node:test");
const assert = require("node:assert/strict");
const cleaner = require("../cleaner.js");

test("parseCSV handles quotes, escaped quotes, CRLF and newlines inside fields", () => {
  const text = 'id,name,memo\r\n1,"Smith, John","He said ""hi"""\r\n2,Taro,"line1\nline2"\r\n';
  assert.deepEqual(cleaner.parseCSV(text), [
    ["id", "name", "memo"],
    ["1", "Smith, John", 'He said "hi"'],
    ["2", "Taro", "line1\nline2"],
  ]);
});

test("detectDelimiter recognizes tab-separated text", () => {
  assert.equal(cleaner.detectDelimiter("a\tb\tc\n1\t2\t3"), "\t");
  assert.equal(cleaner.detectDelimiter("a,b\n1,2"), ",");
});

test("decodeText falls back to Shift_JIS when bytes are not valid UTF-8", () => {
  // "東京" encoded as Shift_JIS.
  const shiftJisBytes = new Uint8Array([0x93, 0x8c, 0x8b, 0x9e]);
  assert.deepEqual(cleaner.decodeText(shiftJisBytes), { text: "東京", encoding: "Shift_JIS" });
  const utf8WithBom = new Uint8Array([0xef, 0xbb, 0xbf, 0x61]);
  assert.deepEqual(cleaner.decodeText(utf8WithBom), { text: "a", encoding: "UTF-8" });
});

test("buildTable names empty and duplicate headers and pads short rows", () => {
  const table = cleaner.buildTable([["名前", "", "名前"], ["a"]]);
  assert.deepEqual(table.headers, ["名前", "列2", "名前(2)"]);
  assert.deepEqual(table.rows, [["a", "", ""]]);
});

test("normalizeText folds width, unifies dashes and spaces", () => {
  assert.equal(cleaner.normalizeText("  ＡＢＣ　 ｶﾌﾞ  "), "ABC カブ");
  assert.equal(cleaner.normalizeText("０３－１２３４ー５６７８"), "03-1234-5678");
  // A long-vowel mark inside katakana words must stay.
  assert.equal(cleaner.normalizeText("コーヒー"), "コーヒー");
});

test("findVariantGroups groups case, spacing and company-form variants", () => {
  const rows = [["株式会社サクラ"], ["(株)サクラ"], ["株式会社サクラ"], ["Standard"], ["standard"], ["Lite"]];
  const groups = cleaner.findVariantGroups(rows, 0, true);
  assert.equal(groups.length, 2);
  assert.equal(groups[0].representative, "株式会社サクラ");
  assert.deepEqual(
    groups[0].variants.map((variant) => variant.value),
    ["株式会社サクラ", "(株)サクラ"],
  );
});

test("cleanTable normalizes, unifies variants, removes blank and duplicate rows", () => {
  const table = {
    headers: ["id", "company", "plan"],
    rows: [
      ["1", "株式会社サクラ", "Standard"],
      ["2", "（株）サクラ", "standard"],
      ["", "", " "],
      ["3", "", "Lite"],
      ["１", "株式会社サクラ ", "Standard"],
      ["4", "N/A", "Lite"],
    ],
  };
  const result = cleaner.cleanTable(table, {
    requiredColumns: [1],
    variantRules: { 1: { "(株)サクラ": "株式会社サクラ" }, 2: { standard: "Standard" } },
  });
  assert.deepEqual(result.table.rows, [
    ["1", "株式会社サクラ", "Standard"],
    ["2", "株式会社サクラ", "Standard"],
  ]);
  assert.deepEqual(
    result.removed.map((row) => [row.origin + 2, row.type]),
    [
      [4, "empty"],
      [5, "required"],
      [6, "duplicate"],
      [7, "required"],
    ],
  );
  assert.equal(result.removed[2].reason, "重複（2行目と同じ）");
  assert.deepEqual(
    result.changes.map((change) => [change.origin, change.column, change.kind]),
    [
      [1, 1, "variant"],
      [1, 2, "variant"],
    ],
  );
});

test("cleanTable compares duplicates only on the chosen key columns", () => {
  const table = { headers: ["email", "name"], rows: [["a@example.com", "A"], ["A@example.com", "B"]] };
  const byEmail = cleaner.cleanTable(table, { duplicateKeyColumns: [0] });
  assert.equal(byEmail.table.rows.length, 1);
  const byAll = cleaner.cleanTable(table, {});
  assert.equal(byAll.table.rows.length, 2);
});

test("toCSV quotes fields that need it and toSheetValue keeps leading zeros", () => {
  assert.equal(cleaner.toCSV(["a", "b"], [['x,y', 'say "hi"']]), 'a,b\r\n"x,y","say ""hi"""\r\n');
  assert.equal(cleaner.toSheetValue("0123"), "0123");
  assert.equal(cleaner.toSheetValue("1200.5"), 1200.5);
});
