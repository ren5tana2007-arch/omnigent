# ととのえ表（CSV・Excel データ整理ツール）

ブラウザだけで動く表データの整理ツールです。ファイルはサーバーに送信せず、
すべて利用者の端末内で処理します。

## できること（MVP）

| 機能 | 内容 |
| --- | --- |
| 読み込み | CSV / TSV（UTF-8・Shift_JIS を自動判別、区切り文字を自動判別）、Excel（.xlsx / .xls、シート選択） |
| 重複の検出 | 全列または指定した列で比較し、2件目以降を削除（どの行と重複したかを記録） |
| 空欄の検出 | 列ごとの空欄数・空欄率を表示。全空欄行と、必須列が空欄の行を削除 |
| 表記揺れの整理 | 全角/半角・前後や連続の空白・ハイフン類を正規化。大文字小文字や「(株)/株式会社」の違いを候補として表示し、統一先を選んで一括置換 |
| 整理前後の比較 | 整理後（修正セルを強調）、整理前（削除行・修正セルを表示）、変更履歴、削除した行 |
| ダウンロード | CSV（BOM付きUTF-8、Excelで文字化けしない）、Excel（整理後・変更履歴・削除した行の3シート） |

## 使い方

```bash
cd csv-cleaner
python3 -m http.server 8000   # http://localhost:8000 を開く
```

`index.html` を直接開いても動きます（Excel 読み書きと Web フォントは CDN から読み込みます）。

1ファイルにまとめた版を作るとき:

```bash
node build-single-file.mjs dist.html
```

## テスト

```bash
node --test csv-cleaner/tests/*.test.cjs
```

## ファイル構成

- `cleaner.js` — 整理ロジック（DOM 非依存。ブラウザと Node の両方で動く）
- `app.js` — 画面の描画と操作
- `index.html` / `styles.css` — 画面
- `build-single-file.mjs` — 1ファイル版の生成
- `tests/` — `cleaner.js` の単体テスト
