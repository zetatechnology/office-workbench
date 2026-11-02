# Office Workbench

VS Code 上で Word / Excel / PowerPoint（と PDF）ファイルを閲覧・操作するための拡張機能。

- 実装方針: [docs/implementation-plan.md](docs/implementation-plan.md)

## 現在の状態（Phase 1 / 1.5: 閲覧、Phase 2: Excel・CSV 編集）

`.docx` / `.xlsx` / `.pptx`（およびマクロ付きの `.docm` / `.xlsm` / `.pptm`）と `.pdf` をエディタで開いて表示します。Excel は編集・保存もできます（それ以外は読み取り専用）。

| 形式 | 表示内容 |
|---|---|
| Word | docx-preview によるページ表示、ズーム、幅に合わせる、ページ番号 |
| Excel | シート切替、数値書式・日付、フォント・塗り・罫線、結合セル、列幅・行高、数式バー、キーボード移動、仮想スクロール |
| PowerPoint | スライド一覧、レイアウト/マスター継承、テーマ色、図形・グラデーション・線・矢印、画像、表スタイル、ノート、非表示スライド |
| PDF | PDF.js による表示、ページ移動、ズーム（自動 / 幅 / ページ全体 / 倍率）、文字の検索・選択、目次（しおり）、日本語 CMap 対応 |

グラフ・SmartArt・条件付き書式などは未対応で、プレースホルダーまたは「Not shown」として表示します。

### Excel の編集

| 操作 | 方法 |
|---|---|
| セル入力 | セルを選んでそのまま入力（日本語 IME 可）、F2 / ダブルクリックで既存値を編集、数式バーでも編集可。Enter / Tab で確定、Esc で取り消し、Alt+Enter で改行、Ctrl+Enter で選択範囲に一括入力 |
| 入力値の解釈 | 数値・`1,234`・`12.5%`・日付（`2026/1/15`）・時刻・TRUE/FALSE・`=数式`。`'` で始めると文字列。書式が「標準」のセルには日付・パーセント等の表示形式を自動で付与 |
| 範囲選択 | ドラッグ、Shift+クリック / Shift+矢印、行・列見出しのクリック、Ctrl+A |
| クリア・コピー・貼り付け | Delete で内容をクリア、Ctrl+C / Ctrl+X / Ctrl+V（Excel 互換のタブ区切りテキスト） |
| 行・列 | 右クリックメニューから挿入・削除（数式・結合セル・グラフの参照範囲・ウィンドウ枠の固定なども追従） |
| シート | タブの「+」で追加、ダブルクリックで名前変更、右クリックで削除（他シート・名前定義・グラフの参照も更新） |
| 元に戻す・保存 | VS Code の Ctrl+Z / Ctrl+Y / Ctrl+S。未保存の変更は VS Code のバックアップ（Hot Exit）で保持 |

### CSV / TSV

`.csv` / `.tsv` / `.tab` は Excel と同じ表の画面で表示・編集できます。CSV はテキストとして編集することも多いため、通常どおりテキストエディタで開き、エディタ右上の「Open as Table」（表アイコン）またはエクスプローラーの右クリックで表に切り替えます（表からは「Open as Text」で戻ります）。常に表で開きたい場合は設定 `workbench.editorAssociations` に `"*.csv": "officeWorkbench.csv"` を追加します。

- 文字コードは UTF-8（BOM あり / なし）、Shift_JIS（Windows-31J）、UTF-16 を自動判別し、保存時も同じ文字コード・区切り文字・改行コードで書き戻します。Shift_JIS で表せない文字（絵文字など）を含むと保存できません
- 編集した行だけを書き直し、それ以外の行は元のテキスト（引用符の付け方も含む）のまま残します
- 値は文字列としてそのまま保存します（`=1+2` も数式ではなく文字列）。数値に見える値は右寄せで表示します

保存は「変更したパートだけを書き換える」方式で、それ以外の ZIP エントリ（グラフ・画像・マクロ等）は元のバイト列のまま残します。数式は再計算しないため、入力した数式の結果は Excel で開いたときに計算されます（保存時に `fullCalcOnLoad` を設定し、`calcChain.xml` を削除）。

## 開発

```bash
npm install
npm run build        # 拡張とWebviewをバンドル（packages/extension/dist）
npm test             # core のテスト（Vitest）
npm run lint
npm run typecheck
```

- **VS Code で動作確認**: このフォルダを VS Code で開き、F5（Run Extension）。`test-fixtures/` を開いた拡張開発ホストが起動します。
- **ブラウザで Webview を確認**: `npm run build && npm run preview` → http://localhost:5178/ （`-- --dir <フォルダ>` で任意のファイルを表示）。Excel は編集でき、Ctrl+Z / Ctrl+Y で元に戻す・やり直し、Ctrl+S で一時フォルダ（`%TEMP%/office-workbench-preview`）へ保存します（元ファイルは変更しません）
- **VSIX 作成**: `npm run package`
- **テスト用ファイルの再生成**: `npm run fixtures`（Windows + Microsoft Office が必要。`sample.pdf` は Word から書き出し。生成物はコミット済み）

他の PDF 拡張機能と併用している場合、VS Code がどちらで開くかを尋ねます。右クリックの「Reopen Editor With...」で切り替えられます。

## 構成

```
packages/
  core/       OOXML 読み書き（xlsx の編集・保存、pptx の読み取り）、CSV / TSV の読み書きと文字コード、数値書式、Webview とのメッセージ定義。VS Code 非依存
  extension/  VS Code 拡張本体（Excel・CSV は CustomEditorProvider、他は CustomReadonlyEditorProvider）
  webview/    Webview UI（Word / Excel / PowerPoint / PDF ビューア）
scripts/      ビルド、プレビュー用サーバー、フィクスチャ生成
test-fixtures/ Office で作成したサンプルファイル（CSV は Excel から Shift_JIS / UTF-8 / UTF-16 で書き出したもの）
```
