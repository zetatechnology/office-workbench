# Office Workbench 実装方針（ドラフト）

VS Code 上で Word / Excel / PowerPoint ファイルを閲覧・操作する拡張機能の実装方針。

## 1. ゴールとスコープ

| 区分 | 内容 |
|---|---|
| 対象形式（主） | `.docx` / `.xlsx` / `.pptx`（OOXML）。マクロ付き `.docm/.xlsm/.pptm` は「マクロを保持するが実行しない」 |
| 対象形式（副） | `.xls` は読み取り専用（SheetJS）。`.doc/.ppt` は LibreOffice 連携がある場合のみ変換表示 |
| 閲覧 | VS Code のエディタタブで開いて表示（ズーム・ページ/シート/スライド切替・検索） |
| 操作 | セル・テキストの編集、行列/スライドの追加削除などの基本編集、Undo/Redo、保存 |
| 開発者向け | テキスト抽出、Git 差分表示、コマンド/Language Model Tool からの読み書き |
| 非ゴール（当面） | Office 完全互換のレイアウト編集、数式の完全再計算エンジン、マクロ実行 |

## 2. 基本方針

1. **Office 非依存・純 TypeScript を核にする**
   Office のインストール有無・OS に依存せず、Remote/WSL/Codespaces/vscode.dev（Web 拡張）でも動くことを優先する。COM 自動化や LibreOffice は「任意のアダプタ」として後付けする。
2. **ラウンドトリップ安全性を最優先（最小差分パッチ方式）**
   ExcelJS 等の「モデルに読み込んで全体を書き直す」ライブラリは、グラフ・ピボット・未対応要素を保存時に落とす危険がある。
   編集時は ZIP 内の **変更したパートの、変更した XML ノードだけ** を書き換え、それ以外のパートはバイト単位で保持する。
3. **表示は既存ライブラリ、編集は自前の OOXML レイヤ**
   閲覧の忠実度は実績あるレンダラに任せ、書き込みは自前の薄い OOXML 操作層（JSZip + XML DOM）で制御する。
4. **正（canonical）の状態は Extension Host が持つ**
   Webview は表示と入力のみ。編集は「操作（Operation）」として Extension Host に送り、そこで適用・Undo 管理・保存する。

## 3. アーキテクチャ

```
┌──────────────────────── VS Code ────────────────────────┐
│ Extension Host (Node / Web Worker)                      │
│  ├ CustomEditorProvider (docx / xlsx / pptx)            │
│  │   └ OfficeDocument : CustomDocument                  │
│  │       ├ OoxmlPackage (JSZip + パート単位の XML DOM)    │
│  │       ├ Operation 適用 / Undo・Redo スタック           │
│  │       └ save / saveAs / revert / backup              │
│  ├ Commands（テキスト抽出・エクスポート・変換）              │
│  ├ Language Model Tools（vscode.lm.registerTool）         │
│  └ Adapters（任意）: LibreOffice headless / Office COM     │
│              ▲ postMessage（Operation / 差分通知）         │
│              ▼                                          │
│ Webview UI（形式ごとのビュー）                             │
│  ├ Word  : docx-preview によるページ表示 + 段落編集         │
│  ├ Excel : 仮想スクロールのグリッド（自前） + セル編集        │
│  └ PPT   : DrawingML サブセットの自前 SVG/HTML レンダラ     │
└─────────────────────────────────────────────────────────┘
```

### パッケージ構成（単一リポジトリ・npm workspaces 想定）

```
packages/
  core/        # OOXML 読み書き・Operation 定義。VS Code 非依存、Node 単体でテスト可能
  extension/   # VS Code 拡張本体（CustomEditor, commands, LM tools）
  webview/     # Webview UI（Word/Excel/PPT ビュー）
test-fixtures/ # 各形式のサンプルファイル（ラウンドトリップ検証用）
```

`core` を VS Code から切り離すことで、ラウンドトリップテスト（開く→編集→保存→再読込→未変更パートが同一か）を高速に回せる。

## 4. 形式ごとの方針

### Excel（.xlsx）— 最初に編集まで作り込む
- **読み込み**: 自前パーサ（`workbook.xml` / `sheetN.xml` / `sharedStrings.xml` / `styles.xml`）。巨大シート対応のためシート単位で遅延パース。
- **表示**: 仮想スクロールのグリッド（数万行でも可視範囲のみ描画）。数値書式・列幅・結合セル・基本スタイルを再現。
- **編集**: セル値・数式文字列の入力、行列挿入削除、シート追加/名前変更。数式は **値を再計算せず** `<f>` を書いてキャッシュ値を消し、`fullCalcOnLoad` で Excel 側に再計算させる（将来 HyperFormula 等で簡易計算を検討）。
- **補助**: `.csv` / `.tsv` は自前のパーサで表示・編集（実装済み）。`.xls` は SheetJS で読み取り専用表示。

### Word（.docx）
- **表示**: `docx-preview`（Apache-2.0、ページ単位の高忠実度 HTML 化）。
- **編集（段階的）**: ① 段落テキストの置換・追記 → ② 段落の追加削除・見出しスタイル適用 → ③ 表セル編集。
  run（`<w:r>`）分割を考慮し、編集した段落の書式を可能な限り維持する。
- 変更履歴（track changes）は表示のみ対応から開始。

### PowerPoint（.pptx）— 最も難易度が高い
- **表示**: 既存 OSS は忠実度・保守性に難があるため、**DrawingML のサブセット（テキストボックス・図形・画像・表・テーマ色）を自前で SVG/HTML 化**。スライド一覧サムネイル + メイン表示。
- 高忠実度が必要な場合は LibreOffice アダプタで PDF/画像化してプレビュー（任意機能）。
- **編集**: テキストの編集 → スライドの複製/削除/並べ替え → 図形の移動・サイズ変更。

## 5. VS Code 統合

| 機能 | API / 方式 |
|---|---|
| エディタ | `CustomEditorProvider`（編集対応）。`retainContextWhenHidden` は大容量時のメモリに注意し状態復元で対応 |
| Undo/Redo・dirty | `onDidChangeCustomDocument` に Operation 単位の edit を発火 |
| Hot exit | `backup()` で ZIP を一時保存 |
| ファイル I/O | `vscode.workspace.fs`（Remote / Web で共通動作） |
| テキスト抽出 | コマンド「プレーンテキスト / Markdown として開く」 |
| Git 差分 | SCM の差分で Office ファイル同士を比較するビュー（テキスト化差分 → 将来セル/段落単位の差分 UI） |
| AI 連携 | `vscode.lm.registerTool` で「読み取り（シート範囲・段落・スライドのテキスト）」「編集（Operation 適用）」を公開。Copilot 等のエージェントから Office ファイルを扱えるようにする |
| Web 拡張 | `browser` エントリを用意し、Node API 依存はアダプタ層に隔離 |

## 6. 技術スタック（推奨）

| 用途 | 候補 | 備考 |
|---|---|---|
| 言語 / ビルド | TypeScript + esbuild | 拡張本体と Webview を別バンドル |
| ZIP | JSZip（MIT） | Web 互換 |
| XML | `@xmldom/xmldom` | DOM で部分パッチしやすい。名前空間・属性順を保持 |
| Word 表示 | docx-preview | |
| Excel 補助 | SheetJS CE（xls 読み取り） | npm 公開版は古いため公式 CDN tarball を固定バージョンで利用 |
| Webview UI | Preact（または素の TS） | バンドルを軽く保つ。`@vscode/webview-ui-toolkit` は非推奨化済みのため使わず、VS Code の CSS 変数でテーマ追従 |
| テスト | Vitest（core）、`@vscode/test-electron`（統合） | ラウンドトリップテストを CI 必須に |
| パッケージ | `@vscode/vsce` | |

## 7. セキュリティ・堅牢性
- Webview は厳格な CSP（nonce 付きスクリプトのみ、外部通信禁止）。
- マクロ（`vbaProject.bin`）・ActiveX・OLE は実行しない。保持のみ。
- 外部リンク・外部画像参照は自動取得しない。
- ZIP bomb 対策（展開サイズ・ファイル数上限）、XXE 無効化。
- 大容量ファイルは遅延パース + 閾値超えで読み取り専用モードへ。

## 8. マイルストーン

| Phase | 内容 | 完了条件 |
|---|---|---|
| 1 | **3 形式の閲覧**（雛形構築を含む） | `.docx` / `.xlsx` / `.pptx` をダブルクリックで開き、内容を表示できる |
| 1.5 | **PDF の閲覧**（PDF.js） | `.pdf` をページ移動・ズーム・検索・目次付きで表示できる |
| 2 | Excel 編集（セル・行列・シート）+ Undo/保存 | 編集後に Excel で開いて破損なし、未変更パートがバイト一致 |
| 3 | Word / PPT のテキスト編集 | 同上 |
| 4 | テキスト抽出・Git 差分・LM Tools | エージェントから読み書きできる |
| 5 | 任意アダプタ（LibreOffice 変換、Windows の Office COM） | 設定で有効化した場合のみ動作 |

### Phase 1 の詳細（閲覧専用）

まず「3 形式すべてが VS Code 上で見られる」状態を最短で作る。編集機能は入れず、`CustomReadonlyEditorProvider` で実装する（Phase 2 で `CustomEditorProvider` に置き換える）。

| 項目 | 内容 |
|---|---|
| 雛形 | npm workspaces（`core` / `extension` / `webview`）、esbuild、ESLint、Vitest、GitHub Actions |
| 共通 | 3 形式の CustomEditor 登録、Webview の CSP、VS Code テーマ追従、読み込み中・エラー表示、ファイル変更時の再読込 |
| Word | docx-preview でページ表示。ズーム、ページ送り |
| Excel | 自前パーサ + 仮想スクロールのグリッド。シートタブ切替、値と基本書式（数値書式・太字・背景色・列幅・結合セル）、行列見出しの固定 |
| PowerPoint | 自前レンダラでテキスト・図形（矩形/楕円等の基本形状）・画像・表を SVG/HTML 化。サムネイル一覧 + 選択スライド表示 |
| テスト | 各形式の fixture を数種用意し、core のパース結果をスナップショットで検証 |

**Phase 1 で割り切ること**: グラフ・SmartArt・数式（Word の数式エディタ）・アニメーション・条件付き書式は「未対応の要素」としてプレースホルダ表示。Excel の数式はファイル内のキャッシュ値を表示する。

**完了条件**: 代表的な fixture（各形式 3〜5 個）が崩れずに読める程度に表示でき、10 万セル規模の xlsx でもスクロールが実用的に動くこと。

**実装メモ（Phase 1 時点）**
- Webview UI は Preact を使わず素の TypeScript で実装（各ビューアのバンドルは Excel / PowerPoint とも約 10 KB）。
- xlsx / pptx の解析は Extension Host 側（`core`）で行い、表示用モデルを Webview に送る。docx はバイト列を送って Webview 内の docx-preview で描画。
- 10 万セルのシート解析は約 0.8 秒（テストで計測）。
- fixture は `scripts/make-fixtures.ps1` で実際の Office（COM）から生成。現状は各形式 1 ファイルで、実ファイルでの確認を追加していく。
- 未実装: Excel のウィンドウ枠の固定（行列見出しは固定済み）、PowerPoint のグラフ以外の図形の一部（未知のプリセットは矩形で代替）。共有数式の数式バー表示は Phase 2 で対応。

### Phase 2 の詳細（Excel 編集）

| 項目 | 内容 |
|---|---|
| VS Code 統合 | xlsx を `CustomEditorProvider` に置き換え。編集ごとに `onDidChangeCustomDocument` を発火（Undo/Redo・dirty 表示）、save / saveAs / revert / backup（Hot Exit）。未保存の変更がなければ外部での変更を再読込 |
| 操作（`EditOp`） | `setCells`（値・数式・クリア）、`clearRange`、`insertRows/deleteRows/insertCols/deleteCols`、`addSheet/renameSheet/deleteSheet`。Webview は操作を送るだけで、適用・Undo は Extension Host（`WorkbookSession`）が行う |
| セル編集 | `<c>` 要素だけを書き換え（スタイル・拡張属性は保持）。文字列は `inlineStr` で書き、sharedStrings は変更しない。日付・% などは必要に応じて `cellXfs` に書式を追加。共有数式の一部を編集するとそのグループを通常の数式に展開 |
| 行列の挿入削除 | セル位置、全シート・名前定義・グラフ系列の数式参照、結合セル、条件付き書式・入力規則・ハイパーリンクの範囲、列幅、ウィンドウ枠の固定、テーブル範囲を移動。テーブルを分断する操作は拒否 |
| シート操作 | workbook.xml / rels / `[Content_Types].xml` を更新。名前変更・削除で他シートの参照、名前定義、グラフ、ピボットのソースを更新（削除は `#REF!`） |
| Undo | セル編集は変更前の `<c>` を保存して復元。構造変更はパッケージのジャーナル（触ったパートの変更前内容）を記録して復元 |
| 保存時の処理 | 変更シートの `<dimension>` 再計算、`calcChain.xml` 削除、`calcPr@fullCalcOnLoad="1"`。変更していないパートは元の圧縮バイト列のまま |

**検証**: core のテスト（未変更パートのバイト一致、Undo、参照の移動）に加え、セル編集・行列操作・シート操作を組み合わせたファイルと VS Code から保存したファイルを Excel（COM）で開き、修復なしで開けること・再計算結果が正しいことを確認。VS Code 上では CDP 経由の実キー入力で、入力・IME・Delete・Ctrl+Z/Y/S、行挿入、シート追加、外部変更の再読込を確認。

**CSV / TSV（Phase 2 の追加）**: 同じグリッドと `SheetSession` インターフェースを使う `CsvSession` で 1 シートのブックとして扱う。拡張子 `.csv/.tsv/.tab` の CustomEditor（priority `option`）で、テキストエディタとの切り替えコマンドをエディタタイトルに追加。文字コード（UTF-8 / BOM / Shift_JIS / UTF-16）・区切り文字・改行コードを保持し、未編集の行は元のテキストのまま書き戻す。Shift_JIS のエンコードは依存ライブラリを増やさず `TextDecoder` の逆引き表で行う。

**Phase 2 で割り切ったこと**: 数式の再計算はしない（入力した数式は Excel で開くまで値が空で、グリッドには数式を斜体で表示）。図形・画像のアンカーは行列操作で移動しない。セル書式（太字・色など）の変更、行高・列幅の変更、シートの並べ替え・非表示化、xls は未対応。

### Phase 3 の詳細（Word / PowerPoint のテキスト編集）

Phase 2 と同じ方針（最小差分保存、正の状態は Extension Host、未対応の要素は壊さず拒否）で、Excel の仕組みを共通化して使う。

| 項目 | 内容 |
|---|---|
| 共通化 | `EditSession<Op>`（`handle` / `edit` / `save` / `reset`）を全形式の共通インターフェースにし、拡張側は `OfficeEditorProvider`（旧 `SheetEditorProvider`）1 つで xlsx / csv / docx / pptx を扱う。Word / PowerPoint は `DocumentSession`（操作キューで直列化、`generation` で reset 前の Undo を無効化、編集のたびにビュー全体を再送）を継承。PDF のみ読み取り専用 |
| Word の操作（`DocxEditOp`） | `setParagraphText(id, text)`、`insertParagraph(after, text, update?)`（`update` 付きで段落の分割）、`deleteParagraph(id)`。id は本文（または表セル）内の要素の位置を `/` でつないだもの（表は `表/行/列/セル内の位置`） |
| Word の書き換え | 段落の `w:pPr` は残し、run を 1 つにまとめる（最初のテキスト run の `w:rPr`、run がなければ段落記号の `w:rPr`）。タブは `w:tab`、改行は `w:br`。段落の両端にあるブックマークは残し、`proofErr` は捨てる。挿入する段落は `w:pPr` をコピー（`sectPr`・`pPrChange` は除く）し、`w14:paraId` は付けない。分割した段落がセクション末尾なら `sectPr` は後ろの段落へ移す |
| Word で拒否するもの | フィールド（`fldChar` / `fldSimple` / `instrText`）、ハイパーリンク、画像・図形・OLE、変更履歴（`ins` / `del` / `move*` / `*PrChange`）、コメント範囲・参照、コンテンツコントロール、段落途中のブックマーク、改ページ・段区切り、脚注参照、ルビ、数式、カスタム XML・スマートタグ、編集権限範囲。削除は加えて、ブックマークを含む段落、セクション区切りの段落、文書・表セルの最後の段落、表と表の間の唯一の段落（Word で表が結合されるため） |
| Word の UI | docx-preview の DOM と XML の段落の対応付けは不確実なため、ページ表示の横に「編集パネル」（段落と表セルの一覧、各段落が textarea）を置く。フォーカスが外れたら確定、Enter で分割、空段落の Backspace で削除。編集後はホストが段落一覧（`docxOutline`）だけを返し、ページ表示は編集が 0.6 秒途切れてから Webview が要求（`preview`）して描き直す（裏で描画してから差し替え、スクロール位置は維持） |
| Word の段落 id のずれ | id は位置なので、挿入・削除の応答（ホストが送り元のパネルだけに返す `edited`）を待つ間の操作は保留し、応答後に id をずらして送る（削除された段落への操作は捨てる）。分割と削除はパネルにすぐ反映し、Enter / Backspace の連打が画面どおりに働くようにする |
| PowerPoint の操作（`PptxEditOp`） | `setShapeText(slide, shape, paragraphs)`、`setTableCellText(slide, shape, row, col, paragraphs)`、`duplicateSlide` / `deleteSlide` / `moveSlide`。shape は `cNvPr@id`。段落内の `\n` は `a:br`。テキスト編集はスライドの `sldId@id`（`slideId`）も持ち、入力中に別の操作でスライドが動いても編集を始めたスライドに書く。スライド番号フィールドは実際の番号として比較する |
| PowerPoint の書き換え | 旧段落と新段落を最長共通部分列で対応付け、テキストが同じ段落は XML ごとそのまま残す。変わった段落は `a:pPr` と `a:endParaRPr` を残して最初の run の `a:rPr` で書き直し、追加した段落は直前の段落の `a:pPr` / `a:endParaRPr` / `a:rPr` をコピー。`txBody` のないプレースホルダー・図形には作成（図形は PowerPoint と同じく中央揃え）。フィールド（`a:fld`）を含む段落の書き換え、レイアウト・マスターの図形、`mc:AlternateContent` 内の図形、グラフ・画像などのプレースホルダー、結合で隠れたセルは拒否 |
| スライド操作 | `presentation.xml` の `sldIdLst`、`presentation.xml.rels`、`[Content_Types].xml`、セクション（`p14:sectionLst`）を更新。複製はスライドとノートをコピーして相互の rels を付け替え、`p14:creationId` を振り直す（コメントは複製しない。グラフ・SmartArt・埋め込み・タグを含むスライドは拒否）。削除はスライドとそのノート・コメントのパートを削除（他スライドからのリンク、目的別スライドショー、最後の 1 枚は拒否）。移動はセクションも移動先に合わせる |
| PowerPoint の UI | 図形・表セルをダブルクリックすると、描画済みの要素をその場で contenteditable にする（Enter = 段落、Shift+Enter = 改行、Ctrl+Enter / 外側クリック = 確定、Esc = 取り消し）。編集後は画像を除いたモデルを再送し、Webview は手元の画像を使い回す。サムネイルの右クリック・ドラッグでスライド操作 |
| Undo | テキスト編集は変更前の要素（Word は段落、PowerPoint は `txBody`）のクローンを位置から引き直して戻す。スライド操作はパッケージのジャーナル（`begin` / `commit` / `restore`）で戻す |
| IME と Undo / 保存キー | 入力は textarea / contenteditable にそのまま任せる（`isComposing` 中の Enter などは処理しない）。入力途中（未確定の変更あり）の間は Ctrl+Z / Ctrl+Y を Webview 内で止め、ブラウザ標準の入力の取り消しにする。変更がなければ VS Code に渡り、文書の Undo になる。入力途中の Ctrl+S は入力を確定する編集を送ってから `save` メッセージを送り、ホストがその編集の適用後に保存する |
| AI 用ツール | `officeWorkbench_editDocument`（段落番号で `setParagraph` / `insertParagraph` / `deleteParagraph`、`replaceText`）と `officeWorkbench_editPresentation`（`setText` / `setTableCell` / `replaceText` / スライド操作）。番号・図形 ID は `officeWorkbench_readFile` の `outline: true` で得る。適用はエディタと同じ `applyEdit` |

**共通の修正**: JSZip は保存時に指定と違う圧縮方式のエントリを再圧縮するため、無圧縮（STORE）で格納されたエントリ（Word は画像をこう格納する）が変わっていた。読み込み時に元の方式を指定しておき、未変更エントリのバイト一致を保つ（Excel にも効く）。

**検証**: core のテスト（各操作、拒否されるケース、Undo / Redo とリセット、未変更パートのバイト一致、Markdown 変換への反映、ツール用の変換）。ブラウザのプレビュー（Playwright）で、ダブルクリック編集・IME（CDP の `Input.imeSetComposition` / `insertText`）・入力途中の Ctrl+Z・確定・取り消し・表セル・空のプレースホルダー・スライドの複製 / 削除 / ドラッグ移動・Undo / Redo・保存、Word の編集パネル（分割・削除・拒否時の通知）を確認。編集したファイルは LibreOffice・python-docx / python-pptx で開けること、パッケージの整合性（Content_Types と rels の参照先）を確認した。Office（COM）と VS Code 本体での確認用に `npm run edited-samples` と `scripts/check-office.ps1` を用意した。

**Phase 3 で割り切ったこと**: 文字単位の書式（太字・色など）の編集はできず、編集した段落は最初の run の書式にそろう。Word のヘッダー・フッター・脚注、PowerPoint のノートは編集できない。Word の編集パネルとページ表示の位置の連動はない。図形の移動・サイズ変更、スライドの新規追加（レイアウト選択）は未対応。

### Phase 4 の詳細

**前半（実装済み）: テキスト抽出と Git 差分**

| 項目 | 内容 |
|---|---|
| 変換（core `text/`） | `toMarkdown(kind, data)`。Word は自前の簡易パーサ（スタイル名・アウトラインレベルで見出し、numbering.xml で番号/箇条書き、`w:del` を除外、`vMerge`/`gridSpan` を考慮した表）。Excel は既存の `Workbook` の表示値と数式（縦方向のコピーを範囲にまとめる）。PowerPoint は既存のスライドモデル |
| 差分に強い出力 | Excel の表は値のある行だけを行番号付きで出し、行末の空セルを省く（列が増えても変わるのは見出しと該当行だけ） |
| VS Code 統合 | 仮想ドキュメント `office-text:<path>.md?<元 URI>`（読み取り専用、元ファイルの変更で更新）。差分は組み込み Git 拡張の API（`toGitUri(uri, 'HEAD')`）で HEAD 版を読み、`vscode.diff` で Markdown 同士を比較。エディタタイトル、エクスプローラー、ソース管理ビューの右クリックから実行 |

**後半（実装済み）: Language Model Tools**

| 項目 | 内容 |
|---|---|
| `officeWorkbench_readFile` | 前半の Markdown 変換を返す。Excel はシート・範囲指定、出力長の上限（既定 4 万文字、超過時は範囲指定を促す）。開いているエディタがあればその現在の状態（未保存の変更を含む）を読む |
| `officeWorkbench_editSpreadsheet` | モデル向けの表現（A1 アドレス、シート名、1 始まりの行番号、列の文字）を core の `toEditOp` で `EditOp` に変換し、エディタと同じ `applyEdit` で 1 操作ずつ Undo 可能な編集として適用。開いていなければ表のエディタを背面で開く。途中で失敗したら、そこまでの結果と理由（例: 存在するシート名の一覧）を返す。`prepareInvocation` で確認を出し、`save` 指定時のみ保存 |
| API | 同じ処理を `activate()` の戻り値として公開（他の拡張・テスト用） |

確認は、VS Code の開発ホストで別のテスト用拡張から `vscode.lm.invokeTool`（読み取り）と公開 API（編集）を呼び出して行った。engines は `vscode.lm.registerTool` が安定版になった 1.95 以上に変更。

**今後の候補**: ステージ済み（index）との比較、コマンドライン用の `git diff` textconv スクリプト、PDF のテキスト抽出。

## 9. 決めておきたい論点
1. **想定利用者と優先形式**: Excel 優先で良いか（編集の需要・実装コスト比が最も高いため推奨）。
2. **Web 拡張（vscode.dev）対応を必須とするか**: 必須なら Node 依存を完全に排除する設計を初期から徹底する。
3. **AI 連携（LM Tools）の優先度**: 早期に入れると差別化要素になる。
4. **Office COM 連携の要否**: Windows ローカル限定だが、完全忠実な PDF 化や再計算が可能。
5. **公開範囲**: Marketplace 公開か社内配布か（ライセンス・依存 OSS の選定に影響）。

## 10. 開発運用
- リポジトリは `office-workbench/main` を main ブランチの作業ツリーとし、機能ブランチは `git worktree add ../<branch> -b <branch>` で兄弟ディレクトリに作成する。
