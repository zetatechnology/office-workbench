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
- **補助**: `.xls` / `.csv` は SheetJS で読み取り専用表示。

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
| Excel 補助 | SheetJS CE（xls/csv 読み取り） | npm 公開版は古いため公式 CDN tarball を固定バージョンで利用 |
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
| 0 | 雛形（workspaces / esbuild / CI / テスト基盤 / fixtures） | 空の CustomEditor で 3 形式が開ける |
| 1 | 閲覧（Word: docx-preview、Excel: グリッド、PPT: 基本レンダラ） | 代表的 fixture が実用的に表示できる |
| 2 | Excel 編集（セル・行列・シート）+ Undo/保存 | 編集後に Excel で開いて破損なし、未変更パートがバイト一致 |
| 3 | Word / PPT のテキスト編集 | 同上 |
| 4 | テキスト抽出・Git 差分・LM Tools | エージェントから読み書きできる |
| 5 | 任意アダプタ（LibreOffice 変換、Windows の Office COM） | 設定で有効化した場合のみ動作 |

## 9. 決めておきたい論点
1. **想定利用者と優先形式**: Excel 優先で良いか（編集の需要・実装コスト比が最も高いため推奨）。
2. **Web 拡張（vscode.dev）対応を必須とするか**: 必須なら Node 依存を完全に排除する設計を初期から徹底する。
3. **AI 連携（LM Tools）の優先度**: 早期に入れると差別化要素になる。
4. **Office COM 連携の要否**: Windows ローカル限定だが、完全忠実な PDF 化や再計算が可能。
5. **公開範囲**: Marketplace 公開か社内配布か（ライセンス・依存 OSS の選定に影響）。

## 10. 開発運用
- リポジトリは `office-workbench/main` を main ブランチの作業ツリーとし、機能ブランチは `git worktree add ../<branch> -b <branch>` で兄弟ディレクトリに作成する。
