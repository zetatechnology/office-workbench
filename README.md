# Office Workbench

VS Code 上で Word / Excel / PowerPoint（と PDF）ファイルを閲覧・操作するための拡張機能。

- 実装方針: [docs/implementation-plan.md](docs/implementation-plan.md)

## 現在の状態（Phase 1 / 1.5: 閲覧）

`.docx` / `.xlsx` / `.pptx`（およびマクロ付きの `.docm` / `.xlsm` / `.pptm`）と `.pdf` をエディタで開くと、読み取り専用で表示します。

| 形式 | 表示内容 |
|---|---|
| Word | docx-preview によるページ表示、ズーム、幅に合わせる、ページ番号 |
| Excel | シート切替、数値書式・日付、フォント・塗り・罫線、結合セル、列幅・行高、数式バー、キーボード移動、仮想スクロール |
| PowerPoint | スライド一覧、レイアウト/マスター継承、テーマ色、図形・グラデーション・線・矢印、画像、表スタイル、ノート、非表示スライド |
| PDF | PDF.js による表示、ページ移動、ズーム（自動 / 幅 / ページ全体 / 倍率）、文字の検索・選択、目次（しおり）、日本語 CMap 対応 |

グラフ・SmartArt・条件付き書式などは未対応で、プレースホルダーまたは「Not shown」として表示します。

## 開発

```bash
npm install
npm run build        # 拡張とWebviewをバンドル（packages/extension/dist）
npm test             # core のテスト（Vitest）
npm run lint
npm run typecheck
```

- **VS Code で動作確認**: このフォルダを VS Code で開き、F5（Run Extension）。`test-fixtures/` を開いた拡張開発ホストが起動します。
- **ブラウザで Webview を確認**: `npm run build && npm run preview` → http://localhost:5178/ （`-- --dir <フォルダ>` で任意のファイルを表示）
- **VSIX 作成**: `npm run package`
- **テスト用ファイルの再生成**: `npm run fixtures`（Windows + Microsoft Office が必要。`sample.pdf` は Word から書き出し。生成物はコミット済み）

他の PDF 拡張機能と併用している場合、VS Code がどちらで開くかを尋ねます。右クリックの「Reopen Editor With...」で切り替えられます。

## 構成

```
packages/
  core/       OOXML 読み取り（xlsx / pptx）、数値書式、Webview とのメッセージ定義。VS Code 非依存
  extension/  VS Code 拡張本体（CustomReadonlyEditorProvider）
  webview/    Webview UI（Word / Excel / PowerPoint / PDF ビューア）
scripts/      ビルド、プレビュー用サーバー、フィクスチャ生成
test-fixtures/ Office で作成したサンプルファイル
```
