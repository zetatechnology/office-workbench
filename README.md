# Office Workbench

VS Code 上で Word / Excel / PowerPoint（と PDF）ファイルを閲覧・操作するための拡張機能。

- 実装方針: [docs/implementation-plan.md](docs/implementation-plan.md)

## 現在の状態（Phase 1 / 1.5: 閲覧、Phase 2: Excel・CSV 編集、Phase 3: Word・PowerPoint のテキスト編集）

`.docx` / `.xlsx` / `.pptx`（およびマクロ付きの `.docm` / `.xlsm` / `.pptm`）と `.pdf` をエディタで開いて表示します。Excel・CSV はセル、Word は段落、PowerPoint は図形・表のテキストとスライドを編集・保存できます（PDF は読み取り専用）。

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

### Word の編集

ツールバーの「Edit」で、ページ表示の右に編集パネル（本文と表セルの段落の一覧）を開きます。段落を編集してフォーカスを外すと確定し、ページ表示が更新されます。

| 操作 | 方法 |
|---|---|
| 段落の編集 | 段落をクリックして入力（日本語 IME 可）。別の段落やページ表示をクリックすると確定、Esc で入力を取り消し。Shift+Enter で段落内の改行、Tab でタブ文字 |
| 段落の追加・分割 | Enter。カーソル位置で段落を分け、後ろの部分を新しい段落にする（段落の書式・スタイル・箇条書きを引き継ぐ） |
| 段落の削除 | 空の段落で Backspace |
| 移動 | 段落の先頭で ↑ / 末尾で ↓ で前後の段落へ |
| 元に戻す・保存 | VS Code の Ctrl+Z / Ctrl+Y / Ctrl+S（入力途中の段落では Ctrl+Z はその入力を戻す） |

- 段落の書式（スタイル、インデント、箇条書き、配置）はそのまま残し、新しいテキストには段落の最初の文字の書式（フォント・太字・色など）を使います。段落の一部だけ太字、といった文字単位の書式は、その段落を編集すると最初の文字の書式にそろいます
- 次を含む段落は壊さないよう編集できません（パネルに理由を表示）: フィールド（ページ番号・目次など）、ハイパーリンク、画像・図形、変更履歴、コメント、コンテンツコントロール、段落の途中にあるブックマーク、改ページ、脚注の参照、ルビ、数式
- 文書の最後の段落、表セルの最後の段落、セクション区切りを持つ段落は削除できません

### PowerPoint の編集

| 操作 | 方法 |
|---|---|
| テキストの編集 | スライド上の図形・テキストボックス・表のセルをダブルクリックしてその場で入力（日本語 IME 可）。空のプレースホルダーや文字のない図形にも入力できる。Enter で段落、Shift+Enter で段落内の改行、Ctrl+Enter または外側のクリックで確定、Esc で取り消し |
| スライドの操作 | サムネイルの右クリックで複製・削除・上下へ移動。サムネイルのドラッグで並べ替え |
| 元に戻す・保存 | VS Code の Ctrl+Z / Ctrl+Y / Ctrl+S（入力途中の Ctrl+Z はその入力を戻す） |

- 段落の書式（箇条書き、レベル、配置）と段落末の書式は残し、テキストを変えた段落には最初の文字の書式を使います。テキストが変わらない段落はそのまま残します
- 編集できるのはスライド自身の図形です（レイアウト・マスターの図形、グラフや画像のプレースホルダーは対象外）。スライド番号・日付のフィールドを含む段落は書き換えません
- スライドの複製はノートも複製します。グラフ・SmartArt・埋め込みオブジェクトを含むスライドの複製、他のスライドからリンクされているスライドや目的別スライドショーで使われているスライドの削除はできません

### CSV / TSV

`.csv` / `.tsv` / `.tab` は Excel と同じ表の画面で表示・編集できます。CSV はテキストとして編集することも多いため、通常どおりテキストエディタで開き、エディタ右上の「Open as Table」（表アイコン）またはエクスプローラーの右クリックで表に切り替えます（表からは「Open as Text」で戻ります）。常に表で開きたい場合は設定 `workbench.editorAssociations` に `"*.csv": "officeWorkbench.csv"` を追加します。

- 文字コードは UTF-8（BOM あり / なし）、Shift_JIS（Windows-31J）、UTF-16 を自動判別し、保存時も同じ文字コード・区切り文字・改行コードで書き戻します。Shift_JIS で表せない文字（絵文字など）を含むと保存できません
- 編集した行だけを書き直し、それ以外の行は元のテキスト（引用符の付け方も含む）のまま残します
- 値は文字列としてそのまま保存します（`=1+2` も数式ではなく文字列）。数値に見える値は右寄せで表示します

保存は「変更したパートだけを書き換える」方式で、それ以外の ZIP エントリ（グラフ・画像・マクロ等）は元のバイト列のまま残します（Word / PowerPoint も同じ）。数式は再計算しないため、入力した数式の結果は Excel で開いたときに計算されます（保存時に `fullCalcOnLoad` を設定し、`calcChain.xml` を削除）。

### テキスト表示と Git 差分（Phase 4）

Word / Excel / PowerPoint ファイルを Markdown テキストとして表示し、Git の変更を行単位で比較できます。

| コマンド | 場所 | 内容 |
|---|---|---|
| Open as Markdown | エディタ右上（Markdown アイコン）、エクスプローラーの右クリック | 読み取り専用の Markdown として開く。元ファイルが変わると自動で更新 |
| Compare with HEAD as Text | エディタ右上（比較アイコン）、エクスプローラー / ソース管理ビューの右クリック | HEAD と作業ツリーの内容を Markdown 同士の差分で表示 |

- **Word**: 見出し（`#`）、箇条書き・番号付きリスト、表、ヘッダー / フッター、脚注。変更履歴は削除部分を除いた現在の文面
- **Excel**: シートごとの表（表示値、1 列目は行番号。値のある行だけ）と数式一覧（下方向にコピーされた数式は `E4:E7` のように範囲でまとめる）、結合セル。非表示シートも含む
- **PowerPoint**: スライドごとのテキスト・箇条書き・表、ノート、非表示スライドの印

### AI（言語モデル）ツール（Phase 4）

Copilot のエージェントモードなど、VS Code のチャットから使えるツールを提供します（VS Code 1.95 以降）。チャットで `#readOffice` / `#editSpreadsheet` / `#editDocument` / `#editPresentation` と指定することもできます。

| ツール | 内容 |
|---|---|
| `officeWorkbench_readFile`（Read Office file） | Word / Excel / PowerPoint / CSV を Markdown で読む。Excel はシート名・範囲（`A1:F50`）を指定可能。Word / PowerPoint は `outline: true` で編集ツール用の一覧（段落番号、スライドごとの図形 ID と表のセル）を返す。エディタで開いていれば未保存の変更も含む |
| `officeWorkbench_editSpreadsheet`（Edit spreadsheet） | Excel / CSV のセル入力・クリア、行列の挿入削除、シートの追加・名前変更・削除 |
| `officeWorkbench_editDocument`（Edit Word document） | Word の段落の書き換え・追加・削除、文字列の置換（`replaceText`） |
| `officeWorkbench_editPresentation`（Edit PowerPoint presentation） | PowerPoint の図形・表セルのテキスト、文字列の置換、スライドの複製・削除・移動 |

編集ツールは実行前に確認が表示されます。編集はファイルを Office Workbench のエディタで開き、手で編集したときと同じ仕組みで適用します。各操作は Ctrl+Z で元に戻せ、`save: true` を指定しない限りファイルは保存されません（エディタ上で確認してから保存）。編集結果は `officeWorkbench_readFile` と Markdown 表示・差分にも反映されます。同じ機能は拡張機能の API（`activate()` の戻り値の `readFile` / `editSpreadsheet` / `editDocument` / `editPresentation`）としても使えます。

## 開発

```bash
npm install
npm run build        # 拡張とWebviewをバンドル（packages/extension/dist）
npm test             # core のテスト（Vitest）
npm run lint
npm run typecheck
```

- **VS Code で動作確認**: このフォルダを VS Code で開き、F5（Run Extension）。`test-fixtures/` を開いた拡張開発ホストが起動します。
- **ブラウザで Webview を確認**: `npm run build && npm run preview` → http://localhost:5178/ （`-- --dir <フォルダ>` で任意のファイルを表示）。Excel / CSV / Word / PowerPoint は編集でき、Ctrl+Z / Ctrl+Y で元に戻す・やり直し、Ctrl+S で一時フォルダ（`%TEMP%/office-workbench-preview`）へ保存します（元ファイルは変更しません）
- **編集結果を Office で確認**: `npm run edited-samples` で fixture に各種編集をしたファイルを `%TEMP%/office-workbench-edited` に作り、`npm run check-office`（Windows + Office）で Word / PowerPoint で修復なしに開けるかと内容を確認します（`-Dir <フォルダ>` で任意のファイルも可）
- **VSIX 作成**: `npm run package`
- **テスト用ファイルの再生成**: `npm run fixtures`（Windows + Microsoft Office が必要。`sample.pdf` は Word から書き出し。生成物はコミット済み）

他の PDF 拡張機能と併用している場合、VS Code がどちらで開くかを尋ねます。右クリックの「Reopen Editor With...」で切り替えられます。

## 構成

```
packages/
  core/       OOXML 読み書き（xlsx・docx・pptx の編集と保存）、CSV / TSV の読み書きと文字コード、数値書式、Webview とのメッセージ定義。VS Code 非依存
  extension/  VS Code 拡張本体（Office・CSV は CustomEditorProvider、PDF は CustomReadonlyEditorProvider）
  webview/    Webview UI（Word / Excel / PowerPoint / PDF ビューア）
scripts/      ビルド、プレビュー用サーバー、フィクスチャ生成
test-fixtures/ Office で作成したサンプルファイル（CSV は Excel から Shift_JIS / UTF-8 / UTF-16 で書き出したもの）
```
