import * as vscode from 'vscode';
import {
  decodeText, describeEdit, EditError, MAX_FILE_BYTES, textKindOf, toEditOp, toMarkdown, type HostMessage, type ToolEdit,
} from '@office-workbench/core';
import type { SheetDocuments } from './sheet-editor';

export interface ReadInput {
  filePath: string;
  sheet?: string;
  range?: string;
  maxCharacters?: number;
}

export interface EditInput {
  filePath: string;
  edits: ToolEdit[];
  save?: boolean;
}

/** The same operations as the language model tools, for other extensions (and tests). */
export interface OfficeWorkbenchApi {
  readFile(input: ReadInput): Promise<string>;
  editSpreadsheet(input: EditInput): Promise<string>;
}

const DEFAULT_MAX_CHARACTERS = 40_000;
const DELIMITED = /\.(csv|tsv|tab)$/i;
const WORKBOOK = /\.(xlsx|xlsm)$/i;

/**
 * Language model tools (Copilot agent mode and other chat participants): read Office/CSV files as
 * Markdown, and edit spreadsheets through the same path as the editor, so every edit is an undo
 * step in an open editor and nothing is written to disk unless asked.
 */
export function registerLanguageModelTools(context: vscode.ExtensionContext, registry: SheetDocuments): OfficeWorkbenchApi {
  const api: OfficeWorkbenchApi = {
    readFile: (input) => readOfficeFile(input, registry),
    editSpreadsheet: (input) => editSpreadsheet(input, registry),
  };
  context.subscriptions.push(
    vscode.lm.registerTool<ReadInput>('officeWorkbench_readFile', {
      prepareInvocation: ({ input }) => ({ invocationMessage: `Reading ${baseName(input.filePath)}` }),
      invoke: async ({ input }) => result(await api.readFile(input)),
    }),
    vscode.lm.registerTool<EditInput>('officeWorkbench_editSpreadsheet', {
      prepareInvocation({ input }) {
        const lines = (input.edits ?? []).map((e) => `- ${describeEdit(e)}`).join('\n');
        const after = input.save ? 'The file will be saved.' : 'Changes stay unsaved in the editor until you save them.';
        return {
          invocationMessage: `Editing ${baseName(input.filePath)}`,
          confirmationMessages: { title: `Edit ${baseName(input.filePath)}`, message: new vscode.MarkdownString(`${lines}\n\n${after}`) },
        };
      },
      invoke: async ({ input }) => result(await api.editSpreadsheet(input)),
    }),
  );
  return api;
}

async function readOfficeFile(input: ReadInput, registry: SheetDocuments): Promise<string> {
  const uri = resolvePath(input.filePath);
  const open = registry.get(uri);
  // An open editor may hold unsaved edits; read its current state rather than the file.
  const data = open ? await open.document.session.save() : await readFile(uri);
  const note = open && open.document.unsavedEdits !== 0 ? 'Note: includes unsaved changes from the open editor.\n\n' : '';
  let text: string;
  if (DELIMITED.test(uri.path)) {
    text = decodeText(data).text;
  } else {
    const kind = textKindOf(uri.path);
    if (!kind) throw new Error('Supported files are .docx, .xlsx, .pptx (and macro-enabled variants), .csv and .tsv.');
    text = await toMarkdown(kind, data, { sheet: input.sheet, range: input.range });
  }
  const max = input.maxCharacters ?? DEFAULT_MAX_CHARACTERS;
  if (text.length > max) {
    text = `${text.slice(0, max)}\n\n[Truncated: ${text.length} characters in total. Read a specific sheet/range or raise maxCharacters.]`;
  }
  return note + text;
}

async function editSpreadsheet(input: EditInput, registry: SheetDocuments): Promise<string> {
  const uri = resolvePath(input.filePath);
  if (!WORKBOOK.test(uri.path) && !DELIMITED.test(uri.path)) throw new Error('Only .xlsx/.xlsm workbooks and .csv/.tsv files can be edited.');
  if (!Array.isArray(input.edits) || !input.edits.length) throw new Error('Provide at least one edit.');
  const { document, provider } = await openSheet(uri, registry);

  const done: string[] = [];
  for (const [i, edit] of input.edits.entries()) {
    try {
      const [info] = await document.session.handle({ type: 'ready' });
      const sheets = (info as Extract<HostMessage, { type: 'workbook' }>).info?.sheets ?? [];
      await provider.applyEdit(document, toEditOp(edit, sheets));
      done.push(`- Done: ${describeEdit(edit)}`);
    } catch (e) {
      const reason = e instanceof EditError || e instanceof Error ? e.message : String(e);
      const after = done.length ? 'Earlier edits were applied and can be undone in the editor.' : 'Nothing was changed.';
      return [...done, `- Failed (edit ${i + 1}): ${describeEdit(edit)}: ${reason}`, '', after].join('\n');
    }
  }
  if (input.save) await vscode.workspace.save(uri);
  const after = input.save ? 'The file was saved.' : 'The changes are open in the editor, unsaved; the user can review, undo (Ctrl+Z) or save (Ctrl+S).';
  return [...done, '', after].join('\n');
}

function result(text: string): vscode.LanguageModelToolResult {
  return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(text)]);
}

/** The open sheet document for a file, opening it in the table editor (in the background) if needed. */
async function openSheet(uri: vscode.Uri, registry: SheetDocuments) {
  let entry = registry.get(uri);
  if (!entry) {
    const viewType = WORKBOOK.test(uri.path) ? 'officeWorkbench.xlsx' : 'officeWorkbench.csv';
    await vscode.commands.executeCommand('vscode.openWith', uri, viewType, { preview: false, preserveFocus: true });
    for (let i = 0; i < 100 && !entry; i++) {
      await new Promise((r) => setTimeout(r, 50));
      entry = registry.get(uri);
    }
  }
  if (!entry) throw new Error(`Could not open ${uri.fsPath} in the spreadsheet editor.`);
  return entry;
}

/** Absolute paths, file URIs, or paths relative to a workspace folder. */
function resolvePath(filePath: string): vscode.Uri {
  if (!filePath) throw new Error('filePath is required.');
  if (/^[a-z][a-z0-9+.-]+:\/\//i.test(filePath)) return vscode.Uri.parse(filePath);
  if (/^([a-zA-Z]:[\\/]|[\\/])/.test(filePath)) return vscode.Uri.file(filePath);
  const folders = vscode.workspace.workspaceFolders ?? [];
  if (!folders.length) throw new Error(`"${filePath}" is relative, but no folder is open. Use an absolute path.`);
  return vscode.Uri.joinPath(folders[0].uri, filePath.replace(/\\/g, '/'));
}

async function readFile(uri: vscode.Uri): Promise<Uint8Array> {
  let data: Uint8Array;
  try {
    data = new Uint8Array(await vscode.workspace.fs.readFile(uri));
  } catch {
    throw new Error(`File not found: ${uri.fsPath}`);
  }
  if (data.byteLength > MAX_FILE_BYTES) throw new Error('The file is too large to read.');
  return data;
}

function baseName(filePath: string | undefined): string {
  return (filePath ?? '').split(/[\\/]/).pop() || 'file';
}
