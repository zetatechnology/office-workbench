import * as vscode from 'vscode';
import {
  decodeText, describeDocumentEdit, describeEdit, describePresentationEdit, EditError, MAX_FILE_BYTES, textKindOf, toDocxOps,
  toEditOp, toMarkdown, toOutline, toPptxOps, type AnyEditOp, type DocumentToolEdit, type HostMessage, type PresentationToolEdit,
  type ToolEdit,
} from '@office-workbench/core';
import type { OfficeDocuments } from './editor';

export interface ReadInput {
  filePath: string;
  sheet?: string;
  range?: string;
  /** Word / PowerPoint: the numbered outline the editing tools refer to, instead of Markdown. */
  outline?: boolean;
  maxCharacters?: number;
}

export interface EditInput<Edit = ToolEdit> {
  filePath: string;
  edits: Edit[];
  save?: boolean;
}

/** The same operations as the language model tools, for other extensions (and tests). */
export interface OfficeWorkbenchApi {
  readFile(input: ReadInput): Promise<string>;
  editSpreadsheet(input: EditInput): Promise<string>;
  editDocument(input: EditInput<DocumentToolEdit>): Promise<string>;
  editPresentation(input: EditInput<PresentationToolEdit>): Promise<string>;
}

const DEFAULT_MAX_CHARACTERS = 40_000;
const DELIMITED = /\.(csv|tsv|tab)$/i;
const WORKBOOK = /\.(xlsx|xlsm)$/i;
const DOCUMENT = /\.(docx|docm)$/i;
const PRESENTATION = /\.(pptx|pptm)$/i;

/**
 * Language model tools (Copilot agent mode and other chat participants): read Office/CSV files as
 * Markdown, and edit spreadsheets, documents and presentations through the same path as the
 * editor, so every edit is an undo step in an open editor and nothing is written to disk unless asked.
 */
export function registerLanguageModelTools(context: vscode.ExtensionContext, registry: OfficeDocuments): OfficeWorkbenchApi {
  const api: OfficeWorkbenchApi = {
    readFile: (input) => readOfficeFile(input, registry),
    editSpreadsheet: (input) => editSpreadsheet(input, registry),
    editDocument: (input) =>
      editWith(input, registry, {
        pattern: DOCUMENT,
        kind: 'Word documents (.docx/.docm)',
        viewType: 'officeWorkbench.docx',
        describe: describeDocumentEdit,
        async convert(edit, state) {
          const message = state.find((m) => m.type === 'docx');
          return toDocxOps(edit, (message?.type === 'docx' && message.blocks) || []);
        },
      }),
    editPresentation: (input) =>
      editWith(input, registry, {
        pattern: PRESENTATION,
        kind: 'PowerPoint presentations (.pptx/.pptm)',
        viewType: 'officeWorkbench.pptx',
        describe: describePresentationEdit,
        async convert(edit, state) {
          const message = state.find((m) => m.type === 'presentation');
          if (message?.type !== 'presentation') throw new Error('The presentation could not be read.');
          return toPptxOps(edit, message.presentation);
        },
      }),
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
    vscode.lm.registerTool<EditInput<DocumentToolEdit>>('officeWorkbench_editDocument', {
      prepareInvocation: ({ input }) => confirmation(input, describeDocumentEdit),
      invoke: async ({ input }) => result(await api.editDocument(input)),
    }),
    vscode.lm.registerTool<EditInput<PresentationToolEdit>>('officeWorkbench_editPresentation', {
      prepareInvocation: ({ input }) => confirmation(input, describePresentationEdit),
      invoke: async ({ input }) => result(await api.editPresentation(input)),
    }),
  );
  return api;
}

function confirmation<Edit>(input: EditInput<Edit>, describe: (edit: Edit) => string): vscode.PreparedToolInvocation {
  const lines = (input.edits ?? []).map((e) => `- ${describe(e)}`).join('\n');
  const after = input.save ? 'The file will be saved.' : 'Changes stay unsaved in the editor until you save them.';
  return {
    invocationMessage: `Editing ${baseName(input.filePath)}`,
    confirmationMessages: { title: `Edit ${baseName(input.filePath)}`, message: new vscode.MarkdownString(`${lines}\n\n${after}`) },
  };
}

async function readOfficeFile(input: ReadInput, registry: OfficeDocuments): Promise<string> {
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
    text = input.outline && kind !== 'xlsx' ? await toOutline(kind, data) : await toMarkdown(kind, data, { sheet: input.sheet, range: input.range });
  }
  const max = input.maxCharacters ?? DEFAULT_MAX_CHARACTERS;
  if (text.length > max) {
    text = `${text.slice(0, max)}\n\n[Truncated: ${text.length} characters in total. Read a specific sheet/range or raise maxCharacters.]`;
  }
  return note + text;
}

async function editSpreadsheet(input: EditInput, registry: OfficeDocuments): Promise<string> {
  const uri = resolvePath(input.filePath);
  if (!WORKBOOK.test(uri.path) && !DELIMITED.test(uri.path)) throw new Error('Only .xlsx/.xlsm workbooks and .csv/.tsv files can be edited.');
  if (!Array.isArray(input.edits) || !input.edits.length) throw new Error('Provide at least one edit.');
  const { document, provider } = await openEditor(uri, registry, WORKBOOK.test(uri.path) ? 'officeWorkbench.xlsx' : 'officeWorkbench.csv');

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

interface EditTool<Edit> {
  pattern: RegExp;
  kind: string;
  viewType: string;
  describe(edit: Edit): string;
  /** Operations for one edit, given what the editor shows now (its `ready` messages). */
  convert(edit: Edit, state: HostMessage[]): Promise<AnyEditOp[]>;
}

/** Applies tool edits to a document or presentation, one undo step per operation. */
async function editWith<Edit>(input: EditInput<Edit>, registry: OfficeDocuments, tool: EditTool<Edit>): Promise<string> {
  const uri = resolvePath(input.filePath);
  if (!tool.pattern.test(uri.path)) throw new Error(`Only ${tool.kind} can be edited with this tool.`);
  if (!Array.isArray(input.edits) || !input.edits.length) throw new Error('Provide at least one edit.');
  const { document, provider } = await openEditor(uri, registry, tool.viewType);

  const done: string[] = [];
  for (const [i, edit] of input.edits.entries()) {
    try {
      const ops = await tool.convert(edit, await document.session.handle({ type: 'ready' }));
      for (const op of ops) await provider.applyEdit(document, op);
      done.push(`- Done: ${tool.describe(edit)}${ops.length > 1 ? ` (${ops.length} places)` : ''}`);
    } catch (e) {
      const reason = e instanceof EditError || e instanceof Error ? e.message : String(e);
      const after = done.length ? 'Earlier edits were applied and can be undone in the editor.' : 'Nothing was changed.';
      return [...done, `- Failed (edit ${i + 1}): ${tool.describe(edit)}: ${reason}`, '', after].join('\n');
    }
  }
  if (input.save) await vscode.workspace.save(uri);
  const after = input.save ? 'The file was saved.' : 'The changes are open in the editor, unsaved; the user can review, undo (Ctrl+Z) or save (Ctrl+S).';
  return [...done, '', after].join('\n');
}

/** The open document for a file, opening it in its editor (in the background) if needed. */
async function openEditor(uri: vscode.Uri, registry: OfficeDocuments, viewType: string) {
  let entry = registry.get(uri);
  if (!entry) {
    await vscode.commands.executeCommand('vscode.openWith', uri, viewType, { preview: false, preserveFocus: true });
    for (let i = 0; i < 100 && !entry; i++) {
      await new Promise((r) => setTimeout(r, 50));
      entry = registry.get(uri);
    }
  }
  if (!entry) throw new Error(`Could not open ${uri.fsPath} in the Office Workbench editor.`);
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
