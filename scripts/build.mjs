// Bundles the extension host code and the webview viewers with esbuild.
import * as esbuild from 'esbuild';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outDir = path.join(root, 'packages/extension/dist');
const watch = process.argv.includes('--watch');
const production = process.argv.includes('--production');

/** @type {esbuild.BuildOptions[]} */
const configs = [
  {
    entryPoints: [path.join(root, 'packages/extension/src/extension.ts')],
    outfile: path.join(outDir, 'extension.js'),
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node20',
    external: ['vscode'],
  },
  {
    entryPoints: {
      docx: path.join(root, 'packages/webview/src/docx/main.ts'),
      xlsx: path.join(root, 'packages/webview/src/xlsx/main.ts'),
      pptx: path.join(root, 'packages/webview/src/pptx/main.ts'),
    },
    outdir: path.join(outDir, 'webview'),
    bundle: true,
    platform: 'browser',
    format: 'iife',
    target: 'es2022',
    loader: { '.css': 'css' },
  },
];

for (const config of configs) {
  const options = { ...config, sourcemap: !production, minify: production, logLevel: 'info' };
  if (watch) {
    const ctx = await esbuild.context(options);
    await ctx.watch();
  } else {
    await esbuild.build(options);
  }
}
