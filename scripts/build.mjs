// Bundles the extension host code and the webview viewers with esbuild.
import * as esbuild from 'esbuild';
import { cp, readdir, readFile, writeFile } from 'node:fs/promises';
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
      pdf: path.join(root, 'packages/webview/src/pdf/main.ts'),
    },
    outdir: path.join(outDir, 'webview'),
    bundle: true,
    platform: 'browser',
    format: 'iife',
    target: 'es2022',
    loader: { '.css': 'css', '.svg': 'dataurl', '.png': 'dataurl', '.gif': 'dataurl' },
  },
  {
    // pdf.js worker as a classic script; the PDF viewer loads it through a blob URL.
    entryPoints: { 'pdf.worker': path.join(root, 'node_modules/pdfjs-dist/build/pdf.worker.mjs') },
    outdir: path.join(outDir, 'webview'),
    bundle: true,
    platform: 'browser',
    format: 'iife',
    target: 'es2022',
  },
];

// Data files pdf.js fetches at runtime: CJK character maps, standard fonts, image decoders, ICC profiles.
const pdfjsDir = path.join(root, 'node_modules/pdfjs-dist');
for (const dir of ['cmaps', 'standard_fonts', 'wasm', 'iccs']) {
  await cp(path.join(pdfjsDir, dir), path.join(outDir, 'webview/pdfjs', dir), { recursive: true });
}

const bundledPackages = new Set();
for (const config of configs) {
  const options = { ...config, sourcemap: !production, minify: production, logLevel: 'info', metafile: !watch };
  if (watch) {
    const ctx = await esbuild.context(options);
    await ctx.watch();
  } else {
    const result = await esbuild.build(options);
    for (const input of Object.keys(result.metafile.inputs)) {
      // Innermost node_modules entry, so nested copies (jszip/node_modules/...) are listed too.
      const match = /^(.*node_modules\/(?:@[^/]+\/)?[^/]+)\//.exec(input);
      if (match) bundledPackages.add(path.resolve(match[1]));
    }
  }
}

if (!watch) {
  await writeThirdPartyNotices([...bundledPackages]);
  // vsce reads the license from the extension folder; the source of truth is the repository root.
  await cp(path.join(root, 'LICENSE'), path.join(root, 'packages/extension/LICENSE'));
}

// Collects the license texts of the npm packages bundled into dist into ThirdPartyNotices.txt.
async function writeThirdPartyNotices(packageDirs) {
  const packages = [];
  for (const dir of packageDirs) {
    const pkg = JSON.parse(await readFile(path.join(dir, 'package.json'), 'utf8'));
    const files = await readdir(dir);
    const licenseFile = files.find((f) => /^(licen[cs]e|copying)(\.|$)/i.test(f));
    let text;
    if (licenseFile) {
      text = await readFile(path.join(dir, licenseFile), 'utf8');
    } else {
      // Some small packages (isarray) only carry the license in their README.
      const readme = files.find((f) => /^readme(\.|$)/i.test(f));
      const section = readme && /^#+\s*licen[cs]e\b[^\n]*\n([\s\S]*?)(?=^#+\s|(?![\s\S]))/im.exec(await readFile(path.join(dir, readme), 'utf8'));
      if (!section) throw new Error(`No license text found for bundled package ${dir}`);
      text = section[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>');
    }
    text = text.replace(/\r\n/g, '\n').trim();
    packages.push({ id: `${pkg.name}@${pkg.version}`, pkg, text });
  }
  packages.sort((a, b) => a.id.localeCompare(b.id));
  const sections = packages.map(({ pkg, text }) => {
    const homepage = pkg.homepage ?? pkg.repository?.url ?? pkg.repository ?? '';
    return `${'-'.repeat(78)}\n${pkg.name} ${pkg.version} (${pkg.license})\n${homepage}\n\n${text}\n`;
  });
  const header =
    'Office Workbench includes the following third-party software.\n\n' +
    'The pdf.js data files in dist/webview/pdfjs (character maps, standard fonts, image\n' +
    'decoders and ICC profiles) carry their own license files in the same folders.\n\n';
  await writeFile(path.join(root, 'packages/extension/ThirdPartyNotices.txt'), header + sections.join('\n'));
  console.log(`ThirdPartyNotices.txt: ${packages.length} packages (${packages.map((p) => p.id).join(', ')})`);
}
