import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const toolsDirectory = dirname(fileURLToPath(import.meta.url));
const moduleDirectory = resolve(toolsDirectory, '..');
const etsPath = join(moduleDirectory, 'src', 'main', 'ets', 'VmlRuntime.ets');
const rawPath = join(
  moduleDirectory,
  'src',
  'main',
  'resources',
  'rawfile',
  'outlook_vml',
  'vml-runtime.iife.js'
);
const sourcePath = rawPath;

const source = await readFile(sourcePath, 'utf8');
const version = source.match(/const VERSION = '([^']+)'/)?.[1];
if (!version) {
  throw new Error(`无法从 ${sourcePath} 读取 VERSION。`);
}
if (!source.includes('global.OutlookVMLParser = api;')) {
  throw new Error('VML 源码未导出 global.OutlookVMLParser。');
}

const generated = `// 此文件由 tools/generate-runtime.mjs 自动生成，请勿手工编辑。\n` +
  `export const VML_RUNTIME_VERSION: string = ${JSON.stringify(version)};\n` +
  `export const VML_RUNTIME_SOURCE: string = ${JSON.stringify(source)};\n`;

await mkdir(dirname(etsPath), { recursive: true });
await writeFile(etsPath, generated, 'utf8');

console.log(`Generated ${etsPath}`);
