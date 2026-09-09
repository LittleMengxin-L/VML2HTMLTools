import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const toolsDirectory = dirname(fileURLToPath(import.meta.url));
const moduleDirectory = resolve(toolsDirectory, '..');
const etsPath = join(moduleDirectory, 'src', 'main', 'ets', 'VmlRuntime.ets');
// 运行时源文件放在 tools/ 下：它是「构建期输入」，而非打包进 HAR 的运行时 rawfile 资源。
const sourcePath = join(toolsDirectory, 'runtime', 'vml-runtime.iife.js');

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

// 幂等写入：内容未变化时不重写文件，避免破坏 Hvigor 的增量编译缓存。
let existing = null;
try {
  existing = await readFile(etsPath, 'utf8');
} catch {
  existing = null;
}
if (existing === generated) {
  console.log(`Up to date: ${etsPath}`);
} else {
  await writeFile(etsPath, generated, 'utf8');
  console.log(`Generated ${etsPath}`);
}
