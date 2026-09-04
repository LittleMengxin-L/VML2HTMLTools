import { access, readFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const moduleDirectory = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const requiredFiles = [
  'Index.ets',
  'oh-package.json5',
  'hvigorfile.ts',
  'build-profile.json5',
  'README.md',
  'CHANGELOG.md',
  'LICENSE',
  'src/main/module.json5',
  'src/main/ets/OutlookVmlSdk.ets',
  'src/main/ets/VmlParagraph.ets',
  'src/main/ets/VmlRuntime.ets',
  'src/main/ets/VmlTypes.ets'
];

for (const relativePath of requiredFiles) {
  await access(join(moduleDirectory, relativePath));
}

const manifestText = await readFile(join(moduleDirectory, 'oh-package.json5'), 'utf8');
const manifest = JSON.parse(manifestText);
for (const field of ['name', 'version', 'main', 'license', 'compatibleSdkVersion', 'compatibleSdkType']) {
  if (!manifest[field]) throw new Error(`oh-package.json5 缺少 ${field}。`);
}
if (manifest.artifactType !== 'original') {
  throw new Error('当前源码 HAR 的 artifactType 必须为 original。');
}

const runtimeSource = await readFile(
  join(moduleDirectory, 'src/main/resources/rawfile/outlook_vml/vml-runtime.iife.js'),
  'utf8'
);
const generatedRuntime = await readFile(join(moduleDirectory, 'src/main/ets/VmlRuntime.ets'), 'utf8');
const sourceVersion = runtimeSource.match(/const VERSION = '([^']+)'/)?.[1];
if (!sourceVersion || !generatedRuntime.includes(`VML_RUNTIME_VERSION: string = "${sourceVersion}"`)) {
  throw new Error('VmlRuntime.ets 与 vmlparse.refactored.js 版本不一致，请执行 npm run generate。');
}

console.log(`Package validation passed: ${manifest.name}@${manifest.version}`);
