import { cp, mkdir, readFile, rm } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';

const moduleDirectory = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const distDirectory = join(moduleDirectory, 'dist');
const stagingRoot = join(distDirectory, 'source-har');
const packageDirectory = join(stagingRoot, 'package');

if (!packageDirectory.startsWith(distDirectory + '\\') && !packageDirectory.startsWith(distDirectory + '/')) {
  throw new Error(`拒绝清理非 dist 目录：${packageDirectory}`);
}

const manifest = JSON.parse(await readFile(join(moduleDirectory, 'oh-package.json5'), 'utf8'));
const harPath = join(distDirectory, `${manifest.name}-${manifest.version}-source.har`);
const entries = [
  'Index.ets',
  'oh-package.json5',
  'README.md',
  'CHANGELOG.md',
  'LICENSE',
  'build-profile.json5',
  'hvigorfile.ts',
  'src'
];

await rm(stagingRoot, { recursive: true, force: true });
await mkdir(packageDirectory, { recursive: true });
for (const entry of entries) {
  await cp(join(moduleDirectory, entry), join(packageDirectory, entry), { recursive: true });
}

await mkdir(distDirectory, { recursive: true });
await new Promise((resolvePromise, rejectPromise) => {
  const command = spawn('tar', ['-czf', harPath, '-C', stagingRoot, 'package'], {
    stdio: 'inherit',
    shell: false
  });
  command.on('error', rejectPromise);
  command.on('exit', code => {
    if (code === 0) resolvePromise();
    else rejectPromise(new Error(`tar 退出码：${code}`));
  });
});

console.log(`Created source HAR: ${harPath}`);
