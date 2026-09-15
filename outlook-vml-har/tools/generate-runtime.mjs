/*
 * Copyright (c) 2026 Huawei Technologies Co., Ltd. All rights reserved.
 *
 * SPDX-License-Identifier: MIT
 * Licensed under the MIT License. See LICENSE in the HAR root for details.
 */

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const toolsDirectory = dirname(fileURLToPath(import.meta.url));
const moduleDirectory = resolve(toolsDirectory, '..');
const etsPath = join(moduleDirectory, 'src', 'main', 'ets', 'VmlRuntime.ets');
// 构建期读取 JavaScript 内核，生成用于 ArkWeb 注入的 ArkTS 字符串常量。
const sourcePath = join(toolsDirectory, 'runtime', 'vml-runtime.iife.js');

const source = await readFile(sourcePath, 'utf8');
const version = source.match(/const VERSION = '([^']+)'/)?.[1];
if (!version) {
  throw new Error(`无法从 ${sourcePath} 读取 VERSION。`);
}
if (!source.includes('global.OutlookVMLParser = api;')) {
  throw new Error('VML 源码未导出 global.OutlookVMLParser。');
}

const licenseHeader = source.slice(0, source.indexOf('*/') + 2) + '\n\n';
const generated = licenseHeader + `// 此文件由 tools/generate-runtime.mjs 自动生成，请勿手工编辑。\n` +
  `export const VML_RUNTIME_VERSION: string = ${JSON.stringify(version)};\n` +
  `export const VML_RUNTIME_SOURCE: string = ${JSON.stringify(source)};\n`;

await mkdir(dirname(etsPath), { recursive: true });

// 仅在生成内容变化时写入，保留增量编译缓存。
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
