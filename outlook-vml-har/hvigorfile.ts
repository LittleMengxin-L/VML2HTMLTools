/*
 * Copyright (c) 2026 Huawei Technologies Co., Ltd. All rights reserved.
 *
 * SPDX-License-Identifier: MIT
 * Licensed under the MIT License. See LICENSE in the HAR root for details.
 */

import { harTasks } from '@ohos/hvigor-ohos-plugin';
import { execFileSync } from 'node:child_process';
import { join } from 'node:path';

/**
 * 构建前自动生成 VmlRuntime.ets。
 *
 * 从 tools/runtime/vml-runtime.iife.js 生成 ArkTS 注入常量，
 * 在 default@HarCompileArkTS 之前执行，保持源码与构建产物一致。
 *
 * 生成内容未变化时保留现有文件，以支持增量编译。
 */
function generateVmlRuntimePlugin() {
  return {
    pluginId: 'vml_generate_runtime',
    apply(pluginContext) {
      pluginContext.registerTask({
        name: 'generateVmlRuntime',
        run(taskContext) {
          const modulePath = taskContext && taskContext.modulePath ? taskContext.modulePath : __dirname;
          const scriptPath = join(modulePath, 'tools', 'generate-runtime.mjs');
          execFileSync(process.execPath, [scriptPath], { cwd: modulePath, stdio: 'inherit' });
        },
        // 该任务在 HarCompileArkTS 之前执行：generateVmlRuntime -> default@HarCompileArkTS
        postDependencies: ['default@HarCompileArkTS']
      });
    }
  };
}

export default {
  system: harTasks,
  plugins: [generateVmlRuntimePlugin()]
};
