import { harTasks } from '@ohos/hvigor-ohos-plugin';
import { execFileSync } from 'node:child_process';
import { join } from 'node:path';

/**
 * 构建前自动生成 VmlRuntime.ets。
 *
 * VmlRuntime.ets 是 tools/generate-runtime.mjs 从 rawfile 中的 JS 内核
 * vml-runtime.iife.js 序列化生成的注入常量。手工修改会被覆盖，因此注册
 * 到 ArkTS 编译（default@HarCompileArkTS）之前自动执行，保证源码与产物一致。
 *
 * 幂等：脚本内容未变化时不重写文件，不影响 Hvigor 增量编译。
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
