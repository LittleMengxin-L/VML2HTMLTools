# 完整示例工程

`VML2HTMLTools-example.zip` 包含 entry 示例、HAR 模块、资源、测试及项目构建配置。

1. 解压到独立目录，使用 DevEco Studio 打开 `VML2HTMLTools-example`。
2. 安装工程依赖（`ohpm install --all`），使用匹配工程版本的 HarmonyOS SDK。
3. 在本机配置自己的应用签名。示例不提供证书、密钥或签名密码。
4. 将待解析的 HTML 放入应用沙箱的 `files/mail-input.html`。
5. 启动 entry 页面，解析结果写入 `files/mail-output.html` 并在页面展示。

压缩包不含 Git 历史、IDE 私有配置、签名材料、依赖缓存、构建产物、本地截图及审核文件。
HAR 模块内不重复包含 example 目录，避免嵌套打包。

更新示例源码后，在工程根目录执行：

```powershell
powershell -ExecutionPolicy Bypass -File outlook-vml-har/tools/pack-example.ps1
```
