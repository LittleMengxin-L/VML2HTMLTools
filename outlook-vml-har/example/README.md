# 完整示例工程

[VML2HTMLTools.zip](VML2HTMLTools.zip) 提供示例工程，用于演示 HTML 文件选择、VML 解析、页面展示和结果导出。

## 打开与运行

1. 将压缩包解压到独立目录。
2. 使用 DevEco Studio 打开包含 `build-profile.json5`、`entry` 和 `outlook-vml-har` 的工程根目录。
3. 安装与工程配置匹配的 HarmonyOS SDK，在工程根目录执行 `ohpm install --all` 安装依赖。
4. 在本机配置自己的应用签名，选择设备或模拟器运行 entry 模块。

## 选择输入与输出文件

1. 启动应用，页面就绪后自动打开系统文件选择器。
2. 从 Download 等系统选择器可访问的目录中选择一个 `.html` 或 `.htm` 文件。
3. 应用将所选文件导入自身沙箱，调用 HAR 解析并展示结果。
4. 解析完成后自动打开系统保存选择器，自行选择输出目录及文件名；默认文件名为 `mail-output-时间戳.html`。
5. 保存成功后，页面显示结果文件的 URI。

“选择 HTML”按钮可重新选择输入文件；“另存为”按钮可再次导出已有结果。取消保存不会丢弃当前沙箱内的解析结果。导出失败时，可重新选择可写位置。

## 文件访问方式

公共目录中的文件通过系统选择器授权，应用直接使用返回的 URI 读写，无需向应用私有目录手动粘贴文件。

应用内部使用以下中转路径：

| 路径 | 用途 |
| --- | --- |
| `context.filesDir/mail-input.html` | 导入的输入文件，供 HAR 读取 |
| `context.filesDir/mail-output.html` | HAR 生成的结果，供页面展示和导出 |

这些路径仅用于应用内部处理。用户实际保存的位置由系统保存选择器决定，不要自行拼接 `/data/app/...` 物理路径。

## 工程职责

- `entry/src/main/ets/pages/Index.ets`：文件选择、沙箱中转、可见 WebView 展示及结果导出。
- `outlook-vml-har/Index.ets`：SDK 公共接口导出。
- `VmlHostController.parseFile(inputPath, outputPath)`：读取沙箱输入文件，解析后写入沙箱输出文件。
- `VmlHost`：提供隐藏的 ArkWeb 解析环境，解析期间应保持挂载。

文件选择和导出由 entry 实现，HAR 不主动打开文件管理器。详细 API 和模块结构见 [HAR README](../README.md)。

## 示例维护

压缩包是打包时的源码快照，修改仓库中的 entry 不会自动更新压缩包。更新分发包时，应重新打包完整工程，并排除 Git 历史、IDE 私有配置、签名材料、依赖缓存、构建产物及 example 目录本身，避免嵌套打包。
