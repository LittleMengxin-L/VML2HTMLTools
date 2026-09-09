# Changelog

## Next

- 运行时选项收敛：删除内核 `DEFAULT_OPTIONS` 中的 `convertVmlParagraphs` 选项（原仅
  `transform` 组合入口消费，而展示链路已统一经 `VML_P2DIV` 无条件完成 `p → div`
  预处理）。`transform` 现于解析前无条件执行该预处理，不再提供开关。
- 新增 `VmlHostController.parseFile`：以「HTML 文件路径入参、解析后 HTML 文件路径 +
  结构化报告出参」的文件级托管解析 API，输入输出均为沙箱可读写绝对路径。
- 新增 `VmlHost` 透明容器组件：将 ArkWeb 页面创建、就绪、运行时注入等生命周期
  封闭在 HAR 内部，调用方无需准备 WebviewController 或维护 onPageEnd 状态机。
- 新增 H7014（输入 HTML 文件读取失败）错误码。
- 修复中文乱码：新增共享编码模块 `VmlText.ets`（BOM / `<meta charset>` 探测，
  gb2312/gbk 归一为 gb18030），`VmlHost` 读取输入文件统一经它解码；解析结果落盘前将
  HTML 内 charset 声明归一为 UTF-8，与写盘字节保持一致。`fileIo.readSync` 修正为仅传
  ArrayBuffer。
- `hvigorfile.ts` 注册 `generateVmlRuntime` 自定义任务，构建 HAR 时在 ArkTS 编译前
  自动重新生成 VmlRuntime.ets；`generate-runtime.mjs` 改为幂等写入，不破坏增量编译。
- 移除旧 rawfile 预处理链路：删除 ArkTS 版 `VML_P2DIV` / `prepareDisplayHtml` /
  `prepareDisplayHtmlFile` 及其 `VmlParagraph.ets` 实现、`PrepareDisplayResult` 类型与
  错误码 H7007～H7009。`p → div` 预处理改由注入的 JS 运行时在解析流程内完成，公共 API
  面收敛为 `VmlHost` 托管解析与 `OutlookVmlSdk` 适配器；同时清理未注册的死页面
  `EntryIndex.ets` 与不再被引用的 `vml-runtime-host.html`。
- 移除无消费者的页面 DOM 链路方法：`OutlookVmlSdk.parseVMLImages` / `getDocumentHtml` /
  `changeVMLSrcs`，以及孤儿类型 `ChangeVmlResult` / `HtmlSnapshotResult` 和孤儿错误码
  H7001/H7003/H7004/H7010。`OutlookVmlSdk` 收敛为 `parseVMLImagesFromHtml` 单一入口，
  托管解析链路为唯一官方用法。
- 新增共享报告工厂 `VmlReport.ets`（`createEmptyStats` / `createHostFailure`），消除
  `OutlookVmlSdk` 与 `VmlHost` 之间两份重复实现。
- 包体瘦身：JS 内核源 `vml-runtime.iife.js` 从 rawfile 移至 `tools/runtime/`（构建输入
  不再随 HAR 打包），删除冗余的 rawfile 样例 `mail/outlook.html`；HAR 自此不再携带任何
  rawfile 资源。
- 收紧公共导出面：`VmlIssueSeverity` / `VmlReportStatus` 两个枚举不再从 `Index.ets`
  导出（`VmlReport` 的状态/严重度本就是字符串字段，且 issue 已按 errors/warnings/infos
  预分类，用户判断无需枚举；枚举定义保留供内部实现使用）。`OutlookVmlSdk` 字符串入口
  与 `HarErrorCode` 错误码契约继续保留。
- 按 SDK 开源规范范式重写 README（中文），对齐最新公共 API 面：补充环境要求、构建/
  安装/调用流程、逐函数 API 参考（含参数表、返回值、示例）、完整使用示例、错误码分层
  说明、平台支持与依赖清单。

## 1.0.0

- 首次封装 Outlook 经典版 VML Web 运行时。
- 提供 ArkWeb `changeVMLSrcs` 和 `parseVMLImages` 适配器。
- 提供 ArkTS `VML_P2DIV` / `prepareDisplayHtml` 预处理函数。
- 提供结构化错误报告和 HAR 适配层 H7001～H7006 错误码。
- 默认以 `document.documentElement` 解析完整 HTML，不再要求 `#mail-content`。
- 新增 `prepareDisplayHtmlFile`，支持读取宿主应用 rawfile HTML 文件。
- 增加 H7007（rawfile 读取失败）和 H7008（rawfile 路径无效）。
