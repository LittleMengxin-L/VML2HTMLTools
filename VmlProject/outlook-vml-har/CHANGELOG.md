# Changelog

## 1.0.0

- 首次封装 Outlook 经典版 VML Web 运行时。
- 提供 ArkWeb `changeVMLSrcs` 和 `parseVMLImages` 适配器。
- 提供 ArkTS `VML_P2DIV` / `prepareDisplayHtml` 预处理函数。
- 提供结构化错误报告和 HAR 适配层 H7001～H7006 错误码。
- 默认以 `document.documentElement` 解析完整 HTML，不再要求 `#mail-content`。
- 新增 `prepareDisplayHtmlFile`，支持读取宿主应用 rawfile HTML 文件。
- 增加 H7007（rawfile 读取失败）和 H7008（rawfile 路径无效）。
