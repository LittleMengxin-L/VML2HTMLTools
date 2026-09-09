# outlook-vml-webview

用于 HarmonyOS ArkWeb 的 Microsoft Outlook 经典版 VML 邮件解析 HAR。

Outlook 经典版导出的 HTML 邮件依赖 IE 条件注释中的 **VML（Vector Markup Language）**
绘制邮件内嵌图形，而 ArkWeb 不解析 VML，导致这些图形无法渲染。本 SDK 在 ArkWeb 的
JavaScript 上下文中注入一个离线解析内核，把 VML `shape` 的坐标/尺寸信息提取并映射为
普通 `<img>` 与绝对定位布局，使邮件在 ArkWeb 中按原版式展示，并保证回复/转发/再次编辑
链路可拿到干净的 HTML。

SDK 采用「ArkTS 适配层 + 注入式 JS 解析内核」架构：ArkWeb 页面从创建、就绪到内核注入的
完整生命周期封闭在 HAR 内部，调用方只需**挂载一个透明容器组件、传入源 HTML 文件路径**，
即可拿到解析后的 HTML 文件路径与结构化解析报告。

## 开始使用

### 环境要求

- DevEco Studio（HarmonyOS 工程，HAR 模块）
- HarmonyOS SDK：本 HAR 声明 `compatibleSdkVersion: 12`（API 12 及以上）；随附演示
  工程基于 API 23（HarmonyOS 6.1.0）构建验证
- Node.js 18+（可选，仅在手动执行运行时生成脚本时需要）

### 构建

1. 克隆/复制本仓库，用 DevEco Studio 打开；
2. 将 `outlook-vml-har` 作为 **Static Library（HAR）** 模块导入工程；
3. 执行 `Build > Make Module 'outlook_vml_webview'`，或使用 Hvigor：

```shell
hvigorw --mode module -p module=outlook_vml_webview@default assembleHar
```

模块的 `hvigorfile.ts` 已注册 `generateVmlRuntime` 自定义任务，会在 ArkTS 编译
（`default@HarCompileArkTS`）**之前**自动把 JS 解析内核（`tools/runtime/vml-runtime.iife.js`）
重新生成为内嵌的 `VmlRuntime.ets` 常量；脚本是幂等的（内容未变化时不重写文件，不影响
Hvigor 增量编译），因此日常构建无需手工步骤。手动执行：

```shell
npm run generate
```

> 注意：`tools/runtime/vml-runtime.iife.js` 只是**构建期输入**，随包分发的运行时是
> `VmlRuntime.ets` 中内嵌的常量副本，iife 本身不会被打进 HAR。

### 安装

构建得到 `.har` 后，在宿主工程的 `oh-package.json5` 中声明本地依赖：

```json5
{
  "dependencies": {
    "outlook-vml-webview": "file:../libs/outlook-vml-webview-1.0.0.har"
  }
}
```

执行：

```shell
ohpm install
```

发布到公共 OHPM 仓库后，即可直接：

```shell
ohpm install outlook-vml-webview
```

### 使用说明

#### 工作原理

一次文件级解析在 HAR 内部完成以下链路：

1. **准备宿主**：透明 `VmlHost` 组件创建 1×1 的 `about:blank` ArkWeb 页面，就绪后由
   HAR 自动注入 JS 解析内核（幂等：已注入且版本一致则跳过，见 `OutlookVmlSdk`）；
2. **读取与解码**：按 BOM → `<meta charset>`（`gb2312`/`gbk` 归一到 `gb18030`）自动识别
   输入 HTML 的字符编码并解码为 Unicode；
3. **离线解析**：内核先完成 `p → div` 结构预处理（避免 `p > span > table` 被 HTML
   内容模型截断），再用 `DOMParser` 离屏解析，按 Extract → Normalize → Plan → Apply
   管线提取 VML `shape`、把坐标映射为普通图片与定位布局，全程不修改页面 DOM、不依赖
   `img/cid` 资源加载完成；
4. **落盘**：结果统一以 UTF-8 字节写出，并同步把 HTML 内 charset 声明改写为 `utf-8`，
   杜绝编码不一致导致的乱码。

#### 调用流程

与任何解析 SDK 相同，调用顺序必须正确：

1. **挂载容器组件**：在页面 `build()` 中挂载 `<VmlHost>`（可放在 `Stack` 任意层）；
2. **准备输入文件**：把待解析的 HTML 写入应用沙箱内可读的绝对路径（公共目录文件需先经
   Picker 授权并拷贝进沙箱）；
3. **执行解析**：调用 `VmlHostController.parseFile(inputPath, outputPath, options?)`；
4. **判定结果**：检查 `report.status` 与 `report.errors`（结构化错误码，见「错误处理」），
   成功则使用返回的 `outputPath` 文件继续展示或编辑。

#### API 参考

##### VmlHost

透明容器组件。内部是一个 1×1、透明、`javaScriptAccess(true)` 的 ArkWeb 页面，负责页面
创建、`onPageEnd` 就绪上报与销毁清理；解析完成后无任何 UI 残留。

**属性：**

| 名称 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `host` | `VmlHostController` | 是 | 与组件绑定的控制器实例，需与调用方持有的是同一个 |

```ts
VmlHost({ host: this.host })
```

##### VmlHostController

托管解析控制器，持有内部页面状态机（就绪等待、15 秒超时、销毁清理与等待队列释放）。

**`parseFile(inputPath, outputPath, options?)`**

| 名称 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `inputPath` | `string` | 是 | 源 HTML 文件路径（沙箱内可读绝对路径） |
| `outputPath` | `string` | 是 | 解析结果落盘路径（父目录须存在且可写） |
| `options` | `VmlParserOptions` | 否 | 解析选项，透传给注入内核 |

**返回值：** `Promise<VmlFileParseResult>`

```ts
interface VmlFileParseResult {
  outputPath: string | null; // 解析或落盘失败时为 null
  report: VmlReport;
}
```

**示例：**

```ts
import { VmlHost, VmlHostController, VmlFileParseResult } from 'outlook-vml-webview';

@Entry
@Component
struct MailPage {
  private host: VmlHostController = new VmlHostController();

  aboutToAppear(): void {
    this.parseHtmlFile();
  }

  private async parseHtmlFile(): Promise<void> {
    const context = this.getUIContext().getHostContext();
    if (!context) {
      return;
    }
    // 输入输出路径均为应用沙箱内可读写绝对路径（例如 filesDir）。
    // 不传 options：解析行为全部采用内核 DEFAULT_OPTIONS 默认值。
    const result: VmlFileParseResult = await this.host.parseFile(
      `${context.filesDir}/mail-input.html`,
      `${context.filesDir}/mail-output.html`
    );
    if (result.report.status === 'failed' || result.outputPath === null) {
      const code: string = result.report.errors.length > 0 ?
        result.report.errors[0].code : 'unknown';
      console.error(`解析失败：${code}`);
      return;
    }
    // result.outputPath 即解析后的 HTML 文件，可直接用于继续展示或回复/转发。
  }

  build() {
    Stack() {
      // 业务页面内容……
      VmlHost({ host: this.host })
    }
  }
}
```

> `notifyPageReady()` 与 `markDestroyed()` 由组件的 `onPageEnd` /
> `aboutToDisappear` 自动回调，调用方无需手动调用。单个页面建议复用同一个
> `VmlHostController`；组件销毁后重新挂载即可继续使用。

##### OutlookVmlSdk.parseVMLImagesFromHtml(html, options?)

字符串入口的离线转换：输入完整 HTML，在 ArkWeb 的离屏 `Document` 中解析 VML，输出完整
HTML。`VmlHostController.parseFile` 内部即此入口的文件级封装；自带 Web 页面、希望直接
处理内存字符串的场景可自行调用。

**前置条件：** 传入的 `WebviewController` 必须已与页面中的 `Web` 组件绑定、页面已完成
加载，且 Web 启用了 `javaScriptAccess(true)`。

| 名称 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `html` | `string` | 是 | 待解析的完整 HTML 字符串 |
| `options` | `VmlParserOptions` | 否 | 解析选项，透传给注入内核 |

**返回值：** `Promise<ParseHtmlResult>`

```ts
interface ParseHtmlResult {
  html: string | null; // 转换失败时为 null
  report: VmlReport;
}
```

**示例：**

```ts
import { webview } from '@kit.ArkWeb';
import { OutlookVmlSdk, ParseHtmlResult } from 'outlook-vml-webview';

const controller: webview.WebviewController = new webview.WebviewController();
const sdk: OutlookVmlSdk = new OutlookVmlSdk(controller);

// 页面 onPageEnd 后调用。
const result: ParseHtmlResult = await sdk.parseVMLImagesFromHtml(html);
if (result.report.status !== 'failed' && result.html !== null) {
  // result.html 为解析后的完整 HTML（含 <!DOCTYPE>）。
}
```

##### VmlParserOptions

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `processAllWrapTypes` | `boolean` | `false` | 默认仅处理携带 wrap 环绕布局信息的 VML `shape`；置 `true` 后对无 wrap 信息的 `shape` 一并处理 |
| `removeEmptyPlaceholderTables` | `boolean` | `true` | 转换后移除 Outlook 为 VML 预留的空占位表格 |
| `skipDirectDivImages` | `boolean` | `true` | 直接位于 `<div>` 下的图片跳过定位，保持兼容的环绕图片结构 |
| `markAttribute` | `string` | `data-vml-processed` | 已处理节点的标记属性名（防止重复处理） |
| `maxCoordinatePt` | `number` | `10000000` | 坐标安全上限，超出视为异常布局，保留原图并跳过定位 |
| `debug` | `boolean` | `false` | 输出调试诊断信息 |

##### VmlReport

所有解析结果统一返回结构化报告，业务只应依赖 `status` 与各 issue 的 `code`：

```ts
interface VmlReport {
  status: string;      // 'ok' | 'partial' | 'failed'
  stats: VmlStats;     // shapes/images/matched/transformed/skipped/…
  issues: VmlIssue[];  // 全部问题
  errors: VmlIssue[];  // 错误（按 severity 归类）
  warnings: VmlIssue[];
  infos: VmlIssue[];
}

interface VmlIssue {
  code: string;        // 稳定错误码（见「错误处理」），唯一可依赖字段
  severity: string;    // 'error' | 'warning' | 'info'
  stage: string;
  message: string;
  cause: string;
  action: string;
  shapeIds?: string[];
  originalError?: string;
}
```

#### 完整使用示例

随附的 `entry` 演示页（`pages/Index.ets`）提供可运行的全流程：读取 rawfile 样例邮件 →
字节落盘为输入文件 → `parseFile` 解析 → 在展示 Web 中渲染解析结果并校验正文非空。
最小完整用法如下：

```ts
import { fileIo } from '@kit.CoreFileKit';
import { VmlFileParseResult, VmlHost, VmlHostController } from 'outlook-vml-webview';

@Entry
@Component
struct MailPage {
  private host: VmlHostController = new VmlHostController();

  aboutToAppear(): void {
    this.parseMailFile();
  }

  // Step 1: 输入 HTML 落到沙箱（示例从 rawfile 拷贝；注意保留原始字节，
  //        编码探测统一由 HAR 完成，不要在此按 UTF-8 解码转写）。
  private writeBytes(path: string, bytes: Uint8Array): boolean {
    try {
      const file: fileIo.File = fileIo.openSync(path,
        fileIo.OpenMode.WRITE_ONLY | fileIo.OpenMode.CREATE | fileIo.OpenMode.TRUNC);
      fileIo.writeSync(file.fd, bytes.buffer);
      fileIo.closeSync(file);
      return true;
    } catch (error) {
      console.error(`写入失败：${String(error)}`);
      return false;
    }
  }

  private async parseMailFile(): Promise<void> {
    const context = this.getUIContext().getHostContext();
    if (!context) {
      return;
    }
    const inputPath: string = `${context.filesDir}/mail-input.html`;
    const outputPath: string = `${context.filesDir}/mail-output.html`;
    const bytes: Uint8Array =
      await context.resourceManager.getRawFileContent('mail/outlook.html');
    if (!this.writeBytes(inputPath, bytes)) {
      return;
    }

    // Step 2: 执行托管解析（内部等待容器就绪 → 解码 → 注入内核 → 离线解析）。
    // 不传 options：解析行为全部采用内核 DEFAULT_OPTIONS 默认值。
    const result: VmlFileParseResult = await this.host.parseFile(inputPath, outputPath);

    // Step 3: 判定结果。
    if (result.report.status === 'failed' || result.outputPath === null) {
      console.error(`解析失败：${JSON.stringify(result.report.errors)}`);
      return;
    }
    // Step 4: outputPath 即解析后的 HTML 文件，可读取后继续渲染或编辑。
    console.info(`解析完成：transformed=${result.report.stats.transformed}`);
  }

  build() {
    Stack() {
      // 业务页面内容……
      VmlHost({ host: this.host })
    }
  }
}
```

### 错误处理

错误码分两层，**业务侧只应依赖 `code`**（`message`/`cause` 文案可能随版本调整）。

**适配层错误码（`HarErrorCode`，H 前缀）：**

| 错误码 | 枚举成员 | 含义 |
|--------|----------|------|
| H7002 | `RUNTIME_INJECTION_FAILED` | Web 运行时注入失败 |
| H7005 | `RUNTIME_VERSION_MISMATCH` | 注入后的运行时版本与 HAR 不一致 |
| H7006 | `INVALID_HTML` | HTML 参数无效（空字符串或空文件路径） |
| H7011 | `OUTPUT_FILE_WRITE_FAILED` | 写入解析结果文件失败 |
| H7012 | `PAGE_READY_TIMEOUT` | 内部 ArkWeb 页面在 15 秒内未就绪 |
| H7013 | `HTML_TRANSFORM_FAILED` | HTML 字符串转换或序列化失败 |
| H7014 | `FILE_READ_FAILED` | 输入 HTML 文件读取失败 |

**运行时错误码（JS 内核，共 45 个）：** 格式为 `<前缀字母><4 位数字>`，前缀区分处理
阶段，例如 `I1001`（输入校验）、`V2001`（VML 提取）、`M3007`（布局规划）、`X9001`
（未归类异常）。脚本执行异常时由适配层统一以 `H7013` 返回。

**编码降级不报错：** 输入文件声明了无法识别的字符编码时，`VmlText` 自动按 UTF-8 降级
解码并输出 hilog 警告，不影响解析流程。

## 平台支持

- HarmonyOS ArkWeb（支持 `WebviewController.runJavaScript` 与 `onPageEnd`）
- 兼容性：HAR 声明 `compatibleSdkVersion: 12`；演示工程基于 API 23（HarmonyOS 6.1.0）
- 输入邮件编码：UTF-8 / UTF-16（BOM）与 `gb2312`/`gbk`（经典版 Outlook 中文邮件，
  按 `gb18030` 解码）、`big5` 等常见 `charset` 声明

## 依赖

- **运行时**：仅使用 HarmonyOS SDK 自带 Kit（`@kit.ArkWeb`、`@kit.CoreFileKit`、
  `@kit.ArkTS`），**无任何三方依赖**；
- **构建期**：Node.js（`tools/generate-runtime.mjs`，经 hvigor 自定义任务自动执行）；
- **解析内核**：`tools/runtime/vml-runtime.iife.js` 为纯 JavaScript，构建时嵌入
  `VmlRuntime.ets`，不随 HAR 以资源形式分发。

## 发布前配置

当前 `oh-package.json5` 中的包名、`author` 与 `license: "UNLICENSED"` 是安全占位配置。
发布到公共 OHPM 前必须：

1. 换成实际拥有的包名或组织 scope，并确认 OHPM 上未被占用；
2. 明确许可证，同步修改根目录 `LICENSE` 与 `oh-package.json5`；
3. 使用目标 DevEco Studio / HarmonyOS SDK 执行 release HAR 构建（勿用手工打包的源码包
   代替正式产物）；
4. 确认 README、CHANGELOG、LICENSE 均非空。

## 许可证

Copyright (c) VML SDK Maintainers.

当前以 `UNLICENSED` 占位，未授权任何使用、修改或分发。正式开源前请替换为具体许可证
（如 MIT / Apache-2.0）并同步更新 `oh-package.json5` 与 `LICENSE` 文件。
