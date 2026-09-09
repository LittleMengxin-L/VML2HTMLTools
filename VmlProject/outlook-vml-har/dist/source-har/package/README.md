# outlook-vml-webview

用于 HarmonyOS ArkWeb 的 Microsoft Outlook 经典版 VML 邮件解析 HAR。

## 业务接口

- `changeVMLSrcs`：回复、转发、再次编辑链路，修正 VML `imagedata` 与普通 `img` 的资源文件名。
- `VML_P2DIV` / `prepareDisplayHtml`：HTML 注入 ArkWeb 前，把含 VML 的 `p` 替换为 `div`，避免 `p > span > table` 被截断。
- `prepareDisplayHtmlFile`：读取宿主应用 `resources/rawfile` 中的 HTML 文件并执行上述预处理。
- `parseVMLImages`：ArkWeb 页面 DOM 创建后，解析 VML 坐标并映射到普通图片。

## 本地集成

将模块源码复制到 HarmonyOS 工程，或引用构建后的 HAR：

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

## 阅读展示

```ts
import { webview } from '@kit.ArkWeb';
import {
  OutlookVmlSdk,
  prepareDisplayHtmlFile,
  PrepareDisplayResult,
  VmlReport
} from 'outlook-vml-webview';

@Entry
@Component
struct MailPage {
  private controller: webview.WebviewController = new webview.WebviewController();
  private vmlSdk: OutlookVmlSdk = new OutlookVmlSdk(this.controller);
  private html: string = '';
  private htmlReady: boolean = false;
  private initialPageEnded: boolean = false;
  private mailLoadStarted: boolean = false;
  private vmlParsed: boolean = false;

  aboutToAppear(): void {
    this.loadMailHtmlFile();
  }

  private async loadMailHtmlFile(): Promise<void> {
    const context = this.getUIContext().getHostContext();
    if (!context) {
      console.error('无法获取应用 Context');
      return;
    }

    // 对应 entry/src/main/resources/rawfile/mail/outlook.html
    const result: PrepareDisplayResult = await prepareDisplayHtmlFile(
      context.resourceManager,
      'mail/outlook.html'
    );
    if (result.report.status === 'failed') {
      console.error(`HTML读取或预处理失败：${JSON.stringify(result.report.errors)}`);
      return;
    }

    this.html = result.html;
    this.htmlReady = true;
    this.tryLoadMailHtml();
  }

  private tryLoadMailHtml(): void {
    if (!this.htmlReady || !this.initialPageEnded || this.mailLoadStarted) {
      return;
    }
    this.mailLoadStarted = true;
    this.controller.loadData(this.html, 'text/html', 'UTF-8', '', '');
  }

  private async parseWholeHtml(): Promise<void> {
    // 不传 selector，默认使用 document.documentElement 匹配整份 HTML。
    const report: VmlReport = await this.vmlSdk.parseVMLImages();
    console.info(`VML transformed=${report.stats.transformed}`);
  }

  build() {
    Web({ src: 'about:blank', controller: this.controller })
      .javaScriptAccess(true)
      .onPageEnd(() => {
        if (!this.initialPageEnded) {
          this.initialPageEnded = true;
          this.tryLoadMailHtml();
          return;
        }

        if (this.mailLoadStarted && !this.vmlParsed) {
          this.vmlParsed = true;
          this.parseWholeHtml();
        }
      });
  }
}
```

HTML 文件放在宿主应用的 `entry/src/main/resources/rawfile` 下。`prepareDisplayHtmlFile` 的路径相对于 rawfile 根目录，不能使用绝对路径或 `..`。HTML 内容通过 `WebviewController.loadData` 加载；`parseVMLImages()` 不传 selector 时使用 `document.documentElement`，即匹配整份 HTML，而不是依赖 `#mail-content`。它必须在邮件页面对应的 `onPageEnd` 后执行。

## 回复、转发、再次编辑

```ts
// 不传 selector，同样处理完整 document.documentElement。
const result = await vmlSdk.changeVMLSrcs();
if (result.report.status !== 'failed' && result.html !== null) {
  const editorHtml: string = result.html;
}
```

## 错误处理

页面解析错误沿用 `vmlparse.refactored.js` 的 I/V/M/D/S/P/X 错误码。HAR 适配层增加：

| 错误码 | 含义 |
|---|---|
| H7001 | selector 无效 |
| H7002 | Web 运行时注入失败 |
| H7003 | ArkWeb 脚本执行失败 |
| H7004 | ArkWeb 返回值无法解析 |
| H7005 | 注入后的运行时版本不一致 |
| H7006 | ArkTS 侧 HTML 参数无效 |
| H7007 | rawfile HTML 文件读取失败 |
| H7008 | rawfile 路径为空、越界或包含目录穿越 |

## 生成与构建

先生成嵌入 ArkTS 的 Web 运行时：

```shell
npm run generate
npm run validate
```

在 DevEco Studio 中把本目录作为 Static Library 模块导入，执行 `Build > Make Module`；或者在已有工程根目录使用 Hvigor：

```shell
hvigorw --mode module -p module=outlook_vml_webview@default assembleHar
```

也可生成供本地/私仓验证的源码 HAR：

```shell
npm run pack:source
```

正式发布前应使用目标 DevEco Studio/HarmonyOS SDK 执行 release HAR 构建，不应把手工生成的源码 HAR 当成已编译 release 产物。

## 发布前配置

当前包名、作者和 `UNLICENSED` 是安全占位配置。发布到公共 OHPM 前必须：

1. 换成实际拥有的包名或组织 scope。
2. 明确许可证并同步修改 `LICENSE` 和 `oh-package.json5`。
3. 使用目标 SDK 构建 release HAR，让 Hvigor写入正确的兼容性元数据。
4. 确认 README、CHANGELOG、LICENSE 非空。
