import { webview } from '@kit.ArkWeb';
import { fileIo } from '@kit.CoreFileKit';
import {
  OutlookVmlSdk,
  ParseHtmlResult,
  prepareDisplayHtmlFile,
  PrepareDisplayResult
} from 'outlook-vml-webview';

@Entry
@Component
struct Index {
  private runtimeController: webview.WebviewController = new webview.WebviewController();
  private displayController: webview.WebviewController = new webview.WebviewController();
  private vmlSdk: OutlookVmlSdk = new OutlookVmlSdk(this.runtimeController);
  private inputHtml: string = '';
  private htmlReady: boolean = false;
  private runtimePageReady: boolean = false;
  private transformStarted: boolean = false;
  private displayInitialPageReady: boolean = false;
  private displayLoadStarted: boolean = false;
  private displayVerified: boolean = false;
  private pendingDisplayHtml: string | null = null;
  private displayOutputPath: string = '';
  @State private statusMessage: string = '正在读取 HTML 文件…';

  aboutToAppear(): void {
    this.loadInputHtml();
  }

  private async loadInputHtml(): Promise<void> {
    const context = this.getUIContext().getHostContext();
    if (!context) {
      this.statusMessage = 'H7011：无法获取应用 Context';
      return;
    }

    const prepared: PrepareDisplayResult = await prepareDisplayHtmlFile(
      context.resourceManager,
      'mail/outlook.html'
    );
    if (prepared.report.status === 'failed') {
      const cause: string = prepared.report.errors.length > 0 ?
        prepared.report.errors[0].cause : '未返回错误原因';
      this.statusMessage = `HTML 预处理失败：${cause}`;
      return;
    }

    this.inputHtml = prepared.html;
    this.htmlReady = true;
    this.statusMessage = 'HTML 已读取，等待 VML 运行时…';
    this.tryTransformHtml();
  }

  private tryTransformHtml(): void {
    if (!this.htmlReady || !this.runtimePageReady || this.transformStarted) {
      return;
    }
    this.transformStarted = true;
    this.transformHtml();
  }

  private async transformHtml(): Promise<void> {
    try {
      const inputPath: string | null = this.writeHtmlFile(
        this.inputHtml,
        'vml-before-parse.html',
        'VML 解析前 HTML'
      );
      if (inputPath === null) {
        return;
      }

      this.statusMessage = '正在解析 VML HTML 字符串…';
      const result: ParseHtmlResult = await this.vmlSdk.parseVMLImagesFromHtml(this.inputHtml, {
        removeEmptyPlaceholderTables: true
      });
      if (result.report.status === 'failed' || result.html === null) {
        const code: string = result.report.errors.length > 0 ?
          result.report.errors[0].code : 'H7013';
        const cause: string = result.report.errors.length > 0 ?
          result.report.errors[0].cause : '未返回错误原因';
        this.statusMessage = `${code}：${cause}`;
        return;
      }

      const outputPath: string | null = this.writeHtmlFile(
        result.html,
        'vml-output.html',
        'VML 解析后 HTML'
      );
      if (outputPath === null) {
        return;
      }

      this.pendingDisplayHtml = result.html;
      this.displayOutputPath = outputPath;
      this.statusMessage = `转换完成：${result.report.stats.transformed} 张；等待窗口渲染…`;
      this.tryDisplayHtml();
    } catch (error) {
      this.statusMessage = `H7013：${String(error)}`;
      console.error(`VML HTML 字符串转换失败：${String(error)}`);
    }
  }

  private tryDisplayHtml(): void {
    if (!this.displayInitialPageReady || this.pendingDisplayHtml === null || this.displayLoadStarted) {
      return;
    }

    try {
      this.displayLoadStarted = true;
      this.statusMessage = '正在渲染 VML HTML…';
      const dataUrl: string = 'data:text/html;charset=utf-8,' +
        encodeURIComponent(this.pendingDisplayHtml);
      this.displayController.loadUrl(dataUrl);
      console.info(`VML HTML 已提交窗口渲染：${this.displayOutputPath}`);
    } catch (error) {
      this.displayLoadStarted = false;
      this.statusMessage = `H7013：窗口加载 HTML 失败：${String(error)}`;
    }
  }

  private async verifyDisplayedHtml(): Promise<void> {
    if (!this.displayLoadStarted || this.displayVerified) {
      return;
    }

    try {
      const bodyLengthResult: string = await this.displayController.runJavaScript(
        'String(document.body ? document.body.innerText.trim().length : 0)'
      );
      const bodyLength: number = Number(bodyLengthResult.replace(/^"+|"+$/g, ''));
      if (!Number.isFinite(bodyLength) || bodyLength <= 0) {
        this.statusMessage = 'H7013：窗口 HTML 已加载，但正文为空';
        return;
      }
      this.displayVerified = true;
      this.statusMessage = 'VML HTML 渲染完成';
      console.info(`VML HTML 窗口正文长度：${bodyLength}`);
    } catch (error) {
      this.statusMessage = `H7013：窗口 HTML 校验失败：${String(error)}`;
    }
  }

  private writeHtmlFile(html: string, fileName: string, description: string): string | null {
    const context = this.getUIContext().getHostContext();
    if (!context) {
      this.statusMessage = 'H7011：无法获取应用 Context';
      return null;
    }

    const outputPath: string = `${context.filesDir}/${fileName}`;
    let outputFile: fileIo.File | null = null;
    try {
      const mode: number = fileIo.OpenMode.WRITE_ONLY |
        fileIo.OpenMode.CREATE | fileIo.OpenMode.TRUNC;
      outputFile = fileIo.openSync(outputPath, mode);
      fileIo.writeSync(outputFile.fd, html);
      console.info(`${description}输出完成：${outputPath}`);
      return outputPath;
    } catch (error) {
      this.statusMessage = `H7011：${String(error)}`;
      return null;
    } finally {
      if (outputFile !== null) {
        try {
          fileIo.closeSync(outputFile);
        } catch (error) {
          console.error(`关闭输出文件失败：${String(error)}`);
        }
      }
    }
  }

  build() {
    Column() {
      Text(this.statusMessage)
        .fontSize(14)
        .fontColor('#333333')
        .width('100%')
        .padding(16);

      // 显示解析完成后的完整邮件 HTML。
      Web({ src: 'about:blank', controller: this.displayController })
        .javaScriptAccess(true)
        .onPageEnd(() => {
          if (!this.displayInitialPageReady) {
            this.displayInitialPageReady = true;
            this.tryDisplayHtml();
            return;
          }
          this.verifyDisplayedHtml();
        })
        .width('100%')
        .layoutWeight(1);

      // 只作为 DOMParser/VML JavaScript 的运行容器，不加载邮件 HTML。
      Web({ src: $rawfile('vml-runtime-host.html'), controller: this.runtimeController })
        .javaScriptAccess(true)
        .onPageEnd(() => {
          this.runtimePageReady = true;
          this.tryTransformHtml();
        })
        .width(1)
        .height(1)
        .opacity(0);
    }
    .width('100%')
    .height('100%');
  }
}
