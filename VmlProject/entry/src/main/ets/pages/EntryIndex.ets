import { webview } from '@kit.ArkWeb';
import {
  OutlookVmlSdk,
  prepareDisplayHtmlFile,
  PrepareDisplayResult,
  VmlReport
} from 'outlook-vml-webview';

/**
 * 宿主应用页面示例。
 *
 * 将本文件内容放到 entry/src/main/ets/pages/Index.ets。
 * HTML 示例文件放到 entry/src/main/resources/rawfile/mail/outlook.html。
 */
@Entry
@Component
struct Index {
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
    // 不传 selector：默认处理 document.documentElement，即完整 HTML。
    const report: VmlReport = await this.vmlSdk.parseVMLImages();
    console.info(`VML transformed=${report.stats.transformed}`);
  }

  build() {
    Column() {
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
        })
        .width('100%')
        .height('100%');
    }
    .width('100%')
    .height('100%');
  }
}
