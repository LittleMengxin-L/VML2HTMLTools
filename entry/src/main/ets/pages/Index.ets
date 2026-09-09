import { VmlHost, VmlHostController, VmlFileParseResult } from 'outlook-vml-webview';

@Entry
@Component
struct Index {
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