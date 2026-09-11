import { webview } from '@kit.ArkWeb';
import { fileIo, picker } from '@kit.CoreFileKit';
import { VmlHost, VmlHostController, VmlFileParseResult } from 'outlook-vml-webview';

@Entry
@Component
struct Index {
  private host: VmlHostController = new VmlHostController();
  private displayController: webview.WebviewController = new webview.WebviewController();
  private displayInitialPageReady: boolean = false;
  private displayLoadStarted: boolean = false;
  private displayVerified: boolean = false;
  private pendingDisplayHtml: string | null = null;
  private displayOutputPath: string = '';
  private startupPickerOpened: boolean = false;
  @State private busy: boolean = false;
  @State private canExport: boolean = false;
  @State private fileMessage: string = '';
  @State private statusMessage: string = '请选择待解析的 HTML 文件';

  private async parseHtmlFile(): Promise<void> {
    if (this.busy) {
      return;
    }
    const context = this.getUIContext().getHostContext();
    if (!context) {
      this.statusMessage = 'H7011：无法获取应用 Context';
      return;
    }
    this.busy = true;
    let stage: string = '选择输入文件';
    try {
      this.statusMessage = '请选择 HTML 文件所在目录及文件…';
      const documentPicker: picker.DocumentViewPicker = new picker.DocumentViewPicker(context);
      const options: picker.DocumentSelectOptions = new picker.DocumentSelectOptions();
      options.maxSelectNumber = 1;
      options.fileSuffixFilters = ['HTML|.html,.htm'];
      const selected: string[] = await documentPicker.select(options);
      if (selected.length === 0) {
        this.statusMessage = '已取消选择，可点击“选择 HTML”重试';
        return;
      }
      stage = '导入输入文件';
      this.statusMessage = '正在导入 HTML…';
      await this.copyDocument(selected[0], `${context.filesDir}/mail-input.html`);
      this.canExport = false;
      this.pendingDisplayHtml = null;
      this.displayLoadStarted = false;
      this.displayVerified = false;
      this.fileMessage = '';
      stage = '解析 HTML';
      this.statusMessage = '正在等待 VML 运行时并解析 HTML…';
      const result: VmlFileParseResult = await this.host.parseFile(
        `${context.filesDir}/mail-input.html`,
        `${context.filesDir}/mail-output.html`
      );
      if (result.report.status === 'failed' || result.outputPath === null) {
        const code: string = result.report.errors.length > 0 ?
          result.report.errors[0].code : 'H7013';
        const cause: string = result.report.errors.length > 0 ?
          result.report.errors[0].cause : '未返回错误原因';
        this.statusMessage = `${code}：${cause}`;
        console.error(`解析失败：${this.statusMessage}`);
        return;
      }
      // parseFile 已将完整 HTML 以 UTF-8 写入结果文件。
      try {
        this.pendingDisplayHtml = fileIo.readTextSync(result.outputPath);
      } catch (error) {
        this.statusMessage = `H7014：读取解析结果失败：${String(error)}`;
        return;
      }
      this.displayOutputPath = result.outputPath;
      this.canExport = true;
      this.statusMessage = `转换完成：${result.report.stats.transformed} 张；等待窗口渲染…`;
      this.tryDisplayHtml();
      await this.chooseOutputFile();
    } catch (error) {
      this.statusMessage = `${stage}未完成：${String(error)}。可重新选择文件。`;
    } finally {
      this.busy = false;
    }
  }

  // 系统选择器返回授权 URI；直接打开 URI，不转换为设备物理路径。
  private async copyDocument(source: string, destination: string): Promise<void> {
    let input: fileIo.File | null = null;
    let output: fileIo.File | null = null;
    try {
      input = await fileIo.open(source, fileIo.OpenMode.READ_ONLY);
      output = await fileIo.open(destination,
        fileIo.OpenMode.READ_WRITE | fileIo.OpenMode.CREATE | fileIo.OpenMode.TRUNC);
      await fileIo.copyFile(input.fd, output.fd);
    } finally {
      if (output !== null) {
        try { await fileIo.close(output); } catch (error) {
          console.error(`关闭输出文件失败：${String(error)}`);
        }
      }
      if (input !== null) {
        try { await fileIo.close(input); } catch (error) {
          console.error(`关闭输入文件失败：${String(error)}`);
        }
      }
    }
  }

  private async chooseOutputFile(): Promise<void> {
    const context = this.getUIContext().getHostContext();
    if (!context || !this.canExport) {
      return;
    }
    try {
      this.fileMessage = '请选择输出目录及 HTML 文件名…';
      const documentPicker: picker.DocumentViewPicker = new picker.DocumentViewPicker(context);
      const options: picker.DocumentSaveOptions = new picker.DocumentSaveOptions();
      options.newFileNames = [`mail-output-${Date.now()}.html`];
      const selected: string[] = await documentPicker.save(options);
      if (selected.length === 0) {
        this.fileMessage = '已取消导出，解析结果已保留，可点击“另存为”';
        return;
      }
      await this.copyDocument(this.displayOutputPath, selected[0]);
      this.fileMessage = `已保存：${selected[0]}`;
    } catch (error) {
      this.fileMessage = `导出未完成：${String(error)}。请点击“另存为”重新选择可写位置。`;
    }
  }

  private async exportAgain(): Promise<void> {
    if (this.busy || !this.canExport) {
      return;
    }
    this.busy = true;
    try {
      await this.chooseOutputFile();
    } finally {
      this.busy = false;
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
        'String(document.body ? document.body.innerText.trim().length + document.body.querySelectorAll("img").length : 0)'
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

  build() {
    Column() {
      Row({ space: 12 }) {
        Button('选择 HTML')
          .enabled(!this.busy)
          .onClick(() => { this.parseHtmlFile(); });
        Button('另存为')
          .enabled(!this.busy && this.canExport)
          .onClick(() => { this.exportAgain(); });
      }
      .padding(12);

      Text(this.statusMessage)
        .fontSize(14)
        .fontColor('#333333')
        .width('100%')
        .padding(16);

      if (this.fileMessage.length > 0) {
        Text(this.fileMessage)
          .fontSize(12)
          .width('100%')
          .padding(12);
      }

      Web({ src: 'about:blank', controller: this.displayController })
        .javaScriptAccess(true)
        .onPageEnd(() => {
          if (!this.displayInitialPageReady) {
            this.displayInitialPageReady = true;
            this.tryDisplayHtml();
            if (!this.startupPickerOpened) {
              this.startupPickerOpened = true;
              this.parseHtmlFile();
            }
            return;
          }
          this.verifyDisplayedHtml();
        })
        .width('100%')
        .layoutWeight(1);

      // VmlHost 内部托管隐藏的解析 WebView。
      VmlHost({ host: this.host });
    }
    .width('100%')
    .height('100%');
  }
}
