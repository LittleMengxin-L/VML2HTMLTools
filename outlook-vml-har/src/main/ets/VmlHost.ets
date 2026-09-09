import { webview } from '@kit.ArkWeb';
import { fileIo } from '@kit.CoreFileKit';
import { OutlookVmlSdk } from './OutlookVmlSdk';
import { decodeHtmlContent, HtmlDecodeResult, normalizeCharsetDeclaration } from './VmlText';
import {
  HarErrorCode,
  VmlFileParseResult,
  VmlReportStatus
} from './VmlTypes';
import { createHostFailure } from './VmlReport';

const PAGE_READY_TIMEOUT_MS: number = 15000;
const HOST_STAGE: string = 'vml-host';
const HOST_ACTION: string = '请检查输入/输出文件路径与解析状态后重试。';

interface PendingReady {
  timer: number;
  resolve: (ready: boolean) => void;
}

/**
 * 内部隐藏 ArkWeb 页面的状态机：从创建、onPageEnd 就绪到注入运行时，
 * 全部由本类封闭管理。调用方只负责把 VmlHost 组件挂载到页面并调用
 * parseFile，无需了解 WebviewController、onPageEnd 或 JS 注入细节。
 */
export class VmlHostController {
  readonly controller: webview.WebviewController = new webview.WebviewController();
  private ready: boolean = false;
  private destroyed: boolean = false;
  private waiters: PendingReady[] = [];

  /**
   * 由 VmlHost 组件的 onPageEnd 回调调用，标记内部页面已就绪。
   * 组件重新挂载时允许从 destroyed 状态恢复。
   */
  notifyPageReady(): void {
    if (this.destroyed) {
      this.destroyed = false;
    }
    if (!this.ready) {
      this.ready = true;
    }
    this.flushWaiters(true);
  }

  /** 由 VmlHost 组件 aboutToDisappear 回调调用，释放等待中的调用方。 */
  markDestroyed(): void {
    this.destroyed = true;
    this.flushWaiters(false);
  }

  /**
   * 解析 HTML 文件中的 VML 图片。
   *
   * @param inputPath 源 HTML 文件路径（应用沙箱内可读绝对路径）。
   * @param outputPath 解析后 HTML 的落盘路径（父目录必须存在且可写）。
   * @returns 输出路径与结构化解析报告。
   */
  async parseFile(
    inputPath: string,
    outputPath: string
  ): Promise<VmlFileParseResult> {
    if (inputPath.trim().length === 0 || outputPath.trim().length === 0) {
      return {
        outputPath: null,
        report: createHostFailure(
          HarErrorCode.INVALID_HTML,
          '输入或输出文件路径为空。',
          HOST_STAGE,
          HOST_ACTION
        )
      };
    }

    const pageReady: boolean = await this.waitUntilReady();
    if (!pageReady) {
      return {
        outputPath: null,
        report: createHostFailure(
          HarErrorCode.PAGE_READY_TIMEOUT,
          '内部 ArkWeb 页面在限定时间内未就绪，无法执行解析。',
          HOST_STAGE,
          HOST_ACTION
        )
      };
    }

    const html: string | null = this.readTextFile(inputPath);
    if (html === null) {
      return {
        outputPath: null,
        report: createHostFailure(
          HarErrorCode.FILE_READ_FAILED,
          `读取输入 HTML 失败：${inputPath}`,
          HOST_STAGE,
          HOST_ACTION
        )
      };
    }

    const sdk: OutlookVmlSdk = new OutlookVmlSdk(this.controller);
    const parsed = await sdk.parseVMLImagesFromHtml(html);
    if (parsed.report.status === VmlReportStatus.FAILED || parsed.html === null) {
      return { outputPath: null, report: parsed.report };
    }

    const writeError: string | null = this.writeTextFile(
      outputPath,
      normalizeCharsetDeclaration(parsed.html)
    );
    if (writeError !== null) {
      return {
        outputPath: null,
        report: createHostFailure(
          HarErrorCode.OUTPUT_FILE_WRITE_FAILED,
          `写入解析结果失败：${writeError}`,
          HOST_STAGE,
          HOST_ACTION
        )
      };
    }
    return { outputPath: outputPath, report: parsed.report };
  }

  private waitUntilReady(): Promise<boolean> {
    if (this.ready) {
      return Promise.resolve(!this.destroyed);
    }
    return new Promise<boolean>((resolve: (ready: boolean) => void) => {
      const timer: number = setTimeout(() => {
        this.removeWaiter(timer);
        resolve(false);
      }, PAGE_READY_TIMEOUT_MS);
      this.waiters.push({ timer: timer, resolve: resolve });
    });
  }

  private removeWaiter(timer: number): void {
    let index: number = -1;
    for (let i: number = 0; i < this.waiters.length; i += 1) {
      if (this.waiters[i].timer === timer) {
        index = i;
        break;
      }
    }
    if (index >= 0) {
      this.waiters.splice(index, 1);
    }
  }

  private flushWaiters(ready: boolean): void {
    const pending: PendingReady[] = this.waiters;
    this.waiters = [];
    for (let i: number = 0; i < pending.length; i += 1) {
      clearTimeout(pending[i].timer);
      pending[i].resolve(ready);
    }
  }

  /**
   * 读取 HTML 文件并按内容实际编码解码为 Unicode 字符串。
   * 编码判定顺序：BOM → <meta charset>（gb2312/gbk 归一到 gb18030）→ 默认 UTF-8。
   * @returns 解码后的文本；文件不存在或读取失败时返回 null（错误细节写入 hilog）。
   *          空文件返回空字符串（由解析入口按“HTML 为空”处理）。
   */
  private readTextFile(path: string): string | null {
    let file: fileIo.File | null = null;
    try {
      file = fileIo.openSync(path, fileIo.OpenMode.READ_ONLY);
      const stat: fileIo.Stat = fileIo.statSync(path);
      if (stat.size <= 0) {
        return '';
      }
      // readSync 的 buffer 参数只接受 ArrayBuffer（不接受 Uint8Array 视图）。
      const buffer: ArrayBuffer = new ArrayBuffer(stat.size);
      const read: number = fileIo.readSync(file.fd, buffer);
      if (read < 0) {
        console.error(`VML readTextFile：readSync 返回异常 ${read}，路径 ${path}`);
        return null;
      }
      const decoded: HtmlDecodeResult = decodeHtmlContent(new Uint8Array(buffer, 0, read));
      if (decoded.unsupportedEncoding.length > 0) {
        console.warn(`VML readTextFile：无法识别编码 ${decoded.unsupportedEncoding}，按 UTF-8 降级，路径 ${path}`);
      }
      return decoded.html;
    } catch (error) {
      console.error(`VML readTextFile 失败：${String(error)}，路径 ${path}`);
      return null;
    } finally {
      if (file !== null) {
        try {
          fileIo.closeSync(file);
        } catch (error) {
          // 关闭失败可忽略，原始错误优先。
        }
      }
    }
  }

  private writeTextFile(path: string, content: string): string | null {
    let file: fileIo.File | null = null;
    try {
      const mode: number = fileIo.OpenMode.WRITE_ONLY |
        fileIo.OpenMode.CREATE | fileIo.OpenMode.TRUNC;
      file = fileIo.openSync(path, mode);
      fileIo.writeSync(file.fd, content);
      return null;
    } catch (error) {
      return String(error);
    } finally {
      if (file !== null) {
        try {
          fileIo.closeSync(file);
        } catch (error) {
          // 关闭失败可忽略，原始错误优先。
        }
      }
    }
  }
}

/**
 * 托管解析容器：内部创建 1×1 透明 ArkWeb 页面并驱动生命周期。
 * 调用方页面只需挂载一次（位置任意），例如：
 *
 *   Stack() {
 *     // 业务内容……
 *     VmlHost({ host: this.vmlHost })
 *   }
 */
@Component
export struct VmlHost {
  host: VmlHostController = new VmlHostController();

  aboutToDisappear(): void {
    this.host.markDestroyed();
  }

  build() {
    Web({ src: 'about:blank', controller: this.host.controller })
      .javaScriptAccess(true)
      .onPageEnd(() => {
        this.host.notifyPageReady();
      })
      .width(1)
      .height(1)
      .opacity(0);
  }
}
