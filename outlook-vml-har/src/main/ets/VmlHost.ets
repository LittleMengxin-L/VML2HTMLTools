/*
 * Copyright (c) 2026 Huawei Technologies Co., Ltd. All rights reserved.
 *
 * SPDX-License-Identifier: MIT
 * Licensed under the MIT License. See LICENSE in the HAR root for details.
 */

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
 * VML 文件转换控制器，管理解析页面的就绪状态和文件读写。
 * 使用时须将同一个控制器实例传入页面中的 VmlHost 组件。
 */
export class VmlHostController {
  readonly controller: webview.WebviewController = new webview.WebviewController();
  private ready: boolean = false;
  private destroyed: boolean = false;
  private waiters: PendingReady[] = [];

  /**
   * 通知内部页面已就绪并释放等待任务，由 VmlHost 的 onPageEnd 回调调用。
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

  /** 标记宿主已销毁并结束等待任务，由 VmlHost 生命周期回调调用。 */
  markDestroyed(): void {
    this.destroyed = true;
    this.flushWaiters(false);
  }

  /**
   * 解析 HTML 文件中的 VML 图片。
   *
   * @param inputPath 源 HTML 文件路径（应用沙箱内可读绝对路径）。
   * @param outputPath 输出 HTML 的绝对路径，父目录须存在且可写；已有文件将被覆盖。
   * @returns 输出路径与诊断报告；解析或写入失败时 outputPath 为 null。
   * @remarks 等待内部页面就绪最多 15 秒，输出文件使用 UTF-8 编码。
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

  /** 等待内部页面就绪；超时或宿主销毁时返回 false。 */
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

  /** 移除指定计时器对应的等待任务。 */
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

  /** 清理计时器并统一结束所有等待任务。 */
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
   * @returns 解码后的文本；读取失败返回 null 并记录日志，空文件返回空字符串。
   */
  private readTextFile(path: string): string | null {
    let file: fileIo.File | null = null;
    try {
      file = fileIo.openSync(path, fileIo.OpenMode.READ_ONLY);
      const stat: fileIo.Stat = fileIo.statSync(path);
      if (stat.size <= 0) {
        return '';
      }
      // 使用 ArrayBuffer 接收文件字节。
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
          // 保留读写操作的返回结果，关闭异常不覆盖该结果。
        }
      }
    }
  }

  /** 以 UTF-8 覆盖写入文件；成功返回 null，失败返回错误描述。 */
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
          // 保留读写操作的返回结果，关闭异常不覆盖该结果。
        }
      }
    }
  }
}

/**
 * VML 解析宿主组件，内部使用 1×1 透明 ArkWeb 页面。
 * 每个控制器关联一个宿主组件；解析期间应保持组件挂载。
 * @example
 *
 *   Stack() {
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
