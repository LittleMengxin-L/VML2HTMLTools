/*
 * Copyright (c) 2026 Huawei Technologies Co., Ltd. All rights reserved.
 *
 * SPDX-License-Identifier: MIT
 * Licensed under the MIT License. See LICENSE in the HAR root for details.
 */

import { webview } from '@kit.ArkWeb';
import { VML_RUNTIME_SOURCE, VML_RUNTIME_VERSION } from './VmlRuntime';
import { HarErrorCode, ParseHtmlResult, VmlReport } from './VmlTypes';
import { createHostFailure } from './VmlReport';

const HOST_STAGE: string = 'harmony-adapter';
const HOST_ACTION: string = '未修改邮件 DOM，请检查 ArkWeb 页面状态和 JavaScript 权限。';

/** 将参数序列化为可嵌入运行时脚本的 JSON 表达式。 */
function encodeForJavaScript(value: Object | string): string {
  const encoded: string | undefined = JSON.stringify(value);
  return encoded === undefined ? 'null' : encoded;
}

/**
 * 归一化 ArkWeb 返回的字符串，最多解除两层 JSON 字符串编码。
 */
function unwrapJavaScriptString(value: string): string {
  let result: string = value.trim();
  for (let depth: number = 0; depth < 2; depth += 1) {
    if (result.length < 2 || result.charAt(0) !== '"' || result.charAt(result.length - 1) !== '"') {
      break;
    }
    try {
      const decoded: Object | string = JSON.parse(result) as Object | string;
      if (typeof decoded !== 'string') {
        break;
      }
      result = decoded.trim();
    } catch (_) {
      break;
    }
  }
  return result;
}

/**
 * 基于 ArkWeb 的 Outlook VML HTML 转换接口。
 * 使用独立 Document 解析 HTML，转换过程不等待图片资源加载。
 * 调用前须确保控制器已关联就绪的 Web 组件，并启用 JavaScript。
 */
export class OutlookVmlSdk {
  private readonly controller: webview.WebviewController;

  /** @param controller 已关联 Web 组件的 ArkWeb 控制器。 */
  constructor(controller: webview.WebviewController) {
    this.controller = controller;
  }

  /** 检查运行时版本，按需注入脚本；成功返回 null，失败返回诊断报告。 */
  private async ensureRuntime(): Promise<VmlReport | null> {
    try {
      const versionScript: string =
        'typeof window.OutlookVMLParser === "object" && ' +
        'String(window.OutlookVMLParser.VERSION || "")';
      const currentVersion: string = unwrapJavaScriptString(
        await this.controller.runJavaScript(versionScript)
      );
      if (currentVersion === VML_RUNTIME_VERSION) {
        return null;
      }

      await this.controller.runJavaScript(VML_RUNTIME_SOURCE);
      const injectedVersion: string = unwrapJavaScriptString(
        await this.controller.runJavaScript(versionScript)
      );
      if (injectedVersion !== VML_RUNTIME_VERSION) {
        return createHostFailure(
          HarErrorCode.RUNTIME_VERSION_MISMATCH,
          `VML 运行时版本不一致，期望 ${VML_RUNTIME_VERSION}，实际 ${injectedVersion}。`,
          HOST_STAGE,
          HOST_ACTION
        );
      }
      return null;
    } catch (error) {
      return createHostFailure(
        HarErrorCode.RUNTIME_INJECTION_FAILED,
        `VML 运行时注入失败：${String(error)}`,
        HOST_STAGE,
        HOST_ACTION
      );
    }
  }

  /**
   * 转换完整 HTML 中的 VML 图片位置与源路径，返回序列化的 HTML。
   * 先转换 VML 段落，再同步图片源并应用布局。
   * @param html 待转换的非空 HTML 字符串。
   * @returns HTML 与诊断报告；调用方应检查 report.status 和 html 是否为 null。
   */
  async parseVMLImagesFromHtml(
    html: string
  ): Promise<ParseHtmlResult> {
    if (html.trim().length === 0) {
      return {
        html: null,
        report: createHostFailure(
          HarErrorCode.INVALID_HTML,
          '待解析的 HTML 字符串为空。',
          HOST_STAGE,
          HOST_ACTION
        )
      };
    }

    const runtimeFailure: VmlReport | null = await this.ensureRuntime();
    if (runtimeFailure !== null) {
      return { html: null, report: runtimeFailure };
    }

    const script: string = `(() => {
      try {
        const sourceHtml = ${encodeForJavaScript(html)};
        const report = window.OutlookVMLParser.createReport();
        // 在创建 Document 前转换 VML 段落，使其进入 div/li/h2 定位容器扫描范围。
        const normalizedHtml = window.OutlookVMLParser.VML_P2DIV(sourceHtml, report);
        const parser = new DOMParser();
        const parsedDocument = parser.parseFromString(normalizedHtml, 'text/html');
        const root = parsedDocument.documentElement;
        window.OutlookVMLParser.parseVMLImages(
          root,
          report
        );
        const doctype = parsedDocument.doctype ?
          '<!DOCTYPE ' + parsedDocument.doctype.name + '>\\n' : '';
        return JSON.stringify({ html: doctype + root.outerHTML, report: report });
      } catch (error) {
        return JSON.stringify({
          html: null,
          report: {
            status: 'failed',
            stats: { shapes: 0, images: 0, matched: 0, transformed: 0, skipped: 0,
              sourcesRewritten: 0, paragraphsConverted: 0 },
            issues: [{ code: 'H7013', severity: 'error', stage: 'html-transform',
              message: String(error), cause: String(error), action: '未输出转换后的 HTML。' }],
            errors: [{ code: 'H7013', severity: 'error', stage: 'html-transform',
              message: String(error), cause: String(error), action: '未输出转换后的 HTML。' }],
            warnings: [], infos: []
          }
        });
      }
    })()`;

    try {
      return JSON.parse(
        unwrapJavaScriptString(await this.controller.runJavaScript(script))
      ) as ParseHtmlResult;
    } catch (error) {
      return {
        html: null,
        report: createHostFailure(
          HarErrorCode.HTML_TRANSFORM_FAILED,
          `HTML 字符串转换失败：${String(error)}`,
          HOST_STAGE,
          HOST_ACTION
        )
      };
    }
  }
}
