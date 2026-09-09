import { webview } from '@kit.ArkWeb';
import { VML_RUNTIME_SOURCE, VML_RUNTIME_VERSION } from './VmlRuntime';
import { HarErrorCode, ParseHtmlResult, VmlParserOptions, VmlReport } from './VmlTypes';
import { createHostFailure } from './VmlReport';

const HOST_STAGE: string = 'harmony-adapter';
const HOST_ACTION: string = '未修改邮件 DOM，请检查 ArkWeb 页面状态和 JavaScript 权限。';

function encodeForJavaScript(value: Object | string): string {
  const encoded: string | undefined = JSON.stringify(value);
  return encoded === undefined ? 'null' : encoded;
}

/**
 * ArkWeb 在不同系统版本中可能返回原始字符串，也可能返回 JSON 编码后的字符串。
 * 最多解包两层，既兼容 `1.0.0`，也兼容 `"1.0.0"` 以及被再次编码的 JSON 报告。
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
 * HarmonyOS HAR 适配器（离线转换入口）。
 * 通过 WebviewController 在托管页面的 JavaScript 上下文里执行转换；
 * 字符串入/出，不把邮件加载为当前页面，也不依赖 img/cid 等资源完成加载。
 */
export class OutlookVmlSdk {
  private readonly controller: webview.WebviewController;

  constructor(controller: webview.WebviewController) {
    this.controller = controller;
  }

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
   * 纯字符串转换：输入完整 HTML，在 ArkWeb 的离线 Document 中解析 VML，输出完整 HTML。
   * 不把邮件加载为当前页面，也不依赖 img/cid 等资源完成加载。
   */
  async parseVMLImagesFromHtml(
    html: string,
    options: VmlParserOptions = {}
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
        // 字符串入口必须自行完成 VML 段落预处理。否则 td > p 内的 VML
        // 不属于 parseVMLImages 扫描的 div/li/h2 容器，几何信息会被整体跳过。
        const normalizedHtml = window.OutlookVMLParser.VML_P2DIV(sourceHtml, report);
        const parser = new DOMParser();
        const parsedDocument = parser.parseFromString(normalizedHtml, 'text/html');
        const root = parsedDocument.documentElement;
        window.OutlookVMLParser.parseVMLImages(
          root,
          ${encodeForJavaScript(options as Object)},
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
