import { webview } from '@kit.ArkWeb';
import { VML_RUNTIME_SOURCE, VML_RUNTIME_VERSION } from './VmlRuntime';
import {
  ChangeVmlResult,
  HarErrorCode,
  VmlIssue,
  VmlIssueSeverity,
  VmlParserOptions,
  VmlReport,
  VmlReportStatus,
  VmlStats
} from './VmlTypes';

function createEmptyStats(): VmlStats {
  return {
    shapes: 0,
    images: 0,
    matched: 0,
    transformed: 0,
    skipped: 0,
    sourcesRewritten: 0,
    paragraphsConverted: 0
  };
}

function createHostFailure(code: HarErrorCode, message: string): VmlReport {
  const issue: VmlIssue = {
    code: code,
    severity: VmlIssueSeverity.ERROR,
    stage: 'harmony-adapter',
    message: message,
    cause: message,
    action: '未修改邮件 DOM，请检查 ArkWeb 页面状态和 JavaScript 权限。'
  };
  return {
    status: VmlReportStatus.FAILED,
    stats: createEmptyStats(),
    issues: [issue],
    errors: [issue],
    warnings: [],
    infos: []
  };
}

function isValidSelector(selector: string): boolean {
  return selector.length > 0 && selector.length <= 1024 &&
    selector.indexOf('\u0000') < 0 && selector.indexOf('\r') < 0 && selector.indexOf('\n') < 0;
}

function encodeForJavaScript(value: Object | string): string {
  const encoded: string | undefined = JSON.stringify(value);
  return encoded === undefined ? 'null' : encoded;
}

/**
 * HarmonyOS HAR 适配器。
 * changeVMLSrcs/parseVMLImages 会通过 WebviewController 在当前页面上下文执行，
 * 因为 HTMLElement/document 只存在于 ArkWeb 页面中。
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
      const currentVersion: string = await this.controller.runJavaScript(versionScript);
      if (currentVersion === VML_RUNTIME_VERSION) {
        return null;
      }

      await this.controller.runJavaScript(VML_RUNTIME_SOURCE);
      const injectedVersion: string = await this.controller.runJavaScript(versionScript);
      if (injectedVersion !== VML_RUNTIME_VERSION) {
        return createHostFailure(
          HarErrorCode.RUNTIME_VERSION_MISMATCH,
          `VML 运行时版本不一致，期望 ${VML_RUNTIME_VERSION}，实际 ${injectedVersion}。`
        );
      }
      return null;
    } catch (error) {
      return createHostFailure(
        HarErrorCode.RUNTIME_INJECTION_FAILED,
        `VML 运行时注入失败：${String(error)}`
      );
    }
  }

  private parseReport(value: string): VmlReport {
    try {
      return JSON.parse(value) as VmlReport;
    } catch (error) {
      return createHostFailure(
        HarErrorCode.INVALID_SCRIPT_RESULT,
        `ArkWeb 返回值不是有效的 VML 报告：${String(error)}`
      );
    }
  }

  /** 阅读展示链路：页面 onPageEnd/DOM 创建完成后调用。 */
  async parseVMLImages(
    selector: string = 'html',
    options: VmlParserOptions = {}
  ): Promise<VmlReport> {
    if (!isValidSelector(selector)) {
      return createHostFailure(HarErrorCode.INVALID_SELECTOR, 'selector 为空、过长或包含控制字符。');
    }

    const runtimeFailure: VmlReport | null = await this.ensureRuntime();
    if (runtimeFailure !== null) {
      return runtimeFailure;
    }

    const script: string = `(() => {
      try {
        const root = ${encodeForJavaScript(selector)} === 'html' ?
          document.documentElement : document.querySelector(${encodeForJavaScript(selector)});
        const report = window.OutlookVMLParser.parseVMLImages(
          root,
          ${encodeForJavaScript(options as Object)}
        );
        return JSON.stringify(report);
      } catch (error) {
        return JSON.stringify({
          status: 'failed',
          stats: { shapes: 0, images: 0, matched: 0, transformed: 0, skipped: 0,
            sourcesRewritten: 0, paragraphsConverted: 0 },
          issues: [{ code: 'H7003', severity: 'error', stage: 'harmony-adapter',
            message: String(error), cause: String(error), action: '未修改邮件 DOM。' }],
          errors: [{ code: 'H7003', severity: 'error', stage: 'harmony-adapter',
            message: String(error), cause: String(error), action: '未修改邮件 DOM。' }],
          warnings: [], infos: []
        });
      }
    })()`;

    try {
      return this.parseReport(await this.controller.runJavaScript(script));
    } catch (error) {
      return createHostFailure(
        HarErrorCode.SCRIPT_EXECUTION_FAILED,
        `parseVMLImages 执行失败：${String(error)}`
      );
    }
  }

  /** 回复、转发、再次编辑链路：返回修正后的编辑区 HTML。 */
  async changeVMLSrcs(selector: string = 'html'): Promise<ChangeVmlResult> {
    if (!isValidSelector(selector)) {
      return {
        html: null,
        report: createHostFailure(HarErrorCode.INVALID_SELECTOR, 'selector 为空、过长或包含控制字符。')
      };
    }

    const runtimeFailure: VmlReport | null = await this.ensureRuntime();
    if (runtimeFailure !== null) {
      return { html: null, report: runtimeFailure };
    }

    const script: string = `(() => {
      const root = ${encodeForJavaScript(selector)} === 'html' ?
        document.documentElement : document.querySelector(${encodeForJavaScript(selector)});
      const report = window.OutlookVMLParser.changeVMLSrcs(root);
      return JSON.stringify({ html: root ? root.innerHTML : null, report: report });
    })()`;

    try {
      const result: ChangeVmlResult = JSON.parse(
        await this.controller.runJavaScript(script)
      ) as ChangeVmlResult;
      return result;
    } catch (error) {
      return {
        html: null,
        report: createHostFailure(
          HarErrorCode.SCRIPT_EXECUTION_FAILED,
          `changeVMLSrcs 执行失败：${String(error)}`
        )
      };
    }
  }
}
