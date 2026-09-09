import { util } from '@kit.ArkTS';
import { resourceManager } from '@kit.LocalizationKit';
import {
  HarErrorCode,
  PrepareDisplayResult,
  VmlIssue,
  VmlIssueSeverity,
  VmlReport,
  VmlReportStatus,
  VmlStats
} from './VmlTypes';

function createStats(): VmlStats {
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

function hasVmlCondition(value: string): boolean {
  return /<!--\s*\[if\s+(?:gte\s+)?vml\b/i.test(value);
}

function isTagBoundary(value: string): boolean {
  return value === '' || value === '>' || value === '/' || /\s/.test(value);
}

function findParagraphStart(lowerHtml: string, fromIndex: number): number {
  let index: number = lowerHtml.indexOf('<p', fromIndex);
  while (index >= 0) {
    const boundary: string = index + 2 < lowerHtml.length ? lowerHtml.charAt(index + 2) : '';
    if (isTagBoundary(boundary)) {
      return index;
    }
    index = lowerHtml.indexOf('<p', index + 2);
  }
  return -1;
}

function findTagEnd(source: string, startIndex: number): number {
  let quote: string = '';
  for (let index: number = startIndex; index < source.length; index += 1) {
    const character: string = source.charAt(index);
    if (quote.length > 0) {
      if (character === quote) {
        quote = '';
      }
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === '>') {
      return index;
    }
  }
  return -1;
}

function createParagraphIssue(code: HarErrorCode, reason: string): VmlIssue {
  return {
    code: code,
    severity: code === HarErrorCode.INVALID_HTML ? VmlIssueSeverity.ERROR : VmlIssueSeverity.WARNING,
    stage: 'paragraph',
    message: code === HarErrorCode.INVALID_HTML ? 'HTML 参数无效。' : 'VML 段落结构不完整。',
    cause: reason,
    action: '保留原始 HTML，不跨段落继续替换。'
  };
}

/**
 * 在 HTML 进入 ArkWeb 解析器前，把包含 VML 的 p 替换为 div，
 * 避免 p > span > table 被 HTML 内容模型提前截断。
 */
export function prepareDisplayHtml(html: string): PrepareDisplayResult {
  const stats: VmlStats = createStats();
  const issues: VmlIssue[] = [];

  if (typeof html !== 'string') {
    const issue: VmlIssue = createParagraphIssue(HarErrorCode.INVALID_HTML, '输入不是字符串。');
    issues.push(issue);
    return {
      html: '',
      report: {
        status: VmlReportStatus.FAILED,
        stats: stats,
        issues: issues,
        errors: issues,
        warnings: [],
        infos: []
      }
    };
  }

  const lowerHtml: string = html.toLowerCase();
  let cursor: number = 0;
  let output: string = '';

  while (cursor < html.length) {
    const openStart: number = findParagraphStart(lowerHtml, cursor);
    if (openStart < 0) {
      output += html.substring(cursor);
      break;
    }

    const openEnd: number = findTagEnd(html, openStart);
    if (openEnd < 0) {
      if (hasVmlCondition(html.substring(openStart))) {
        issues.push(createParagraphIssue(HarErrorCode.INVALID_HTML, '含 VML 的 p 开始标签缺少 >。'));
      }
      output += html.substring(cursor);
      break;
    }

    const closeStart: number = lowerHtml.indexOf('</p', openEnd + 1);
    if (closeStart < 0) {
      if (hasVmlCondition(html.substring(openStart))) {
        issues.push(createParagraphIssue(HarErrorCode.INVALID_HTML, '含 VML 的 p 缺少闭合标签。'));
      }
      output += html.substring(cursor);
      break;
    }

    const closeEnd: number = findTagEnd(html, closeStart);
    if (closeEnd < 0) {
      if (hasVmlCondition(html.substring(openStart))) {
        issues.push(createParagraphIssue(HarErrorCode.INVALID_HTML, '含 VML 的 p 闭合标签不完整。'));
      }
      output += html.substring(cursor);
      break;
    }

    const block: string = html.substring(openStart, closeEnd + 1);
    output += html.substring(cursor, openStart);
    if (hasVmlCondition(block)) {
      const openTag: string = html.substring(openStart, openEnd + 1);
      output += '<div' + openTag.substring(2) +
        html.substring(openEnd + 1, closeStart) + '</div>';
      stats.paragraphsConverted += 1;
    } else {
      output += block;
    }
    cursor = closeEnd + 1;
  }

  const errors: VmlIssue[] = issues.filter((item: VmlIssue) => item.severity === VmlIssueSeverity.ERROR);
  const warnings: VmlIssue[] = issues.filter((item: VmlIssue) => item.severity === VmlIssueSeverity.WARNING);
  return {
    html: output,
    report: {
      status: errors.length > 0 ? VmlReportStatus.FAILED :
        (warnings.length > 0 ? VmlReportStatus.PARTIAL : VmlReportStatus.OK),
      stats: stats,
      issues: issues,
      errors: errors,
      warnings: warnings,
      infos: []
    }
  };
}

/** 与 Web 运行时同名的兼容函数。 */
export function VML_P2DIV(html: string): string {
  return prepareDisplayHtml(html).html;
}

function createFileFailure(code: HarErrorCode, cause: string): PrepareDisplayResult {
  const issue: VmlIssue = {
    code: code,
    severity: VmlIssueSeverity.ERROR,
    stage: 'rawfile',
    message: code === HarErrorCode.INVALID_RAWFILE_PATH ?
      'HTML 文件路径无效。' : '读取 HTML 文件失败。',
    cause: cause,
    action: '不加载 HTML，保留当前 Web 页面。'
  };
  return {
    html: '',
    report: {
      status: VmlReportStatus.FAILED,
      stats: createStats(),
      issues: [issue],
      errors: [issue],
      warnings: [],
      infos: []
    }
  };
}

function isValidRawFilePath(rawFilePath: string): boolean {
  return rawFilePath.length > 0 && rawFilePath.length <= 1024 &&
    !rawFilePath.startsWith('/') &&
    !rawFilePath.startsWith('\\') &&
    rawFilePath.indexOf('..') < 0 &&
    rawFilePath.indexOf('\u0000') < 0 &&
    rawFilePath.indexOf('\r') < 0 &&
    rawFilePath.indexOf('\n') < 0;
}

/**
 * 从宿主应用 resources/rawfile 读取 HTML 文件，并在进入 ArkWeb 前修复 VML 段落。
 * rawFilePath 相对于 src/main/resources/rawfile，例如 mail/outlook.html。
 */
export async function prepareDisplayHtmlFile(
  resMgr: resourceManager.ResourceManager,
  rawFilePath: string
): Promise<PrepareDisplayResult> {
  if (!isValidRawFilePath(rawFilePath)) {
    return createFileFailure(HarErrorCode.INVALID_RAWFILE_PATH, '路径为空、过长、包含控制字符或目录穿越。');
  }

  try {
    const content: Uint8Array = await resMgr.getRawFileContent(rawFilePath);
    const decoder: util.TextDecoder = new util.TextDecoder();
    const html: string = decoder.decodeToString(content);
    return prepareDisplayHtml(html);
  } catch (error) {
    return createFileFailure(HarErrorCode.RAWFILE_READ_FAILED, String(error));
  }
}
