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
 * 清理只服务于 IE/MSO 的旧式样式包装，避免 ArkWeb 把后续正文吞进注释或 style。
 * VML shape 条件注释会原样保留，供页面运行时从 innerHTML 中提取坐标。
 */
function normalizeOutlookHtmlForArkWeb(html: string): string {
  return html
    // 这段 behavior:url(#default#VML) 只对 IE 生效，ArkWeb 不需要。
    .replace(
      /<!--\s*\[if\s+!mso\]\s*>\s*<style\b[^>]*>[\s\S]*?<\/style>\s*<!\[endif\]\s*-->/gi,
      ''
    )
    // Word 的 shapedefaults/shapelayout 只服务于 Office，不参与 WebView 布局。
    .replace(
      /<!--\s*\[if\s+gte\s+mso\s+\d+\]\s*>\s*<xml\b[\s\S]*?<\/xml>\s*<!\[endif\]\s*-->/gi,
      ''
    )
    // downlevel-revealed 条件语法不是标准 HTML，移除外壳并保留普通图片正文。
    .replace(/<!\[if\s+!vml\]>([\s\S]*?)<!\[endif\]>/gi, '$1')
    // REC-html40 默认 namespace 会让部分 ArkWeb 版本走非预期的 XML/外来内容分支。
    .replace(/\s+xmlns(?::[a-z][\w.-]*)?\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, '')
    // HTML 已经解码成 ArkTS Unicode 字符串，交给 loadData 时统一声明 UTF-8。
    .replace(/charset\s*=\s*["']?\s*[a-z0-9._-]+/gi, 'charset=utf-8')
    // Outlook 会在普通 style 内再套一层 HTML 注释，现代 WebView 不需要该包装。
    .replace(/(<style\b[^>]*>)\s*<!--/gi, '$1')
    .replace(/-->\s*(<\/style\s*>)/gi, '$1');
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

  const sourceHtml: string = normalizeOutlookHtmlForArkWeb(html);
  const lowerHtml: string = sourceHtml.toLowerCase();
  let cursor: number = 0;
  let output: string = '';

  while (cursor < sourceHtml.length) {
    const openStart: number = findParagraphStart(lowerHtml, cursor);
    if (openStart < 0) {
      output += sourceHtml.substring(cursor);
      break;
    }

    const openEnd: number = findTagEnd(sourceHtml, openStart);
    if (openEnd < 0) {
      if (hasVmlCondition(sourceHtml.substring(openStart))) {
        issues.push(createParagraphIssue(HarErrorCode.INVALID_HTML, '含 VML 的 p 开始标签缺少 >。'));
      }
      output += sourceHtml.substring(cursor);
      break;
    }

    const closeStart: number = lowerHtml.indexOf('</p', openEnd + 1);
    if (closeStart < 0) {
      if (hasVmlCondition(sourceHtml.substring(openStart))) {
        issues.push(createParagraphIssue(HarErrorCode.INVALID_HTML, '含 VML 的 p 缺少闭合标签。'));
      }
      output += sourceHtml.substring(cursor);
      break;
    }

    const closeEnd: number = findTagEnd(sourceHtml, closeStart);
    if (closeEnd < 0) {
      if (hasVmlCondition(sourceHtml.substring(openStart))) {
        issues.push(createParagraphIssue(HarErrorCode.INVALID_HTML, '含 VML 的 p 闭合标签不完整。'));
      }
      output += sourceHtml.substring(cursor);
      break;
    }

    const block: string = sourceHtml.substring(openStart, closeEnd + 1);
    output += sourceHtml.substring(cursor, openStart);
    if (hasVmlCondition(block)) {
      const openTag: string = sourceHtml.substring(openStart, openEnd + 1);
      output += '<div' + openTag.substring(2) +
        sourceHtml.substring(openEnd + 1, closeStart) + '</div>';
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

function detectBomEncoding(content: Uint8Array): string | null {
  if (content.length >= 3 && content[0] === 0xEF && content[1] === 0xBB && content[2] === 0xBF) {
    return 'utf-8';
  }
  if (content.length >= 2 && content[0] === 0xFF && content[1] === 0xFE) {
    return 'utf-16le';
  }
  if (content.length >= 2 && content[0] === 0xFE && content[1] === 0xFF) {
    return 'utf-16be';
  }
  return null;
}

function normalizeHtmlEncoding(label: string): string | null {
  const value: string = label.trim().toLowerCase().replace('_', '-');
  if (value === 'utf-8' || value === 'utf8') {
    return 'utf-8';
  }
  if (value === 'gb2312' || value === 'gbk' || value === 'x-gbk' || value === 'gb18030') {
    // GB18030 向下兼容 GBK/GB2312，能覆盖经典版 Outlook 中文邮件。
    return 'gb18030';
  }
  if (value === 'big5' || value === 'big5-hkscs') {
    return 'big5';
  }
  if (value === 'utf-16' || value === 'utf-16le') {
    return 'utf-16le';
  }
  if (value === 'utf-16be') {
    return 'utf-16be';
  }
  if (value === 'windows-1252' || value === 'iso-8859-1') {
    return value;
  }
  return null;
}

function detectHtmlEncoding(content: Uint8Array): string | null {
  const bomEncoding: string | null = detectBomEncoding(content);
  if (bomEncoding !== null) {
    return bomEncoding;
  }

  // charset 标记本身由 ASCII 字符组成，用 iso-8859-1 扫描头部不会破坏字节。
  const previewLength: number = Math.min(content.length, 4096);
  const preview: string = util.TextDecoder.create('iso-8859-1')
    .decodeToString(content.subarray(0, previewLength));
  const charsetMatch: RegExpMatchArray | null = preview.match(
    /charset\s*=\s*["']?\s*([a-z0-9._-]+)/i
  );
  return charsetMatch === null ? 'utf-8' : normalizeHtmlEncoding(charsetMatch[1]);
}

function appendEncodingWarning(result: PrepareDisplayResult, declaredEncoding: string): void {
  const issue: VmlIssue = {
    code: HarErrorCode.UNSUPPORTED_HTML_ENCODING,
    severity: VmlIssueSeverity.WARNING,
    stage: 'rawfile',
    message: 'HTML 声明了不支持的字符编码。',
    cause: `无法识别字符编码 ${declaredEncoding}。`,
    action: '使用 UTF-8 降级解码；正文可能出现乱码。'
  };
  result.report.issues.push(issue);
  result.report.warnings.push(issue);
  if (result.report.status === VmlReportStatus.OK) {
    result.report.status = VmlReportStatus.PARTIAL;
  }
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
    let encoding: string | null = detectHtmlEncoding(content);
    let unsupportedEncoding: string = '';
    if (encoding === null) {
      const preview: string = util.TextDecoder.create('iso-8859-1')
        .decodeToString(content.subarray(0, Math.min(content.length, 4096)));
      const match: RegExpMatchArray | null = preview.match(/charset\s*=\s*["']?\s*([a-z0-9._-]+)/i);
      unsupportedEncoding = match === null ? 'unknown' : match[1];
      encoding = 'utf-8';
    }
    const decoder: util.TextDecoder = util.TextDecoder.create(encoding);
    const html: string = decoder.decodeToString(content);
    const result: PrepareDisplayResult = prepareDisplayHtml(html);
    if (unsupportedEncoding.length > 0) {
      appendEncodingWarning(result, unsupportedEncoding);
    }
    return result;
  } catch (error) {
    return createFileFailure(HarErrorCode.RAWFILE_READ_FAILED, String(error));
  }
}
