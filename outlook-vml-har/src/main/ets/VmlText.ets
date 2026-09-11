import { util } from '@kit.ArkTS';

/** HTML 字节解码结果，包含文本、实际编码及未识别的编码声明。 */
export interface HtmlDecodeResult {
  /** 解码后的 HTML 文本。 */
  html: string;
  /** 实际使用的字符编码。 */
  encoding: string;
  /** 未识别的编码标签；空字符串表示未发生此类降级。 */
  unsupportedEncoding: string;
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
    // 将 GBK/GB2312 编码标签统一映射为 GB18030。
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

function extractDeclaredCharset(content: Uint8Array): string | null {
  // 将前 4096 字节映射为单字节字符，查找 charset 声明。
  const previewLength: number = Math.min(content.length, 4096);
  const preview: string = util.TextDecoder.create('iso-8859-1')
    .decodeToString(content.subarray(0, previewLength));
  const charsetMatch: RegExpMatchArray | null = preview.match(
    /charset\s*=\s*["']?\s*([a-z0-9._-]+)/i
  );
  return charsetMatch === null ? null : charsetMatch[1];
}

/**
 * 解码 HTML 字节内容：优先 BOM，其次 <meta charset> 声明（gb2312/gbk 归一到
 * gb18030），无声明时默认 UTF-8；声明无法识别时按 UTF-8 降级并返回 unsupportedEncoding。
 * @param content HTML 文件字节。
 * @returns 解码后的文本及编码信息。
 */
export function decodeHtmlContent(content: Uint8Array): HtmlDecodeResult {
  const bomEncoding: string | null = detectBomEncoding(content);
  if (bomEncoding !== null) {
    const decoder: util.TextDecoder = util.TextDecoder.create(bomEncoding);
    return {
      html: decoder.decodeToString(content),
      encoding: bomEncoding,
      unsupportedEncoding: ''
    };
  }

  const declared: string | null = extractDeclaredCharset(content);
  if (declared === null) {
    const decoder: util.TextDecoder = util.TextDecoder.create('utf-8');
    return {
      html: decoder.decodeToString(content),
      encoding: 'utf-8',
      unsupportedEncoding: ''
    };
  }

  const normalized: string | null = normalizeHtmlEncoding(declared);
  if (normalized === null) {
    const decoder: util.TextDecoder = util.TextDecoder.create('utf-8');
    return {
      html: decoder.decodeToString(content),
      encoding: 'utf-8',
      unsupportedEncoding: declared
    };
  }
  const decoder: util.TextDecoder = util.TextDecoder.create(normalized);
  return {
    html: decoder.decodeToString(content),
    encoding: normalized,
    unsupportedEncoding: ''
  };
}

/**
 * 将文本中匹配到的 charset 声明替换为 UTF-8。
 * @param html 已解码、准备以 UTF-8 写入的 HTML 文本。
 * @returns 替换后的文本。
 */
export function normalizeCharsetDeclaration(html: string): string {
  return html.replace(/charset\s*=\s*["']?\s*[a-z0-9._-]+/gi, 'charset=utf-8');
}
