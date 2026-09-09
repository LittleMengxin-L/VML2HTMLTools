import { util } from '@kit.ArkTS';

/** 字节内容解码结果。encoding 为实际采用的解码编码；unsupportedEncoding 非空表示声明无法识别、已按 UTF-8 降级。 */
export interface HtmlDecodeResult {
  html: string;
  encoding: string;
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

function extractDeclaredCharset(content: Uint8Array): string | null {
  // charset 标记本身由 ASCII 字符组成，用 iso-8859-1 扫描头部不会破坏字节。
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
 * 把 HTML 文本内声明的字符编码统一为 UTF-8。
 * 文本已被解码为 Unicode 且落盘/传输统一使用 UTF-8 字节时，声明必须与字节一致，
 * 否则下游按旧声明（如 gb2312）解读 UTF-8 字节会乱码。
 */
export function normalizeCharsetDeclaration(html: string): string {
  return html.replace(/charset\s*=\s*["']?\s*[a-z0-9._-]+/gi, 'charset=utf-8');
}
