export enum VmlIssueSeverity {
  ERROR = 'error',
  WARNING = 'warning',
  INFO = 'info'
}

export enum VmlReportStatus {
  OK = 'ok',
  PARTIAL = 'partial',
  FAILED = 'failed'
}

/** HAR/ArkWeb 适配层错误码；页面内解析错误码由运行时 ERROR_CODES 提供。 */
export enum HarErrorCode {
  RUNTIME_INJECTION_FAILED = 'H7002',
  RUNTIME_VERSION_MISMATCH = 'H7005',
  INVALID_HTML = 'H7006',
  OUTPUT_FILE_WRITE_FAILED = 'H7011',
  PAGE_READY_TIMEOUT = 'H7012',
  HTML_TRANSFORM_FAILED = 'H7013',
  FILE_READ_FAILED = 'H7014'
}

export interface VmlStats {
  shapes: number;
  images: number;
  matched: number;
  transformed: number;
  skipped: number;
  sourcesRewritten: number;
  paragraphsConverted: number;
}

export interface VmlIssue {
  code: string;
  severity: string;
  stage: string;
  message: string;
  cause: string;
  action: string;
  shapeIds?: string[];
  originalError?: string;
}

export interface VmlReport {
  status: string;
  stats: VmlStats;
  issues: VmlIssue[];
  errors: VmlIssue[];
  warnings: VmlIssue[];
  infos: VmlIssue[];
}

export interface VmlParserOptions {
  processAllWrapTypes?: boolean;
  removeEmptyPlaceholderTables?: boolean;
  skipDirectDivImages?: boolean;
  markAttribute?: string;
  maxCoordinatePt?: number;
  debug?: boolean;
}

export interface ParseHtmlResult {
  html: string | null;
  report: VmlReport;
}

/** VmlHostController.parseFile 的文件级解析结果。outputPath 为 null 表示解析或落盘失败。 */
export interface VmlFileParseResult {
  outputPath: string | null;
  report: VmlReport;
}
