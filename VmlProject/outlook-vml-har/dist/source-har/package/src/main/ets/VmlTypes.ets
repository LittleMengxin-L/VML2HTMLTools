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
  INVALID_SELECTOR = 'H7001',
  RUNTIME_INJECTION_FAILED = 'H7002',
  SCRIPT_EXECUTION_FAILED = 'H7003',
  INVALID_SCRIPT_RESULT = 'H7004',
  RUNTIME_VERSION_MISMATCH = 'H7005',
  INVALID_HTML = 'H7006',
  RAWFILE_READ_FAILED = 'H7007',
  INVALID_RAWFILE_PATH = 'H7008'
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

export interface ChangeVmlResult {
  html: string | null;
  report: VmlReport;
}

export interface PrepareDisplayResult {
  html: string;
  report: VmlReport;
}
