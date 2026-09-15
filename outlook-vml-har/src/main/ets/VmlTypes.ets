/*
 * Copyright (c) 2026 Huawei Technologies Co., Ltd. All rights reserved.
 *
 * SPDX-License-Identifier: MIT
 * Licensed under the MIT License. See LICENSE in the HAR root for details.
 */

/** 诊断信息的严重程度。 */
export enum VmlIssueSeverity {
  ERROR = 'error',
  WARNING = 'warning',
  INFO = 'info'
}

/** 解析报告状态：成功、存在警告或失败。 */
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

/** 当前解析调用的处理统计。 */
export interface VmlStats {
  /** 通过环绕规则筛选的 shape 数量。 */
  shapes: number;
  /** 带有 v:shapes 属性的普通图片数量。 */
  images: number;
  /** 生成了布局应用计划的图片数量。 */
  matched: number;
  /** 成功应用布局的图片数量。 */
  transformed: number;
  /** 跳过布局或应用失败的图片数量。 */
  skipped: number;
  /** changeVMLSrcs 改写的 VML shape 块数量。 */
  sourcesRewritten: number;
  /** 转换为 div 的 VML 段落数量。 */
  paragraphsConverted: number;
}

/** 单条诊断信息，调用方可使用 code 判断错误类型。 */
export interface VmlIssue {
  /** 错误码或提示码。 */
  code: string;
  /** 严重程度：error、warning 或 info。 */
  severity: string;
  /** 产生诊断信息的处理阶段。 */
  stage: string;
  /** 诊断信息摘要。 */
  message: string;
  /** 诊断原因。 */
  cause: string;
  /** 处理行为说明或调用建议。 */
  action: string;
  /** 相关 shape ID；无法关联时省略。 */
  shapeIds?: string[];
  /** 底层异常的文本描述。 */
  originalError?: string;
}

/** 解析报告，汇总处理统计与按严重程度分类的诊断信息。 */
export interface VmlReport {
  /** 汇总状态：ok、partial 或 failed。 */
  status: string;
  /** 本次处理统计。 */
  stats: VmlStats;
  /** 全部诊断信息。 */
  issues: VmlIssue[];
  /** 错误级诊断信息。 */
  errors: VmlIssue[];
  /** 警告级诊断信息。 */
  warnings: VmlIssue[];
  /** 提示级诊断信息。 */
  infos: VmlIssue[];
}

/** HTML 字符串转换结果，须结合 report.status 判断转换状态。 */
export interface ParseHtmlResult {
  /** 转换后的 HTML；适配层转换失败时为 null。 */
  html: string | null;
  /** 本次转换的诊断报告。 */
  report: VmlReport;
}

/** VmlHostController.parseFile 的文件级解析结果。outputPath 为 null 表示解析或落盘失败。 */
export interface VmlFileParseResult {
  /** 成功写入的文件路径；解析或写入失败时为 null。 */
  outputPath: string | null;
  /** 本次转换的诊断报告。 */
  report: VmlReport;
}
