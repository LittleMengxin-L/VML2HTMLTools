/*
 * Copyright (c) 2026 Huawei Technologies Co., Ltd. All rights reserved.
 *
 * SPDX-License-Identifier: MIT
 * Licensed under the MIT License. See LICENSE in the HAR root for details.
 */

import {
  HarErrorCode,
  VmlIssue,
  VmlIssueSeverity,
  VmlReport,
  VmlReportStatus,
  VmlStats
} from './VmlTypes';

/** 创建各项计数均为零的解析统计对象。 */
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

/**
 * 创建包含单条错误的适配层失败报告。
 * @param code 适配层错误码。
 * @param message 错误描述，同时作为 cause 字段。
 * @param stage 错误发生阶段。
 * @param action 调用方可采取的处理建议。
 * @returns status 为 failed 的诊断报告。
 */
export function createHostFailure(
  code: HarErrorCode,
  message: string,
  stage: string,
  action: string
): VmlReport {
  const issue: VmlIssue = {
    code: code,
    severity: VmlIssueSeverity.ERROR,
    stage: stage,
    message: message,
    cause: message,
    action: action
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
