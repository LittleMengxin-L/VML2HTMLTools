import {
  HarErrorCode,
  VmlIssue,
  VmlIssueSeverity,
  VmlReport,
  VmlReportStatus,
  VmlStats
} from './VmlTypes';

/** 空的解析统计对象。所有「失败报告」与「空成功报告」共用同一结构。 */
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
 * 构造一个 FAILED 状态的适配层错误报告。
 * @param stage 出错层次（如 'vml-host' / 'harmony-adapter'），用于定位问题来源。
 * @param action 面向调用方的修复建议，随出错场景不同而不同。
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
