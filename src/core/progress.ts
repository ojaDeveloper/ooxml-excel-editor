/** 解析进度(分阶段)。read/build 有真实 ratio;parse 是 exceljs 黑盒,ratio 缺省=不确定态。 */
export interface ParseProgress {
  stage: 'read' | 'parse' | 'build'
  /** 0..1;parse 阶段为 undefined(不确定态,UI 走脉冲) */
  ratio?: number
}

export type ProgressFn = (p: ParseProgress) => void
