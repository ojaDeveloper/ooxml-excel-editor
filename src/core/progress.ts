/** 解析进度(分阶段)。read/build 有真实 ratio;parse 是 exceljs 黑盒,ratio 缺省=不确定态。 */
export interface ParseProgress {
  stage: 'read' | 'parse' | 'build'
  /** 0..1;parse 阶段为 undefined(不确定态,UI 走脉冲) */
  ratio?: number
}

export type ProgressFn = (p: ParseProgress) => void

/**
 * 导出 / 批量转换进度。所有耗时操作(PDF / PNG / XLSX / 批量浮动↔嵌入)统一报这套。
 * stage 语义:
 *  - 'render':canvas 渲染一张表(大表分块时按 ratio 报子进度)
 *  - 'compose':多表合成一张大图 / 合并几何
 *  - 'paginate':PDF 分页 / 每页布局
 *  - 'write':PDF / 图片 blob 编码写出
 *  - 'zip':XLSX zip 压缩(exceljs 黑盒,只前/后两次)
 *  - 'convert':批量图片浮动↔嵌入(P2)
 */
export type ExportStage = 'render' | 'compose' | 'paginate' | 'write' | 'zip' | 'convert'
export interface ExportProgress {
  stage: ExportStage
  /** 当前正在处理的表(可选) */
  sheetIndex?: number
  /** PDF 分页阶段:当前页 */
  pageIndex?: number
  /** 0..1;黑盒阶段省略 */
  ratio?: number
  /** 给 UI 显示的标签(可选,如"渲染 Sheet1 (3/5)") */
  label?: string
}
export type ExportProgressFn = (p: ExportProgress) => void
