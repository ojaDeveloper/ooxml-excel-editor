# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与 [语义化版本](https://semver.org/lang/zh-CN/)。

## [1.2.0] - 未发布(开发中)

编辑 UX 补齐 + 性能 + 导出错误可见性(向后兼容)。

### 新增
- **富粘贴(从 Excel/WPS 复制整块)**:`Ctrl+V` 现在优先读剪贴板 **text/html**,完美解析**字体/颜色/填充/边框/对齐/合并单元格**(与 WPS 一致),整块**单次撤销**(整簿快照逆)。值优先取 Excel 的 `x:num`/`x:fmla`(原始值),否则取文本交类型推断。无 html 时回退原 TSV(`pasteText`,不变)。新 `clipboard-html.ts parseClipboardHtml`;`EditController.pasteRich`;API `pasteRichHtml(html,at?)`。
  - **图片(多通道,硬需求)**:① HTML 里的 data-uri `<img>` → 落格转内嵌图;② 剪贴板单张图片(`image/png` 等)→ 落活动格(`pasteImageBlob`);③ 拖图片文件进网格(消费方接 `pasteImageBlob`)。**已知边界**:Excel/WPS **区域复制的内嵌图一般进不了浏览器剪贴板**(只给 text/html+text/plain),所以"复制一整块带图、粘贴时图一起来"做不到 —— 这是浏览器固有限制,需用上述替代通道(单独复制一张图 / 拖文件)。
- **图片点击放大 + 下载原图**:网格里的图(WPS 内嵌图 DISPIMG / 浮动图)可点开看大图、下载原始字节。框架无关 `LightboxHost`(body 级暗背景 + 居中大图 + 「下载原图」+ 点背景/Esc/关闭按钮)。触发:**只读模式单击图**放大;**编辑模式右键**「查看大图 / 下载原图」(不抢选区/编辑)。新 prop `imageLightbox`(默认 true);新 API `openImageLightbox(src,fileName?)` / `getCellImageAt(row,col)`。顺带给点击加 3px 拖拽死区(微抖不再被当拖动,单击语义更稳)。修了 Vue 布尔 prop 缺省被判 false 的坑(`imageLightbox` 加进 withDefaults)。
- **虚拟空行/空列(滚动自动延伸,不动 dimension)**:滚到数据末尾下方仍有空行/空列可滚动、选中、编辑,像 Excel/WPS 的"无限网格";但**不写进 dimension/文件**(避免体积变大)——只有真去编辑某空格,它才靠 `growDimension` 变实。`GridMetrics` 加虚拟范围(`vRows/vCols/virtualWidth/Height`,封顶 Excel 上限 1048576×16384);`totalWidth/Height` 仍按 dimension(**导出/data-access 不含虚拟空行**);spacer 尺寸 / 可视区 / 命中夹取改用虚拟范围;控制器 `recomputeVirtualExtent()`(滚动/缩放/resize 时只增不减、按需延伸)。纯 core,双壳自动继承;新 API `getVirtualExtent()`。
- **背景色 / 字体色(回显 + 修改)**:新 API `getActiveFillColor()` / `getActiveFontColor()`(回显活动格当前色,#RRGGBB)+ `setSelectionFill(color|null)` / `setSelectionFontColor(color)`(改选区,入命令栈)。两壳 demo 加 WPS 风格的背景/字体取色器 + 清除填充。
- **修复内嵌图"灰底"**:DISPIMG 图加载中 / 缺登记项时,之前画 `#f2f4f7` 灰底盖住了单元格自身填充色(白),看着像默认灰。改为加载中不画底色(露出单元格白填充),仅缺图时画个淡图标、不盖色。
- **公式栏(Fx 内容条)可编辑 + 联动**:顶部公式栏从只读改为可编辑 `<input>`(editable + 该格非只读时)。在栏里输入提交(回车下移、Esc 取消、失焦提交)→ 改活动格;切选区 / 格内编辑 → 栏即时反映。栏显示**可编辑字符串**:公式 `=...`、数值原始数字串(非格式化,避免编辑货币/千分位被当文本)、布尔 TRUE/FALSE。新 API:`getCellEditString()` / `canEditActiveCell()` / `commitActiveCellValue(value, move?)`。仅值真变化才入命令栈。两壳同构(React 顺带补:`cell-change` 触发 chrome 重渲)。
- **合并/拆分单元格**:`mergeCells(range)`(吸收相交旧合并、清空被覆盖格只留左上锚点)/ `unmergeCells(range)`,入命令栈可撤销。
- **粘贴**:`Ctrl+V` / `pasteText(text, at?)` —— TSV(制表符+换行)→ 区域写入,类型自动推断、跳过只读、整块一次撤销。
- **右键上下文菜单**:框架无关 body 级 DOM 菜单 —— 插入/删除行列、合并/拆分、清除内容、复制/粘贴;点外部/Esc/滚动关闭、贴边翻转。只读仍用浏览器默认菜单。
- **WPS 单元格内嵌图(DISPIMG)**:
  - **展示**:解析 WPS 私有件 `xl/cellimages.xml` + 单元格 `=DISPIMG("id",1)` 公式,把图按行高列宽画进单元格内(随网格滚动/裁剪/冻结/缩放),普通工具打不开的 WPS 内嵌图现在能正常显示。新 `parser/cell-image-parser.ts`;模型加 `WorkbookModel.cellImages` + `CellModel.dispImgId`;canvas 渲染带图片解码缓存 + onload 重绘。
  - **贴合方式可配置** `cellImageFit`:`contain`(默认,等比缩放,与 WPS 渲染一致——WPS 打开导出文件时 DISPIMG 固定按 contain 显示)/ `fill`(拉伸铺满随格变形)/ `cover`(等比裁剪铺满),两壳 prop + `setCellImageFit()` 运行时切换、即时重绘。
  - **行 customHeight 保真**:解析 `<row customHeight="1">` 标记(ExcelJS 不暴露),自动行高跳过手动设高的行——避免长文本把"作者设矮放图"的行撑大,渲染/导出行高都与 WPS 一致。顺带放宽 fast-xml-parser 实体展开上限(大表几百个 `&quot;` 会撞默认 1000 上限致 drawing/row-meta 解析静默失败)。
  - **就近 / 批量嵌入**:`convertImageToCellAuto(imgIdx)`(图压在哪格就嵌哪格,几何反推)/ `convertAllImagesToCells(col?)`(整表或整列批量,一次进撤销栈)/ `convertCellImageToFloat(row,col,size?)`(嵌入→浮动);右键单格菜单为「将此处浮动图嵌入 / 整列嵌入(N 张)/ 整表嵌入(N 张)/ 内嵌图转浮动图」。两向入命令栈(整簿快照逆)、发 `cell-change`/`image-change`、翻脏标记。`getCellImages()` 读登记表。`convertImageToCell(imgIdx,row,col)` 保留(显式目标格)。
  - **导出往返**:ExcelJS 写出后在 zip 层回注 `cellimages.xml` + rels + media + `[Content_Types].xml`/`workbook.xml.rels` 补丁(新 `export/wps-cellimages.ts`,从模型重建),rebuild / overlay 两模式均覆盖。原有的 + App 内新转的内嵌图导出后用 WPS 打开都显示。验证:解析→导出→再解析 往返存活(单测)。
  - **逐字节对齐真·WPS(修正 #REF!)**:首版用了 `cellimages+xml`(复数)内容类型 + `2017/etCustomData` 关系类型 + 空 `<xdr:spPr/>`,导致 WPS 加载不了登记表、DISPIMG 显示 `#REF!`。据真·WPS 文件实测修正为 **单数** `cellimage+xml`、关系类型 **`2020/cellImage`**、`<xdr:spPr>` 补全 `xfrm`+`prstGeom rect`、`cNvPr` 加 `descr`、DISPIMG 格缓存值 `<v>` 写为 `=DISPIMG("id",1)`。单测锁定这些字段。

### 性能
- **undo 快照轻量化**:增删行列的整簿快照从 `structuredClone`(深拷图片字节/图表)改为手写轻量克隆 —— 只克隆编辑会动的部分,**图片字节/图表/形状/条件格式按引用共享**(编辑期间不可变),大文件 + undo 栈内存大幅下降。惠及结构编辑 undo、脏状态 baseline、还原原件。

### 修复
- **导出失败不再静默**:React 壳「导出 PDF」之前只把错误转给 `onError` prop,未接就被吞(无反应、控制台无报错)。两壳统一为 `console.error` + 上报(emit/onError)+ alert。**提醒**:PDF 导出需可选 peer `jspdf`(`npm i jspdf`);未装时现在会明确报错。

## [1.1.0] - 2026-06-05

把 1.0.0 编辑能力的三处已知 v1 限制做成增强(向后兼容)。

### 公式引用自动重写(增删行列)
- 增删行/列后,自动重写全簿公式里指向该表的 A1 引用(`=A5` 插一行→`=A6`;删被引用行→`#REF!`),
  含绝对/相对 `$`、跨表 `Sheet1!A5`/`'My Sheet'!A5`、区域(删除收缩、全删 `#REF!`);跳过字符串字面量
  与函数名。结构命令改为整簿快照(`cloneWorkbook`)→ 跨表重写也可撤销;开 `recalc` 时引擎按新公式重建。
- 新 `formula/refs.ts`:`shiftFormulaRefs` / `rewriteWorkbookFormulas`。

### 图片导出保真
- 区分锚型:有 `to` 的双格锚 → ExcelJS `br`(随单元格缩放);单格锚 → `tl`(含子格 EMU 偏移转分数列/行)
  + 像素 `ext`;`editAs` 跟随模型。不再一律导成 oneCell+ext、不再丢子格偏移。

### .xlsx 高保真 overlay 导出
- 新 `exportXlsx({ fidelity: 'overlay' })`:重载原始 .xlsx,只把编辑后的 值/样式/合并/行高列宽/冻结
  叠加上去,**保留** ExcelJS 能往返的其余部分(条件格式/数据验证/打印设置/定义名/图表 等)——默认 `rebuild`
  会丢这些。组件加载时留存原件字节供其使用,缺原件自动回退 `rebuild`。overlay 不反映 增删行列/图片 编辑。

### 测试
- 188 单测 + 40 e2e + build 全绿;core 仍零 vue/react/hyperformula/exceljs 静态 import。

## [1.0.0] - 2026-06-05

**只读 → 可编辑** —— 在 0.2.0(只读双壳)基础上,把组件升级成**可选编辑器**(默认仍只读、零回归)。
所有编辑能力建在**一份可变内存模型**上,读 / 写 / 事件 / 导出共用同一层;core 始终框架无关、两壳同构。

### 编辑能力(props `editable` 开启)
- **单元格编辑**:双击 / F2 / 打字进编辑,内置文本编辑器 + `editor` 钩子自定义(下拉/日期/图片选择器,返回任意 DOM,框架无关)。
- **命令栈**:撤销/重做(Ctrl+Z/Y),所有写操作(值/样式/宽高/图片/结构)统一进栈。
- **样式编辑**:`setStyle(range, patch)`(font/fill/borders/对齐/数字格式)。
- **列宽行高**:拖拽 + `setColumnWidth/setRowHeight`,入命令栈。
- **图片**:浮动图拖拽移动 + `addImage/removeImage/moveImage/resizeImage`。
- **行列结构**:`insertRows/deleteRows/insertCols/deleteCols`(重键 cells、移合并/宽高/图片,快照逆撤销)。
- **公式重算**:`recalc` 开启 → 依赖格自动级联;默认 HyperFormula(可选 peer,GPL/商业双授权),`formulaEngine` 可换自研/持牌引擎(`FormulaEngine` 接口)。
- **脏状态**:`isDirty()` + `dirty-change` 事件 + `resetToOriginal()`(放弃修改还原原件)。

### 事件 + 查询(底层机制)
- **`cell-change` 携前后完整快照**(`CellSnapshot`:底层 CellModel + 解析 style + raw/computed/text),与 `getCellSnapshot` 查询 API 同一份结构。
- `dim-change` / `image-change` / `struct-change` / `dirty-change` / `edit-start` / `edit-commit`。

### 导出(一份数据层,所见即所得)
- **`.xlsx`**(`exportXlsx/downloadXlsx`):从编辑后模型重建(需 `exceljs`),覆盖值/公式/样式/合并/宽高/冻结/图片。
- **`.json`**(raw 类型值)/ **`.csv`**(格式化显示值,UTF-8 BOM):复用 `getSheetData/getWorkbookJSON`,三格式天然一致。

### 限制
- 增删行列**不自动重写公式引用文本**(缓存值随格移动);写回 .xlsx 丢 VBA/工作表保护/复杂 DrawingML;图片导出忽略部分子格偏移。

### 测试
- 175 单测 + 38 真浏览器 e2e(Vue + React 双覆盖)+ build 全绿;`dist/core.js` 零 vue/react/hyperformula/exceljs 静态 import(重依赖全动态懒加载)。

## [0.2.0] - 2026-06-05

**只读** —— 在 0.1.0(Vue-only v1)基础上加:**Vue + React 双壳共享框架无关 core**、列排序、
跨框架插件、边框还原、多入口分包(core/vue/react)、完整文档。后续「编辑」能力按 semver 推进(0.3.0+)。

### 交互 / 插件
- **列排序**:自动筛选下拉加「升序/降序」,按列重排数据区(整行移动),合并区相交则拒绝。
- **插件跨框架**:`overlay` 钩子从返回 Vue VNode 改为返回 **DOM 节点**,`core/plugin` 不再 import vue(core 彻底框架无关)。**同一份 `definePlugin` 在 Vue 与 React 壳通用**;React 壳新增 `plugins` prop,支持 theme/transformModel/cellStyle/events/overlay/toolbar/setup 全套。

### 边框还原(对齐 Excel/WPS)
- **合并单元格内部不再画网格线**(之前无填充的合并格会透出内部浅灰网格线)。
- **斜线边框(对角线 ↘/↗)**:parser 解析 `diagonal{up,down,style,color}`,canvas 与矢量 PDF 都绘制。
- **相邻共享边按权重取较重者**(hair<…<medium<thick<double):普通格的边框绘制顺序无关、与 Excel/WPS 一致(合并区仍画自身四周)。

### 新增
- **React 壳**:`ooxml-excel-editor/react` 导出 `<ExcelViewer>`(`forwardRef` + 命令式 `ExcelViewerHandle`)与 `useExcelDocument`,与 Vue 壳**共用 ~100% core 引擎**。
- **框架无关 core 入口**:`ooxml-excel-editor/core` 暴露引擎(`ViewerController` / `CanvasRenderer` / `WorkbookExporter` / `OverlayManager` / `PluginOverlayHost`)+ 解析 + 读数据 + 类型,零框架依赖。
- **多入口构建**:产物拆为 `dist/core.js`(引擎)+ `dist/index.js`(Vue 壳)+ `dist/react.js`(React 壳),后两者共享同一份 core 引擎 chunk;各自 `.d.ts`。
- React demo(`/react.html`)+ React 真浏览器 e2e(渲染 / 选区 / 查找 / 数据 API / 导出 / 插件)。

### 变更(重构,行为零回归)
- 把 `ExcelViewer.vue` 的非框架编排逐步下沉到框架无关 `src/core/viewer/`:
  - `OverlayManager`(图片/图表/形状叠加层)+ `PluginOverlayHost`(插件 overlay DOM)
  - `ViewerController`:渲染引擎、选区 + 鼠标/键盘交互、查找、自动筛选、排序、导出编排桥接
  - `WorkbookExporter`(`src/core/export/exporter.ts`):导出/打印编排,靠 `ExporterHost` 与壳解耦
- `ExcelViewer.vue` 收薄为薄壳:props/插件桥接 + chrome + 经 hooks 桥接控制器响应式。
- `package.json`:`exports` 增 `./react`、`./core`;`vue`/`react`/`react-dom` 改为可选 peer(按框架二选一)。

### 修复
- React 壳:控制器创建与 rebuild 改用 `useLayoutEffect`,避免晚到的 passive rebuild 清掉刚设置的交互态。
- 库构建:`worker-client` stub 别名兼容 React 壳的 `@/composables/worker-client` 引入,避免误把 1.4MB exceljs 打进产物。

### 基线
- 测试:**111 单测 + 16 e2e(Vue + React)**全绿;`dist/core.js` 无 vue/react import;exceljs 仅运行时 `import()` 不打包。

## [0.1.0] - 2026-06-03

高保真 .xlsx 预览组件 v1:从零实现解析 + Canvas 渲染(Vue,只读)。冻结窗格 / 合并单元格 /
条件格式 / 图表 / 图片形状 / 数字日期格式 / 超链接批注 / 查找 / 自动筛选 / 可配置可插件工具栏 /
导出(图片·位图PDF·矢量PDF·打印) / 读数据 API / 主题·钩子·插件扩展点。
