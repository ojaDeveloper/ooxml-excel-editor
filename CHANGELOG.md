# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与 [语义化版本](https://semver.org/lang/zh-CN/)。

## [1.14.1] - 2026-06-15

> 文档审计 + 入口出口修正(无运行时行为变更)。一次"保证接入/二开都没问题"的体检。

### 修复 — 四入口出口同源(接入)

- **入口出口不一致 bug**:公式引擎工厂(`builtinFormulaEngineFactory` / `hyperFormulaEngineFactory` / `FUNCTION_NAMES` / `FormulaEngine` 类型)此前只在 `/core` 出口,**主入口 / `/react` / `/vue2` 都拿不到** —— 但 `:formula-engine` 是主组件上的 prop,文档让注入 `hyperFormulaEngineFactory`,实际 `import { hyperFormulaEngineFactory } from 'ooxml-excel-editor'` 会失败。`/react` 入口更是只导出组件 + hook,`parseWorkbook`/类型/`definePlugin` 全够不着。
- **修法**:主 / `/react` / `/vue2` 入口统一 `export * from core`(各自再加自己的组件)→ **四入口同源**,任一入口都拿到完整 core 公共 API(解析/读数据/类型/插件/导出/公式引擎工厂…),不再各维护清单致漂移。`CellStyleCtx` / `DataValidationRule` 补进 core 出口。typecheck + build + 419 单测验证。

### 文档 — 使用 vs 二开 分离 + 准确性

- 抽出 **`EXTENDING.md`(二开 / 扩展 API 手册)**:主题 `:theme` / 数据·渲染钩子 / 自定义编辑器 / 右键菜单 transform / 工具栏自定义 / 分层 UI slots / 命令式 API / 导出·打印高级选项 / 插件 `definePlugin`。README 瘦身成**调用方**文档(装/用/props/编辑/导出),顶部留指路;深度二开看 ARCHITECTURE.md。
- README **具名导出表**重写(原列 5 项、实际 ~40 项 → 按 解析/读数据/格式/插件/公式引擎/导出/类型 分类 + 标注"四入口同源");**props 表**补 `toolbar` / `plugins` / `openLinks` + 数据验证说明;修跨章节断锚点。

## [1.14.0] - 2026-06-15

> 新增**内置公式引擎**(MIT,零依赖)设为 recalc 默认引擎(取代 GPL 的 HyperFormula),+ **公式自动补全**。覆盖日常 ~60 个常用函数;需更全覆盖可注入 HyperFormula 或自研引擎。

### 新增 — 内置 MIT 公式引擎(默认)

- **从零实现**(`formula/builtin/`):词法 + 表达式解析(运算符优先级 / `A1` 绝对相对 / 区域 `A1:B2` / 跨表 `Sheet1!A1` / 函数 / 一元 ± / `%`)→ AST;求值器(错误就近传播 `#DIV/0! #N/A #NAME? #NUM! #REF! #VALUE!`);**依赖图 + 拓扑级联重算 + 循环引用检测**(环上格 → `#REF!`)。
- **函数 ~60**:聚合(SUM/AVERAGE/MAX/MIN/COUNT/COUNTA/PRODUCT/SUMPRODUCT)、数学(ROUND/ROUNDUP/ROUNDDOWN/ABS/INT/MOD/SQRT/POWER/CEILING/FLOOR…)、逻辑(IF/IFERROR/IFNA/AND/OR/NOT/XOR/IFS)、文本(LEFT/RIGHT/MID/LEN/CONCAT/UPPER/LOWER/TRIM/SUBSTITUTE/FIND/SEARCH…)、查找(VLOOKUP/HLOOKUP/INDEX/MATCH/CHOOSE)、条件聚合(SUMIF/COUNTIF/AVERAGEIF)、信息(ISNUMBER/ISBLANK/ISERROR…)、日期(TODAY/NOW/DATE/YEAR/MONTH/DAY/WEEKDAY/DAYS)。
- **★ 默认引擎变更**:recalc(`:recalc`,opt-in)的默认引擎由 **HyperFormula → 内置引擎**。好处:**MIT 无 GPL 顾虑、零额外依赖、不再懒加载 ~400KB**。代价:函数集比 HyperFormula 的 ~395 小(日常足够)。
- **HyperFormula 仍可用**:注入 `:formula-engine="hyperFormulaEngineFactory"`(或自研)获得更全覆盖。`core` 新导出 `builtinFormulaEngineFactory` / `BuiltinFormulaEngine` / `FUNCTION_NAMES` / `hyperFormulaEngineFactory`;`defaultFormulaEngineFactory` 保留为 HyperFormula 别名(向后兼容显式注入它的旧代码)。
- **解析/显示不受影响**:打开文件仍显示原件缓存的计算结果(与引擎无关);引擎只在编辑后重算时介入。
### 新增 — 公式自动补全

- 在框架无关默认单元格编辑器(`edit/formula-autocomplete.ts`,三壳自动都有)里:输 `=SU` 时下方弹**函数名列表 + 参数提示**(`FUNCTION_SIGNATURES`);↑↓ 选、Enter/Tab 接受(插入 `NAME(` 并把光标移进括号)、Esc 关、点选即填。只在公式(`=` 开头)且光标处于函数名 token 时弹,不影响普通文本编辑。列表来源 = 引擎实际支持的函数(`FUNCTION_NAMES`),所见即所得。
- 测试:`formula/builtin/__tests__`(parse 12 + eval 11 + engine 7 = 30 例:运算符/函数/级联/循环/跨表);`e2e/formula-autocomplete.e2e.ts`(=SU→弹 SUM→点选插入,三壳);现有 `edit-formula.e2e.ts` 重算 e2e 改由内置引擎驱动仍全过。基线:**419 单测 + 192 e2e**。

## [1.13.0] - 2026-06-15

> 新增 **不连续多区域选择**(Ctrl/⌘ + 点击)。选区模型从单矩形扩成多矩形;纯框架无关 core 交互,壳只转发鼠标。

### 新增 — Ctrl 多区域选择

- **Ctrl/⌘ + 点击** 行头 / 列头 / 单元格 → 把当前选区收进多选集再起新区,**加选不相邻**区域(Shift 仍是连续区间,已有);普通点击 / 键盘导航 / `selectCell` / 全选回到单选。
- 选区模型加 `selRanges[]` + `getSelectionRanges()`(全部矩形,末个为活动区)/ `hasMultiSelection()`;`canvas-renderer` 加 `setExtraSelection` 画所有附加区(填充 + 边框;多选时不画自动填充柄,对齐 Excel)。
- **复制**:多选时各区按出现顺序**逐行堆叠**成块 → TSV + HTML 表写剪贴板(覆盖最常见的"Ctrl 点多个行头复制非相邻行",粘到 Excel/WPS / app 内都成堆叠块)。
- **状态栏统计**跨所有选区聚合(count/sum/avg/min/max);新增控制器 `getSelectionStats()`,三壳状态栏改用它。
- 三壳句柄 + 插件 `ViewerApi` 暴露 `getSelectionRanges` / `hasMultiSelection`。测试:`e2e/multi-select.e2e.ts`(Ctrl+点击两不邻行头 → 2 区域 + 回单选,Vue/React/Vue2)。基线:**389 单测 + 189 e2e**。

## [1.12.0] - 2026-06-15

> 新增 **格式刷**(Format Painter)。纯框架无关 core 交互(控制器采样 + onMouseUp 刷),壳只加工具栏按钮;需 `editable`。

### 新增 — 格式刷

- 工具栏 `format-painter` 入口:先选**源格**点按钮采样其完整样式(字体/填充/边框/对齐/换行/数字格式),再**点或拖**目标格/区域即刷上(单次撤销);`Esc` 或再点按钮退出,待刷时光标变 `copy`。
- 控制器 `startFormatPainter(sticky?)` / `isFormatPainterArmed()` / `cancelFormatPainter()`;刷动作在 `onMouseUp` 选区完成后应用(复用 setStyle)。三壳句柄 + 插件 `ViewerApi` 暴露;三 demo 工具栏加 `format-painter` 入口(工具栏按钮 active 态反映待刷)。
- 测试:`e2e/format-painter.e2e.ts`(采样红底 → 刷到目标格 + undo,Vue/React/Vue2 三壳)。
- 顺手:`toolbar` 溢出相关 e2e 宽屏断言 1280→1680(工具栏又加了按钮,宽屏才全部容纳)。基线:**389 单测 + 186 e2e**。

## [1.11.0] - 2026-06-15

> 三个编辑小件合并:**查找替换补全 + 数字格式编辑器 + 批注编辑**。都复用已有引擎/对话框套路;对话框均为框架无关 DOM(三壳共用一份,UI 天然 1:1)。

### 新增 — 查找替换

- 此前只有「查找」(高亮定位),现补「替换」:查找栏开 `editable` 时多出替换行(替换输入 + 替换 / 全部替换);控制器 `setFindReplace` / `replaceCurrent`(替换当前并查找下一个,替换后重算命中)/ `replaceAll`(全部替换,**整体单次撤销**)。支持区分大小写 / 全字匹配;跳过只读格。三壳查找栏都加替换行(Vue3 `FindBar.vue` + React/Vue2 内联)。

### 新增 — 数字格式编辑器

- 框架无关对话框 `viewer/number-format-dialog-host.ts`(工具栏 `number-format` 入口):分类(常规/数值/货币/百分比/日期/时间/文本/自定义)+ 选项(小数位数 / 千分位 / 负数红色 / 货币符号 / 日期时间预设)→ **实时预览**(复用 number-format 引擎)+ 可直接编辑「格式代码」→ 确定即 `setStyle({ numFmt })`(单次撤销)。控制器 `setSelectionNumberFormat(code)` / `openNumberFormatDialog()`。

### 新增 — 批注编辑

- 批注此前只读显示,现可**新建/编辑/删除**:新 `set-comment` 命令(逆 = restore-cells 精确还原,单次撤销)+ `model/mutations.ts` 的 `setCellComment`(空批注清除、空格挂批注/清空);右键菜单单格加「插入/编辑/删除批注」;框架无关对话框 `viewer/comment-dialog-host.ts`(多行文本 + 确定/删除/取消)。控制器 `getCellComment` / `setCellComment` / `openCommentEditor`。**导出回写**:`xlsx-writer` rebuild + overlay 都把批注写成 ExcelJS note(此前 rebuild 丢批注)。

- 三壳句柄 + 插件 `ViewerApi` 都暴露上述 API;三 demo 工具栏加 `number-format` 入口。测试:`find-replace` / `number-format` / `comment` e2e(各 Vue/React/Vue2)+ `autofill`/`comment` 等单测。基线:**389 单测 + 183 e2e**。

## [1.10.0] - 2026-06-15

> 新增 **自动填充柄**(Excel/WPS 拖拽填充)。纯框架无关 core canvas 交互(渲染层画柄 + 控制器拖拽),三壳零改动自动获得;需 `editable`。

### 新增 — 自动填充柄(拖拽填充序列)

- **序列引擎**(`edit/autofill.ts`,纯函数可单测):接续源值产出新值 —— 全数值(1 个复制 / ≥2 个等差外推)、全日期(每格 +1 天 / 按相邻差)、"前缀+末尾整数"文本(`Item 1`→`Item 2`,保留前导零位宽)、星期/月份名(中英常见写法循环接续)、其它循环复制。
- **渲染层**(`canvas-renderer`):选区右下角画**填充柄**小方块(`editable` 才画,随 editable 实时显隐)+ `fillHandleAt` 命中检测 + 拖拽时画目标范围**虚线预览**。
- **控制器**:`onMouseDown` 命中填充柄 → `fill` 拖拽模式;`onMouseMove` 按鼠标位置算目标范围(主轴 = 偏移更大的方向,支持上/下/左/右)+ 设预览;`onMouseUp` 按源选区逐列/行接续序列,**整体单次撤销**(`set-cells` 命令),填充后选区扩到整片。填充柄上悬停显示十字光标。跳过只读格。
- **Ctrl/⌘ 修饰键翻转"复制 ↔ 序列"**(对齐 Excel):普通拖 单个数字→复制、≥2 数字→等差;按住 Ctrl 拖 单个数字→递增(+1)、序列→复制;日期/星期月份/文本递增普通→序列、Ctrl→复制。
- 注:v1 填充**值**(不复制源格式);列/行各自独立成序列。
- 测试:`edit/__tests__/autofill.test.ts`(13 例:复制/等差/日期/文本递增/星期/月份/循环 + Ctrl 翻转)+ `e2e/autofill.e2e.ts`(拖填充柄出等差序列 + undo + Ctrl 单数字递增,Vue/React/Vue2 三壳)。基线:**386 单测 + 171 e2e**。

## [1.9.0] - 2026-06-14

> 路线图「保真/编辑完整化」第二阶段:**条件格式从只读渲染 → 可编辑**。整个功能由 `conditionalFormat` prop 开启(三壳同名,默认 false = 关闭、与历史一致;三 demo 已开启)。支持全 6 类规则的新建/编辑/删除 + 导出回写,overlay 模式保留原件未编辑规则原样。

### 新增 — 条件格式可编辑(全类型)

- **模型 + 解析**:`ConditionalRule` 加 `id` / `origin`('parsed'|'user')/ `dirty` / `raw`(原始 ExcelJS rule,导出原样回写用)+ `top10`(rank/percent/bottom)/ `iconSet.reverse`。`parseConditional` 给每条规则派 id、存 raw、补全 top10/iconSet 字段。
- **命令栈**:新增 `set-conditional` 命令(整张 `conditional` 数组不可变替换,逆=换回前态)→ 新建/编辑/删除规则**整体单次撤销**。
- **编程 API**(控制器 + 插件 viewer + 三壳句柄):`getConditionalRules` / `addConditionalRule` / `updateConditionalRule` / `removeConditionalRule` / `setConditionalRules` / `openConditionalFormatDialog`。改完即 live 重渲。需 `conditionalFormat` + `editable`。
- **管理对话框**(框架无关 DOM `viewer/conditional-format-dialog-host.ts`,三壳共用一份 → UI 天然 1:1):列出当前表所有规则(可删/可编辑)+ 新建。6 类编辑器:突出显示单元格(cellIs:大于/小于/介于/等于… + 填充/字体色/加粗)、公式(expression)、色阶(colorScale 2/3 色)、数据条(dataBar 颜色 + 渐变)、图标集(iconSet 7 种 + 反向)、项目选取(top10 前/后 N + 百分比)。新建默认套到当前选区。
- **工具栏入口** `conditional-format`(三壳 + 三 demo,`conditionalFormat` 关时不渲染)。
- **导出回写**(rebuild + overlay 共用 `xlsx-writer` 的 `writeConditionalFormatting`):清空 ExcelJS 现有 CF 后按模型重建 —— **未编辑的 parsed 规则用 `raw` 原样回写**(含 cfvo 阈值,零退化);用户新建/编辑过的按模型 `buildExcelCfRule` 构造(全 6 类)。**1.9.0 起 rebuild 也回写条件格式**(此前 rebuild 丢弃);overlay 满足"原件规则原样留、只增改用户改的"。
- 测试:`edit/__tests__/conditional-format.test.ts`(解析保真 / 命令撤销 / rebuild + overlay 往返 5 例)+ `e2e/conditional-format.e2e.ts`(API 新增+撤销、对话框新建 cellIs,Vue + React + Vue 2 三壳)。基线:**373 单测 + 165 e2e**。

## [1.8.0] - 2026-06-14

> 路线图「保真/编辑完整化」第一阶段:① 把 1.7.0 起步的数据验证**做完整**(从"只能选值"到"编辑拦截非法输入 + 输入/出错提示");② 补上 **Vue 2 壳的 e2e 回归网**(此前 Vue 2 零 e2e,改 Vue 2 全靠手测,是 CLAUDE.md 点名的空洞)。

### 新增 — 数据验证完整化:编辑时拦截非法输入 + WPS 式提示

- **解析全类型规则**(`SheetModel.dataValidationRules`):list / 整数(whole)/ 小数(decimal)/ 日期(date)/ 时间(time)/ 文本长度(textLength)/ 自定义(custom),连同 operator(between/greaterThan/…)、约束操作数、`allowBlank`、出错信息(errorStyle/errorTitle/error)、输入提示(promptTitle/prompt)。1.7.0 的 `dataValidations`(下拉箭头区域)/`dataValidationLists`(选项)从这里派生,**零回归**。
- **编辑时拦截**(框架无关引擎 `edit/data-validation.ts`,纯函数可单测):内置编辑器提交 / 公式栏提交时校验。`errorStyle='stop'`(默认)→ **硬拒,不写入**,弹模态出错提示,**编辑器保持打开让用户改正**;`warning`/`information` → toast 软提示但放行;`custom` 公式与以 `=` 开头的公式不拦(结果未知);空值放行(允许清空)。
- **提示 UI**(框架无关 DOM,三壳共用 `viewer/validation-prompt-host.ts`):出错模态 / toast + 选中带"输入提示"的格时格旁弹黄色气泡(随选区/滚动跟手)。
- **顺手修一个编辑器 UX bug**:校验拒绝后,内置编辑器的"已提交"锁(`done`)曾被卡死 → 用户改正后回车无反应。现 `commit()` 返回成功与否,拒绝时解锁并记住被拒值(避免点弹窗按钮的 blur 二次触发叠弹),改内容即可再次提交;弹窗关闭后焦点还给编辑器。
- 测试:`edit/__tests__/data-validation.test.ts`(12 例:各类型/operator/空值/坏约束/软提示/自定义信息)+ `date-locale.test.ts` 加全类型规则解析;`data-validation.e2e.ts` 加"整数 1-100 校验:输入 999 → 弹拒、值不变、改 80 → 写入"(Vue + React + **Vue 2** 三壳)。

### 新增 — Vue 2 壳 e2e 回归网

- 此前 Vue 2 壳**完全没有 e2e**(只有 Vue 3 + React 双覆盖),改 Vue 2 只能手测 → 高风险点(patch 复用 controller 持有的 DOM 致 stale、Vue 2.6 函数 ref 被忽略)迟早回归。现补上:
  - `playwright.config.ts` 加**第二个 dev server**(端口 5302,`npm run dev:vue2`,plugin-vue2 SFC 编译器隔离);`vue2-demo` DEV 下把命令式 API 挂 `window.__excelViewerVue2`(对齐 Vue3 `__excelViewer` / React `__excelViewerReact`)。
  - `e2e/vue2-smoke.e2e.ts`:加载示例→canvas 渲染+模型有 sheet、编辑模式 editCell+undo、rectOf 几何、demo 顶栏按钮 1:1 —— 把"只有真 Vue 2 浏览器才暴露"的 DOM 复用/imperative DOM 回归钉死。
  - `data-validation.e2e.ts` 加 Vue 2 行,数据验证下拉 + 整数拦截在 Vue 2 上同样覆盖。
- 基线更新:**368 单测 + 159 e2e**(Vue 3 / React / Vue 2 三壳)。

## [1.7.0] - 2026-06-12

### 新增 — 列表型数据验证:点下拉箭头选值(B4,审计后续专项)

- 以前只对"列表型数据验证"的格画下拉箭头**指示**,点了没反应(没解析选项)。现在:**解析选项**(内联 `"a,b,c"` 拆分 + 同表区域引用 `$A$1:$A$5` 读那几格)存进 `SheetModel.dataValidationLists`;编辑模式下点格内下拉箭头 → **弹可选值菜单**(复用右键菜单宿主,框架无关、三壳共用)→ 点选即 `editCell` 填入(走命令栈,**可撤销**)。只读 / 只读格不弹。
- 解析后续:仅"列表型";其它校验类型(整数/日期/自定义公式范围)仍只解析、不强制(强制是后续编辑增强)。
- 测试:`date-locale.test.ts` 加"内联选项解析进 dataValidationLists";`data-validation.e2e.ts`(Vue+React)点箭头→弹菜单→点选填值→撤销回退。

### 关于审计另一后续专项(矢量 PDF 富文本矢量化)—— 评估后不做

- 审计列的"矢量 PDF 富文本走栅格兜底"经核实:矢量 PDF 对**任何非拉丁(中文)文本**(无自定义字体时)本就走栅格(jsPDF 内置字体画不了 CJK)。本组件主力是中文/WPS 文件,文本在矢量 PDF 里**无论如何都栅格**,所以"富文本矢量化"只对"拉丁文 + 自定义字体"的多色格有意义,价值很窄;且栅格兜底**渲染正确**(只是变图片/不可搜索)。故保持现状。

## [1.6.2] - 2026-06-12

> 一次**深度保真审计**(针对"为什么这么多 bug 等用户拖真实文件才暴露")的产出。审计把同类潜在问题归为两大根因 ——
> **ExcelJS 有损中间层** + **多条平行渲染路径漂移**,逐条用代码核实。本版修掉其中已确认的几条;另有数条经核实为
> **非问题 / 已处理 / 故意为之**(见下);两条大件(数据校验下拉编辑器、矢量 PDF 富文本矢量化)作为后续专项。

### 修复 — 富文本(多色/多段)往返导出丢每段字体

- `xlsx-writer` 导出富文本只写 `{text}`,把每段的颜色/粗斜/下划线/删除线/字号全丢了 → 打开多色文件、编辑、导出后变纯黑文本。现在带回每段字体(`toExcelRichFont`)。往返单测锁定。

### 修复 — 富文本渲染只补了换行,缺 indent / 下划线删除线 / shrinkToFit

- `drawRichText` 1.6.1 加了 wrap,但仍缺普通文本路径早有的:**缩进**、**逐 run 下划线/删除线**、**shrinkToFit**(超宽统一缩放塞进列宽)。现在补齐,跟普通文本路径同档(溢出仍顶对齐)。

### 修复 — 竖排文本忽略垂直对齐;矢量 PDF 文本溢出/双线边框

- 竖排文本恒顶对齐 → 现在尊重 vAlign(middle/bottom,超高顶对齐)。
- 矢量 PDF 导出:wrap 文本折行超过格高时不再被裁掉文头(**溢出顶对齐**,跟屏幕一致);**double 边框**画成两条平行线(原来退化成单线)。

### 修复 — 空格上的批注被丢

- 解析器对"空且无边框无填充"的格返 null 跳过,连**批注**也一起丢。现在空格带批注也保留入模型。

### 审计核实为非问题 / 已处理(本版不改,记录在案)

- **隐藏 sheet**:三壳 tab 早已 `filter(state==='visible')`、导出已保留 state —— 已处理。
- **空格超链接**:ExcelJS 把超链接格给成 object 值(非 null),走非空分支,不会丢 —— 非问题。
- **内置日期格式 15-21**:14/22(短日期/日期+时间)已重映射;15-17 的英文月名、20/21 的 `h:mm` 本就是正确 OOXML 显示,故意不动。
- **合并格边框**:合并边框存锚点格、渲染按锚点 box 画整片,正确 —— 加回归单测锁定。

### 后续专项(审计已列,本版未做,见路线图)

- 数据校验:现仅"列表型"画下拉箭头指示;**全类型解析 + 真下拉选取编辑器**是独立编辑功能。
- 矢量 PDF 富文本:现走**栅格兜底**(渲染正确,只是变图片、不可搜索),**矢量化富文本**是较大专项。

## [1.6.1] - 2026-06-12

### 修复 — 打开 .xlsx 边框线不完整(空但带边框的结构格被丢)

- **现象**:打开带"树形/网格结构"的 .xlsx(大量**空单元格只承载边框**组成框线),渲染出来的框线残缺、盒子合不拢。
- **根因**:解析器用 ExcelJS `eachCell({ includeEmpty: false })` 遍历,**跳过 value 为空的格** —— 而结构框线全靠这些"空但有边框"的格,于是它们的边框(及底色)整批丢失。(实测该文件 H3 = 空格 + 上边框,解析后**根本不在模型里**。)
- **修法**:改 `includeEmpty: true` 遍历;`toCellModel` 对空格**只在有可见样式(边框/填充)时**才入模型,真正空白格(无边框无填充)返 null 跳过 —— 既保住结构格边框/底色,又不把空白格塞进来膨胀。`eachCell` 只扫到该行最右有格的列,大表也不会扫到无限远(大表维度单测仍过)。
- 测试:`date-locale.test.ts` 加"空但带边框的格入模型(上边框保住)、纯空白格不入模型"。

### 修复 — 只读模式右键没有「复制」(复制不改数据,理应可用)

- 之前右键内置菜单只在编辑模式给(`editable ? build : []`),只读模式直接空 → 没法右键复制。改为:**复制**(不改数据)任何模式都给;编辑项(粘贴/插入/删除/合并/清除…)仍仅编辑模式。
- 测试:`edit-contextmenu.e2e.ts` 加"只读模式右键 → 有复制、无编辑项"(Vue+React)。

### 修复 — 富文本(多色/多段)单元格不换行,渲染只显示中间一截

- **现象**:打开真实 WPS 文件,J 列「货品名称」(多色富文本 + `wrapText`)只显示中间几个字(如"配12件套套"),WPS 是从顶部完整换行显示。多点几次工具栏「自动换行」会显示更多 —— 因为那是另一条 autofit 撑高路径。
- **根因**:`canvas-renderer.drawRichText` **完全没实现换行** —— 把所有富文本 run 拍在**一整行**上画,水平居中导致起点远在格左外、再裁切到格内,于是只看到水平居中的那一截"中间文字";而普通文本路径早有折行 + 溢出顶对齐。富文本走的是另一条路,漏了。
- **修法**:重写 `drawRichText` 支持 `wrapText` —— 逐字符按列宽折行(保留各 run 字体/颜色),并套用跟普通文本**完全一致**的垂直对齐 + **溢出顶对齐**(超过行高 → 顶对齐显示文头,WPS 行为)。非 wrapText 仍单行。实测真实文件 J 列现与 WPS 渲染一致(顶部完整换行)。

### 修复 — 打开 .xlsx 日期显示成 `04-01-26`(应为 `2026/4/1`)

- **根因**:OOXML 内置短日期格式(`numFmtId=14`)的"显示"本应跟随区域设置,但 ExcelJS 把它**硬编码成美式串** `mm-dd-yy` → 我们照渲染成 `04-01-26`;WPS/Excel 中文环境显示 `2026/4/1`。(实测真实 WPS 文件:该格 `numFmtId=14`,wrapText/垂直居中/行高都解析正确,**唯一**出入就是这个内置日期格式。)
- **修法**:`exceljs-adapter` 把 ExcelJS 的内置日期串重映射成中文 locale —— `mm-dd-yy → yyyy/m/d`、`m/d/yy h:mm → yyyy/m/d h:mm`(纯时间 / 带英文月名的 `d-mmm-yy` 不动)。需自定义可经 `transformModel` 钩子覆盖。
- 测试:`date-locale.test.ts`(ExcelJS 写 `mm-dd-yy` → 解析回应得 `yyyy/m/d` 渲染 `2026/4/1`;普通货币格式不受影响)。

### 修复 — Vue 2 / React demo 拖文件不解析(弹空白标签 `about:blank#blocked`)

- 拖文件加载是 demo 层能力,Vue 3 `App.vue` 有 `@drop.prevent`,而 vue2-demo / react-demo 漏了 → 浏览器走默认行为把文件当导航打开。按"三 demo 1:1"补上:两 demo 根容器加 `dragover/dragleave/drop` 的 `preventDefault` + `onDrop` 落 `src` 解析。

## [1.6.0] - 2026-06-12

### 新增 — 可配置粘贴行为(覆盖 / 合并 / 仅值)+ 右键选择性粘贴 + 工具栏配置面板 + 只读提示

**问题**:粘贴(尤其从 WPS 粘真实表格)有两类痛点 —— ① 目标区**原有合并/结构没清掉**,旧合并吞列致数据错位(如示例 A1:E1 旧合并把粘进来的前 5 列吞了);② "贴近源(覆盖)"还是"保留目标格式(仅值)"应由用户选,而非写死。

**方案 —— 框架无关的 `PasteBehavior` 配置系统**(core,三壳共用):

- **逐项可配**(默认 = 覆盖式 1:1):`cellStyle` / `fill`(覆盖/合并/不粘)· `rowHeight`(搬源/不动)· `colWidth`(仅首行搬源/总搬/不动)· `sourceMerges`(应用/不应用)· **`targetMerges`(清掉/保留 —— 默认清,修数据错位)** · `images`(落格/不粘)。首行 vs 中间唯一差异收敛在 `colWidth: 'firstRowOnly'`(列宽整列共享,仅粘到首行才取源宽,粘到中间不动上方表头)。
- **两条粘贴路径都按配置走**:`pasteRich`(外部 WPS/Excel)+ `pasteSnapshot`(应用内/跨实例 1:1);样式按「覆盖式以中性默认为基 / 合并式以目标为基 / 仅值留目标」三档算(见 `resolvePastedCellStyle`)。
- **右键「选择性粘贴」子菜单**(core context-menu 加一级 flyout):`覆盖格式(贴近源)` / `保留原样式(仅值)`,逐次预设,不改默认。
- **工具栏「⚙ 粘贴配置」面板**(`paste-config-host.ts`,框架无关 DOM,**三壳/三 demo 共用一份**,UI 天然 1:1):列出全部项下拉自由定制 + 两个快捷预设(覆盖式 1:1 / 仅值)+ 恢复默认;应用即 `setPasteBehavior`。
- **API / prop**(三壳同名):组件 `:paste-behavior`(`Partial<PasteBehavior>`,缺项回落默认)· `viewer.getPasteBehavior()` / `setPasteBehavior(cfg)` / `openPasteConfigDialog()` · `pasteRichHtml(html, at, behaviorOverride)` 第 3 参逐次预设。导出 `PasteBehavior` / `DEFAULT_PASTE_BEHAVIOR` / `PASTE_PRESET_VALUES_ONLY` / `resolvePasteBehavior`。
- **只读检查 + 提醒 = 核心层统一(不按输入方式重写)**:只读**判定**本就是核心层唯一真相源 `resolveEditable()`(`isCellEditable` 全走它),任何改数据的操作(粘贴/编辑/合并/拆分/图片互转)在 EditController 里逐格 `isEditable` 拦截 + 收集 denied + emit `permission-denied`;**提醒**统一收口在控制器 `emitEditEvent` —— 所有 `permission-denied`(dimension 列宽行高=布局除外)经此**一处**按 `readOnlyPrompt` 配置弹内置提醒。新增输入方式只要照常走 EditController API,**无需各自重写只读检查/提醒**(避免遗漏)。
- **只读提醒(逐格精确 + 可配 dialog/toast/none)**:撞只读(**编辑模式下也可能有只读格**,如 `readOnlyRanges`)不再静默 —— 收集**所有被跳过的只读格**(不止落点),按 `readOnlyPrompt`:`'dialog'`(默认,弹窗**列出具体哪些格**A1 引用)/ `'toast'`(顶部气泡)/ `'none'`(只发事件)。框架无关 DOM(`readonly-prompt-host.ts`)三壳共用。粘贴落点只读不再整次中止 —— 可编辑的格照常粘、只读格跳过并在提醒里列出。
- **右键「选择性粘贴」二级菜单可达性修复**:菜单宿主加 flyout 关闭延时 + 紧贴父菜单(消除缝隙),从父项滑到子项不再"还没点到就消失"。
- **Ctrl+V vs 右键差异(已知,浏览器限制)**:`Ctrl+V` 走 `paste` 事件拿**原始 HTML**(WPS 的 `<style>` 类格式全在);右键「粘贴/选择性粘贴」走 `navigator.clipboard.read()`,**浏览器会净化** HTML(删 `<style>`)→ 从 WPS 粘的类格式不如 Ctrl+V 全(应用内 1:1 复制因带 `data-ooxml-clip` 属性不受影响)。无法在右键路径绕过(无 paste 事件)。
- 测试:`paste-behavior.test.ts`(9 个,各档样式解析)+ `paste-behavior.e2e.ts`(8 个:仅值保留目标 / 默认覆盖清目标合并修数据错位 / 配置面板弹出+预设+应用 / **粘到只读区弹对话框列出哪些格+只读格不被覆盖**,Vue+React)。

### 修复 — 从 WPS/Excel 富粘贴丢"自动换行 + 垂直居中"(连带水平居中也看不出)

- **现象**:WPS 里开了自动换行 + 水平/垂直居中的格,粘进来后不换行(长文本溢出/裁切)、垂直贴底,连水平居中也看不出来。
- **根因(两处解析漏洞,均非回归,一直没实现)**:
  - **`white-space` 没解析** → `wrapText` 永远 false。Excel/WPS 用 `white-space:normal` 标记开了自动换行(全局默认 `td{white-space:nowrap}`,换行格用 `normal` 覆盖)。
  - **裸 `td` 元素默认层没收集** → 垂直居中丢。WPS 把"所有单元格默认"(如 `vertical-align:middle`、`white-space:nowrap`、`font-size:11pt`)放在 `td{...}` 选择器上,各 `.etN` 类只覆盖要改的;而 `parseClassStyles` **只收 `.类名` 规则、不收裸 `td` 规则**,于是没写 `vertical-align` 的格全回落成默认 `bottom`。
  - **渲染器本就支持** `wrapText`/水平居中/垂直居中(canvas-renderer 1122/1152/1140 行),只是解析没喂字段。
- **修法**:① `cssToStyleOverride` 增加 `white-space: normal|pre-wrap|pre-line → wrapText=true`;② `parseClassStyles` 额外收集裸 `td` 默认声明,`rawCssOf` 按 CSS 优先级 **td 默认 < 类 < 内联** 三层合并 → 没写垂直对齐的格拿到 `td` 的 `vertical-align:middle`;③ 字号解析认 `pt` 单位(`font-size:11.0pt` → 11,不再当 px 算成 8 —— 之前没喂 td 默认层不暴露,引入默认层后必须修对)。
- 测试:`edit-paste-rich.e2e.ts` 的 WPS 夹具加裸 `td{vertical-align:middle;white-space:nowrap;font-size:11pt}` + 类 `white-space:normal`,断言粘贴后 `hAlign=center` + `vAlign=middle` + `wrapText=true` + `font.size=11`。

> **路径隔离(为什么不会互相串)**:本组件三种"进数据"路径**各走各的解析,零交叉**——① 打开 .xlsx 走 `exceljs-adapter`(OOXML);② 应用内/跨实例复制走 `clipboard-snapshot`(自带 `data-ooxml-clip` 完整快照);③ 外部 WPS/Excel 粘贴才走 `clipboard-html`。本次只改 ③。`pasteRichHtml` 先认快照(② )、不是才退到外部解析(③),所以 ① ② 完全不受影响。每个外部粘贴格的 CSS 也是**独立**算出一份 `CellStyleOverride`(td默认+自己的类+自己的内联),不跨格混用。

### 修复 — 外部富粘贴样式是"合并目标"而非"覆盖"(粘到带色表头会漏出表头底色)

- **现象**:从首行(带绿底表头)开始粘,源里**没写填充**的格(如 WPS 的 `.et6` 没 `background`)粘完仍保留表头的绿底,不是干净覆盖成源的样子。
- **根因**:`pasteRich` 用 `applyStyleOverride` 套源样式,而它以**目标格现有样式为基**做浅合并(给工具栏"加粗"那种增量编辑设计的)。源 patch 没写的属性(填充/边框)就保留目标原有的 → 粘到带色区会漏底色。
- **修法**:`pasteRich` 套样式前先把目标格 styleId 归 0(中性默认),再合并源 patch —— 即**以中性默认为基的覆盖式**,源没写的属性回落默认(无填充/无边框),不再漏目标底色。结果贴近源,同 Excel"粘贴替换格式"、也同应用内 1:1 快照粘贴(后者本就清空目标再落)。
- **为什么所有落点都这么做、不限首行**:样式是**逐格**的,覆盖只影响被粘的那几格、不波及邻格,所以任何位置都安全;列宽因**整列共享**才需限定首行。纯文本粘贴(源无样式 → 无 patch)不受影响,照常保留目标格式。
- 测试:`edit-paste-rich.e2e.ts` 加"粘到红底格 → 源白底覆盖成白、源没写填充的格清成无填充(不漏红底)、边框照常"。

### 变更 — 粘贴只带行高、列宽默认不改(例外:粘到首行时套用源列宽)

- **问题**:粘贴(WPS/Excel 外部富粘贴 + 应用内 1:1 快照粘贴)原会把源**列宽**搬到目标列。但列宽是**整列共享**的,粘到第 18 行却把同列上方的表头(第 1~7 行)宽度一起改了 → 破坏现有表格布局。
- **改法**:两条粘贴路径(`pasteRich` 外部 / `pasteSnapshot` 应用内)都**只搬行高;列宽默认不搬**(以现有表头为准,同 Excel 默认粘贴)。行高是逐行的,只影响被粘的那几行;内嵌图按**目标格尺寸**填充,不依赖源列宽,图照样填满。
- **例外 — 粘到首行(`start.row === 0`)套用源列宽**:此时上方没有任何内容可被破坏,粘贴块本身就是新表头/新布局,列宽应以它为准。`row > 0`(粘进已有表格中间)才保持目标列宽不动。
- 源列宽仍随 `ParsedClipboard.colWidths` / `ClipSnapshot.colWidths` 解析、携带,粘到首行即应用、否则保留。
- 测试:`edit-paste-rich.e2e.ts`(WPS)断言粘到 row 2 列宽不被源覆盖、粘到 row 0 列宽=源 72/120;`clipboard-snapshot.test.ts`(应用内)断言粘到 row 5 目标列宽不变、粘到 row 0 套用源列宽 120。

### 修复 — 空格/新建格/粘贴串入首格底色(根因:解析时 `styles[0]` 不是中性默认,而是第一个被解析到的格样式)

- **现象**:加载示例(A1 是绿底表头)后,从 WPS 粘贴一段内容,部分**本应无底色**的格冒出绿底(看起来像"混入了第 1 行样式");散落分布(有的格白、有的格绿)。打开任意"首格带底色"的本地 .xlsx 也会有同类隐患(空格/编辑新建的格染上首格底色)。
- **根因**:ExcelJS 解析器 `buildSheet` 按**遇到顺序** intern 样式,首个被解析到的单元格(通常是 A1 表头)样式就占据了 `styles[0]`。而全 core 多处把 `styleId 0` / `styles[0]` 当成"中性空白默认基样式"用——空格、`setCellValue` 新建的格、`applyStyleOverride` 对空格的兜底基样式都回落到它。于是 A1 的绿底成了"默认底色":凡是没有显式指定填充的格(如 WPS 类里**没写 `background:`** 的 `.et6`/`.et8`),`mergeStyleOverride` 保留基样式的 `fill` → 冒出绿底;写了 `background:#FFFFFF` 的格才是白 → 散落串色。
- **修法(深修,非兜底)**:抽出唯一规范工厂 `makeDefaultStyle()`([src/core/model/types.ts](src/core/model/types.ts)),`buildSheet` 解析前**预置 `styles[0] = makeDefaultStyle()`**(无填充/无边框中性默认)并登记进 styleIndex,真实格样式从 index 1 起;首格 A1 的绿底样式仍在(只是换了 index),引用它的格不受影响。loader-json / clipboard-snapshot 原本各有一份重复的默认样式工厂,一并改为复用此唯一来源。
- **影响面**:不止修了 WPS 粘贴——所有"空格/新建格/兜底基样式"路径现在都正确回落到中性默认,文件打开渲染零变化(带样式的格按各自 styleId 渲染如初),只有"默认格"不再染上首格底色。
- 测试:`parse.test.ts` 加回归"styles[0] 恒为中性空白默认(首格 A1 有绿底也不占 index 0)";现有 339 单测 + e2e 全绿。

### 修复 — Ctrl+V 改走 paste 事件(根因:clipboard.read() 会净化 HTML 删掉 `<style>`)

- **真正的根因**:我们 Ctrl+V 原先走 `navigator.clipboard.read()`,这个 Async Clipboard API **会净化 HTML** —— 把 `<style>` 块、注释整个删掉。而 WPS/Excel 复制的格式(CSS 类)、数字格式(`mso-number-format`)、内嵌图(VML `o:gfxdata` 注释)**全在这三样里** → 过一遍 `read()` 就没了(实测:写 `<style>` 进剪贴板再 `read()` 读回,`<style>`/类定义/注释全被删)。这就是"直接打开 Excel 文件能完整解析、复制粘贴反而丢格式"的原因:打开文件走 ExcelJS 解 .xlsx(无损 OOXML),粘贴走系统剪贴板的 `text/html`(本就有损)且**还被浏览器二次净化**。
- **修法**:控制器在 `scroller` 上绑 `paste` 事件,Ctrl+V 改走 `onPaste(e)` —— `e.clipboardData.getData('text/html')` 拿的是**原始未净化** HTML(WPS 的 `<style>`/VML 都在),不再走净化的 `read()`。`onKeyDown` 不再拦截 Ctrl+V(在那 `preventDefault` 反而会阻止 paste 事件)。我们自己复制的 `data-ooxml-clip` 快照也照样原样拿到,1:1 不受影响。
- 右键菜单"粘贴"无 paste 事件,仍走 `pasteFromClipboard()`(`read()`,会净化)→ 从 WPS 粘的格式不如 Ctrl+V 全(已在方法注释/文档说明)。
- 测试:`edit-paste-rich.e2e.ts` 加"派发 paste 事件(原始 HTML)→ onPaste → 类格式/numFmt/VML 图都还原"用例(Vue+React);自家 1:1 复制 e2e 仍过(快照不被净化)。

### 修复 — 从 WPS/Excel 富粘贴丢格式/数字/图片(对照真实 WPS 剪贴板 HTML)

- **解析 `<style>` 类样式**:Excel/WPS 复制的剪贴板 HTML 把单元格格式放在 `<style>` 块的 CSS 类里(`<td class=et3>` + `.et3{border:…;background:…}`,整段还包在 `<!-- -->` 里),而旧解析只读每个 `<td>` 的内联 `style=`(`DOMParser` 不会把类规则套到元素上)→ 边框/底色/字体全丢。现在 `parseClipboardHtml` 先收集所有 `<style>` 块的「类名→声明」(剥掉 `<!-- -->` 壳),落格时把命中类的声明合并进 `td.style`(类在前、内联在后,内联优先)再解析 → 还原边框/填充/字体/对齐。
- **数字格式(日期/货币不再变成裸序列号)**:格式码在 `mso-number-format`(CSSOM 会丢弃这种私有属性),且值是 CSS 转义的(`2`→`"`、`\#`→`#`、`\;`→`;`、`\(`→`\(`)。新增 `parseMsoNumberFormat`/`unescapeMsoNumFmt` 从原始声明串解析并解转义 → `numFmt`(如 `yyyy/m/d`、`"￥"#,##0.00_);[Red]\("￥"#,##0.00\)`),配合 `x:num` 原始序列号 → 日期/货币正确显示(之前 46113 直接显示成裸数字)。
- **列宽/行高 1:1**:剪贴板 HTML 带了 `<col width=N span=M>`(列宽 px)和 `<tr height=N>`(行高 px),`parseClipboardHtml` 现在解析出 `colWidths`/`rowHeights`,`pasteRich` 用 `restoreDimension` 搬到目标列/行 → 列宽行高 1:1。内嵌图填满单元格,**格尺寸对了图也就对了**(之前列宽行高没搬 → 图也显得不对)。(实测真实 WPS:行高 106px、列宽 72/124/197… 与 `<col span>` 完全对应)
- **图片(WPS 区域复制内嵌图能救回来)**:之前以为是浏览器硬限制——其实 `<img src="file:///…">` 确实读不了,但 WPS 同时把图放在 VML `<v:shape o:gfxdata="base64">`(在 `<!--[if gte vml 1]>…<![endif]-->` 注释里),**那段 base64 是个 zip,内含 `media/imageN.png`**。新增 `extractVmlImageDataUrl`:从 td 的注释节点取 `o:gfxdata` → `unzipSync`(fflate)→ 取图 → data-uri → 走现有图片落格(转 DISPIMG 单元格图)。
- **不影响应用内 1:1 复制**:本组件自己复制的内容带 `data-ooxml-clip`,粘贴时先走快照路径(`parseSnapshotHtml`/`pasteSnapshot`),根本不进 `parseClipboardHtml`,零影响。
- 图片方向澄清(文档):WPS/Excel **区域复制**的图片在剪贴板里是 `file:///` 本地路径,浏览器读不到 → 区域里的图必然丢(浏览器限制);单图复制走 `pasteImageBlob` 仍可。
- 测试:`e2e/edit-paste-rich.e2e.ts` 加 Excel/WPS 类样式(`<style> .xl 类`)用例(Vue + React),断言边框/填充/字体/对齐还原。

## [1.5.0] - 2026-06-11

### 新增 — 应用内复制粘贴 1:1 保真(走剪贴板嵌入快照)

**问题**:复制粘贴走系统剪贴板的 HTML/TSV 交换(为了能贴进 Excel/WPS),而复制端只 emit `<td style=CSS>格式化文本`,导致:① 合并单元格被拍平;② 图片(DISPIMG / 浮动图)整个丢失;③ 数字按格式化文本复制(`¥237` → 文本"¥237",丢原始值);④ 边框/数字格式/行高列宽不带。

**方案(全走剪贴板,跨实例 1:1)**:复制时把**完整模型快照**序列化(base64 UTF-8 JSON)嵌进剪贴板 HTML 的 `data-ooxml-clip` 属性;粘贴时 `pasteRichHtml` 优先识别该快照走 1:1 还原,否则回退原有外部 HTML 近似解析。因为快照随剪贴板走、不依赖内存,**Vue3 / Vue2 / React 三壳之间、跨标签页互相复制结果都一致**。外部应用(Excel/WPS/Word)忽略该属性,只读可见 `<table>`(同时增强:补 `colspan/rowspan` + `<img data:>`,所以贴进 Excel 也能带上合并和图片)。

- 新增 [src/core/edit/clipboard-snapshot.ts](src/core/edit/clipboard-snapshot.ts):`ClipSnapshot` 序列化/反序列化(每格 原始值/类型/公式/超链/批注/富文本/dispImgId + **完整 CellStyle**;合并;浮动图 base64;DISPIMG 字节;**行高/列宽/手动行高标记**),UTF-8 安全 base64。
- 新增 `EditController.pasteSnapshot`(覆盖式 1:1 落格:intern 样式到目标表、登记 DISPIMG 字节、平移合并/图片/行高列宽;整体单次撤销 + 前后快照 cell-change + 只读 permission-denied),`mutations.setCellModel`。
- `copySelection` 增强 HTML(合并 colspan/rowspan + 图片 `<img data:>`)+ 嵌入快照;`pasteRichHtml` 先试快照再回退。
- 不改三壳(纯 core);`Ctrl+C`/`Ctrl+V`/右键菜单复制粘贴全部受益。

**大图护栏(三项)**:
- **不双重 base64**:图片字节不再既进快照又进可见 `<img>`,改为只在可见 `<img data-clip-img="key">` 存一份(`key`=DISPIMG `c:id` / 浮动 `f:序号`),快照只引用;粘贴时 `parseSnapshotHtml` 从 `<img>` 回填字节(`reattachImages`)。剪贴板体积从约 3× 原始图字节降到约 1.4×,有效上限翻倍。(`serializeSnapshot(..., { withImageBytes: false })`)
- **字节预算 + 优雅降级**:复制区图片总字节超 `CLIP_IMAGE_BUDGET_BYTES`(6 MB)→ 自动**降级为"无图 1:1 复制"**(样式/数字/边框/合并/行高列宽仍 1:1,只跳过图片,DISPIMG 格中性化为空格),避免剪贴板超限导致整次复制静默失败。(`withoutImages`)
- **降级通知**:降级时经现有 `permission-denied` 通道发事件(`reason: 'copy'` + 中文 message:"复制内容含图过多…已按无图复制"),壳/插件可 toast,不再静默。
- **贴进 WPS/Excel 图片不再巨大**:可见 `<img>` 改为带 `width`/`height` **属性**(按单元格大小:DISPIMG 用所在格列宽×行高,浮动图用其 EMU 尺寸),WPS/Excel 认属性不认 CSS `max-width` —— 之前只给 CSS 导致按原图像素(常几百上千 px)贴入显得巨大。

- 测试:`src/core/edit/__tests__/clipboard-snapshot.test.ts`(6:序列化往返 + `pasteSnapshot` 1:1 + undo + 瘦身回填 1:1 + 降级中性化 + 脏数据回退);`e2e/copy-paste-fidelity.e2e.ts`(真系统剪贴板 Ctrl+C→Ctrl+V,数字仍是数字不退化文本)。

## [1.4.0] - 2026-06-11

**透视表(Pivot Table)完整闭环** — WPS 式创建入口 → 右侧字段面板 → 编程 API → **导出真实 OOXML 透视表零件**,并支持**活刷新 / 折叠展开 / 多选筛选**。整个功能由 **`pivotTable` 配置开启,默认关闭**(三 demo 已开启)。另含 `scrollToCell` 导航 API 与工具栏 `sort` 内置项。

### 新增

#### 透视结果"活"化 + 多选筛选(2026-06-11)

- **活刷新**:编辑源数据区任意单元格(含撤销/重做,统一走 `onModelChange` chokepoint)后,所有透视表按其 `source` 区域自动重算 —— 从"静态快照"变"活对象"。重算经唯一入口 `recomputePivot`(面板改布局 / 活刷新 / 折叠展开三处共用),直接改模型不入命令栈(派生态),`pivotRefreshing` 防重入,无透视表 / 功能关闭时零开销。
- **行分组折叠/展开**:放 ≥2 个行字段时 `buildPivotRows` 产出"大纲"——外层分组行带小计、内层缩进明细;分组表头行首画 [−]/[+] 折叠按钮(canvas 绘制 + `CanvasRenderer.pivotToggleAt` 命中,与 autofilter 下拉同款),点击折叠/展开该组。折叠状态存 `PivotTableModel.collapsed`,运行时 `rowGroups` 记录可折叠表头行供渲染/命中。单行字段退化为扁平结果(无折叠),既有行为不变。
- **空白起步 + 字段勾选联动 + 新建表显示**(2026-06-11,修用户反馈):① 对话框创建的透视表不再自动猜字段(此前会把"首个数值列=值、首个其它列=行"硬塞,常把"编号/ID"列当行字段),改为空白占位 + 面板选字段填充,对齐 WPS/Excel;`createPivotTable` 仅在显式传 `layout.rows/values` 时沿用字段(编程 API 不变)。② 字段列表复选框从 `disabled` 改为可点 —— 勾选即加入(数值→值/其它→行),取消即移出,顶部字段列表与底部四区双向联动。③ 修「新建工作表」输出后新表 tab 不显示:Vue 3 壳 `workbook` 是 `shallowRef`,`sheets.push` 不触发 `SheetTabs` 的 `computed` 重算 → 给 SheetTabs 加 `:key` 版本号,`onActiveSheetChange` 时 +1 强制重渲(React/Vue 2 壳内联重读,本就正常)。
- **四区大白话 tooltip + 列交叉表确认**:字段面板「筛选器 / 列 / 行 / 值」每个区加 hover 说明(区标题带 ⓘ),筛选器/列空着时方框内还有一句引导;补 e2e 证实「列」字段确实横向展开成二维交叉表。
- **多选筛选**:`PivotFilterMode` 增 `'include'`(`PivotFilterRule.values` 列出保留值)。字段面板筛选 chip 点开底部明细面板,可选 全部 / 非空 / 多选(勾选具体值,带全选/清空)。导出对齐 WPS:`include`/`non-empty` → `multipleItemSelectionAllowed` + 未选项 `item@h=1`。

#### `pivotTable` 功能开关(默认关闭,opt-in)

#### `pivotTable` 功能开关(默认关闭,opt-in)

- 三壳同名 boolean prop(`:pivot-table` / `pivotTable`),进 `EditConfig.pivotTable`。**默认 `false`:工具栏 `pivot-table` 入口不渲染、`createPivotTable`/`openPivotTableDialog` 等 API 返回 `false` 并提示、导出不回注 pivot 零件 —— 行为与历史版本完全一致(零回归)**。开启后(还需 `editable`)下述全部能力生效。直接用 core `workbookToXlsxBlob` 时对应 `XlsxExportOptions.pivotTables`(经 viewer 导出时自动随开关注入)。

#### 透视表创建(WPS 风格,core 落地、三壳共用)

- **工具栏 `pivot-table` 入口**(需 `editable`):选中带表头数据区 → 「创建透视表」对话框选择生成位置(现有工作表指定单元格 / 新建工作表)→ 写出静态透视汇总表,入命令栈(undo 整体还原)。
- **右侧「数据透视表」字段面板**(框架无关 body 级 DOM,[src/core/viewer/pivot-dialog-host.ts](src/core/viewer/pivot-dialog-host.ts)):字段搜索;按钮/拖拽把字段加入 **筛选器 / 列 / 行 / 值** 四区;拖到移除区删除;筛选器支持 全部/非空/具体值;值字段可多个、汇总方式可切 求和/计数/平均值/最大值/最小值;每次变更重建静态结果。
- **编程 API**:`createPivotTable({ sourceRange, sourceSheetIndex, output, layout, showPanel })` 不经页面直接创建;`createPivotTableFromSelection()` 选区快捷;`openPivotTableDialog()` 打开入口对话框。三壳(Vue3 ref / React handle / Vue2 viewerApi)+ 插件 `viewer` 均已暴露。
- **模型元数据**:`PivotTableModel` 保存 `source`(源表 + 源区域)与 `layout`(四区布局),`cloneWorkbook` 深克隆,undo 快照不被面板操作污染。

#### 导出真实 OOXML 透视表零件 ([src/core/export/pivot-tables.ts](src/core/export/pivot-tables.ts))

- ExcelJS 不建模 pivot 零件,写出后在 **zip 层回注**(同 WPS cellimages 模式):`pivotCacheDefinition`(cacheSource + cacheFields/sharedItems)+ `pivotCacheRecords`(源数据行)+ `pivotTableDefinition`(location / pivotFields / row/col/page/dataFields / 样式)+ workbook `<pivotCaches>` + 全套 rels + `[Content_Types].xml`。
- `refreshOnLoad="1"`:Excel/WPS 打开导出件即识别为**真透视表**并按源区域重算原生布局;静态汇总结果仍写在单元格里,不支持透视的查看器也能看。
- **筛选器导出语义对齐 WPS**:"= 具体值"写 `pageField@item` 指向选中项(打开还原筛选状态);"非空"映射为多选 + 隐藏空白项(`multipleItemSelectionAllowed` + `item@h`,即 WPS"去掉空白"语义);"全部"不写选中。
- **overlay 导出保留原文件透视表**(`restoreOriginalPivotPartsIntoZip`):原文件已有的透视表(解析为只读,ExcelJS load→write 会整套丢掉)在 overlay 模式下从原件 zip **原样搬运**整套零件(pivotCache/pivotTables 目录 + workbook 注册 + worksheet 隐式关系按表名重挂 + Content_Types),cacheId 保留、r:id 重新分配;后续 App 新建透视表的零件编号/cacheId 自动避开。"打开 → 编辑 → 另存,透视表仍在"。rebuild 模式因结构可能被增删行列改动,不搬运(退化为普通单元格)。
- 回注/搬运失败自动降级为纯静态结果,不影响主体导出。

#### 其它

- `scrollToCell(row, col, { select? })` 命令式导航 API(三壳 + 插件 viewer),超出当前虚拟区自动扩展。
- 工具栏内置 `sort` 项(按活动单元格所在列升/降序;未开自动筛选先按选区/已用区建立范围),`viewer.sortActiveColumn(dir)` 同步暴露。

### 修复

- **pivot-parser 支持标准 ECMA-376 隐式关联**:真 Excel 文件的透视表零件靠 worksheet rels 关联(sheet XML 里没有元素),此前只认 worksheet XML 内的 `pivotTableDefinition` 引用 → 标准文件解析不到只读透视表按钮。现在两条路径都认(rels 扫描 + 兼容引用),导出件可往返解析。

### 测试

- 单测 316 → 330(pivot-parser 解析 + clone 元数据 + 导出回注零件结构/往返 + equals/non-empty/include 筛选语义 + overlay 原件搬运/编号避让等);e2e 118 → 124(`e2e/pivot.e2e.ts`:UI 入口全链路 + 面板切换汇总方式 + undo;API 新建工作表 + 求和;导出 zip 零件断言;活刷新随源编辑/撤销;2 行字段折叠/展开;include 多选筛选;面板筛选复选框)。

## [1.3.3] - 2026-06-09

**Vue 2.6 真实兼容修复合集** — 1.3.2 上线后消费方 Vue 2.6.12 + vue-cli 4 (webpack 4) 项目验证暴露两个 Vue 2.6 特有 bug (函数 ref 不支持 / `ctx.expose` shim 语义不同), 1.3.3 一并修掉. **现在 Vue 2.6 / 2.7 / Vue 3 三个版本都真正可用**.

> 1.3.2 已发布到 npm, 但 Vue 2.6 仍不可用 — 1.3.3 是必须升的真兼容补丁.

### 修复

#### Vue 2.6 上 `$refs.viewer.*` 命令式 API 全拿不到

**根因**: 之前用 `expose?.(viewerApi)` 暴露 60+ 命令式方法 (load / getSelection / setStyle / downloadXlsx / beginEdit / commitActiveCellValue / undo / ...). Vue 3 / Vue 2.7 原生 expose 走标准路径 OK, **但 Vue 2.6 + @vue/composition-api 1.7.x shim 下 `ctx.expose` 的语义是"暴露 setup 返回值的指定 key"** —— 本组件 setup 返回的是 render function (没有可暴露的 key), shim 下 expose() 是 no-op. 消费方 `this.$refs.viewer.downloadXlsx(...)` 全部 `undefined is not a function`.

**修复** ([src/vue2/ExcelViewer.ts](src/vue2/ExcelViewer.ts)): 在 Vue 2 下用 `Object.assign(vm, viewerApi)` 直接挂到 Vue 2 instance proxy (Vue 2 instance 是普通对象, 可 assign). Vue 3 仍走 `expose()`. 通过 `vm._isVue` 检测 Vue 2 (Vue 3 instance proxy 没这个 flag):

```ts
expose?.(viewerApi)
if (vm && (vm as any)._isVue) {
  Object.assign(vm, viewerApi)
}
```

**验证 (Vue 2.6.12 + composition-api 1.7.2 真实环境)**:
- `$refs.viewer.{getWorkbook, getActiveSheet, getSelection, setSelection, getCellText, getSheetJSON, beginEdit, commitActiveCellValue, isCellEditable, downloadXlsx, downloadJson, undo, canUndo}` 13 个核心方法 `typeof === 'function'` ✓
- `getWorkbook()` 真返工作簿 (3 sheets, 销售报表) ✓
- 编辑流: `beginEdit(0,0) → isEditing=true → commitActiveCellValue('VUE26_OK', 'down')` → 单元格值变 "VUE26_OK" ✓
- `undo()` → 恢复 "2026 年度销售汇总" ✓
- 0 控制台错误 ✓

#### Vue 2.6 上 renderArea / fb / templateInput 三个 DOM 全拿不到 → canvas 不挂

**根因**: `src/vue2/ExcelViewer.ts` 之前用**函数 ref (callback ref)** 拿 DOM:
```ts
function domSlot<T>() {
  const slot = { value: null, bind: (el) => { slot.value = el } }
  return slot
}
h('div', { ref: renderAreaSlot.bind, class: 'ov-render-area' }, [...])
```

**但函数 ref 是 Vue 3 引入、Vue 2.7 才 backport 的特性**. **Vue 2.6 的 vnode ref 只认字符串 ref** — 传函数根本不会被 Vue 调用, `slot.value` 永远是 null, onMounted 里 `renderAreaSlot.value`、`fbSlot.value`、`templateInputSlot.value` 全是 null, 画布初始化整段挂不上.

(为什么内部 demo (`vue2-demo`) 一直 OK? 因为 demo 的 `vue2` 别名解析到 `vue@2.7.16` —— 2.7 起函数 ref 已 backport, 跑起来正常. 直到消费方在真 Vue 2.6.12 项目里装包才暴露.)

**修复**: 把 `domSlot` 从函数 ref 改成**字符串 ref + `vm.$refs` getter**, 三个版本 (Vue 2.6 / 2.7 / Vue 3) 统一支持字符串 ref:
```ts
function makeDomSlotFactory(vm) {
  return function domSlot<T>(refName: string) {
    return { refName, get value() { return vm.$refs[refName] ?? null } }
  }
}
// render: `ref: slot.refName` (string)
// access: `slot.value` getter → 实时读 vm.$refs[name]
```

并保留 onMounted 体内的 `nextTick` 兜底, 让 Vue 3 (refs 在 mounted 后 flush) 也能拿到. 修复影响 3 处 DOM 槽: `renderAreaSlot` (画布根 div) / `fbSlot` (公式栏 textarea) / `templateInputSlot` (隐藏的模板 file input).

**验证**: 本地把 `vue2` devDep 临时改成 `npm:vue@2.6.12` + `Vue.use(VueCompositionAPI)`, 跑 `npm run dev:vue2`, headless 检查: renderArea (1280×509) / canvas (1280×509) / 8 个 controller 子节点全到位, 控制台 0 错误.

### 修复

#### 1. dist 不再含 `import.meta.url` + module worker (webpack 4 SyntaxError 致命)

之前 lib-vue2 build 没用 worker stub, `dist/vue2.js` 含:
```js
new Worker(new URL("/assets/parse.worker-xxx.js", import.meta.url), { type: "module" })
```
- `import.meta.url` 是 ESM 专属语法, webpack 4 解析失败 (Module parse failed: Unexpected token)
- `/assets/parse.worker-xxx.js` 是 Vite 构建产物, 在别的打包器/项目里根本不存在

**修复**: vite.config.ts 把 `isLibBuild` 判定从 `command === 'build' && !isDemoSite && !isVue2Build` 改成 `command === 'build' && !isDemoSite` (Vue 2 lib build 也用 worker-client.stub.ts 走主线程解析). 现在所有 lib 产物 (`dist/index.js` / `react.js` / `vue2.js`) 都不含 `new Worker` 也不含 `import.meta`.

#### 2. dist 降级 `??` / `?.` 等 ES2020 语法 (webpack 4 SyntaxError 致命)

之前用户报:
```js
r.onerror = () => o(r.error ?? new Error("文件读取失败"))  // ?? 是 ES2020, webpack 4 不识别
```
- 多数项目默认不转译 node_modules (webpack 4 / vue-cli 4 尤甚), 用户被迫加 transpileDependencies

**修复**: vite.config.ts 给所有 lib build 加 `target: 'es2018'`. ES2018 没有 `??` (ES2020) / `?.` (ES2020), 但保留 async/await + spread/rest (ES2017+, webpack 4 都支持), 不引入大量 polyfill.

#### 3. 包根加 vue2.js / react.js / core.js stub (webpack 4 不读 package.json#exports)

之前用户报:
```js
import ExcelViewer from 'ooxml-excel-editor/vue2'  // webpack 4 解析失败
// 被迫: import ExcelViewer from 'ooxml-excel-editor/dist/vue2.js'
```

**修复**: 包根新增 3 个 stub 文件:
- `vue2.js` — re-export `./dist/vue2.js`
- `react.js` — re-export `./dist/react.js`
- `core.js` — re-export `./dist/core.js`

webpack 4 解析 `ooxml-excel-editor/vue2` → 找包根 `vue2.js` → re-export 到 dist. webpack 5 / Vite / Rollup 仍走 `package.json#exports` → `./dist/vue2.js` (stub 不会被用到, 0 性能影响).

文件加进 `package.json` `files` 数组, npm publish 时一并发布.

#### 4. echarts / exceljs / jspdf / hyperformula **inline 进 dist chunks/** (彻底)

试过 3 个方案都不彻底:
- 1.3.1 `peerOptional`: npm 7+ 触发 ERESOLVE 冲突
- 1.3.2 第一版 文档说明 + 用户按需装: webpack 4 静态扫描 `await import()` 找不到模块发 warning
- 1.3.2 第二版 改 `dependencies` 自动装: webpack 4 解析这些 lib **源码** 时仍报 named-export / class-fields 错误 (因为 lib 用 ES2020+ 语法)

**终极修复**: rollupOptions.external 区分两类:

**inline 编进 chunks/** (老打包器零解析):
- `exceljs` — 私有依赖, 用户项目一般没用
- `hyperformula` — class fields 语法 webpack 4 解析必炸, 必 inline
- `jspdf` — 同样现代语法

**external + dependencies (npm 自动装)**:
- `echarts` — 用户项目大概率已有自己的 echarts (常见 viz lib), inline 会:
  - 浪费 ~1 MB bundle 体积
  - 跟用户 `echarts.registerTheme()` 注册的主题 dual instance, 用户主题在我们 inline 的实例上失效
  现代 echarts 5.x CJS / UMD 入口 webpack 4 兼容良好, external 安全.

target: `es2017` 把 class-fields / `??` / `?.` 全降级.

消费方:
- **完全不用手动装任何 lib** — `npm i ooxml-excel-editor` 自动装 echarts
- webpack 不解析 exceljs / jspdf / hyperformula 源码 (已 inline 嚼碎)
- echarts 跟消费方自己的 echarts 同实例 (theme 共享, 不 dual)
- 不再有 named-export / class-fields / `import.meta` / module worker / `??` 等老打包器报错

变化:
- **dependencies** 仅保留 `echarts` (+ 小 lib `fast-xml-parser` / `fflate`)
- **dependencies** 移除: `exceljs` / `jspdf` / `hyperformula` (已 inline)
- **peerDependencies** 剩余: `vue` / `react` / `react-dom` / `@vue/composition-api`
- **dist 体积**: 5.4 MB (含 3 个 inline lib)
- **tgz 压缩后**: 1.33 MB
- vite/rollup target: `es2017`

消费方 build 后 bundle (不用 PDF/编辑公式) **影响很小** — dynamic import 保留 code-split, 不调到 chunk 就不下载.

### 总效果

任何打包器都能消费 1.3.2:
- ✅ Vite / webpack 5: 走 `package.json#exports`, 用 dist/ 真实产物
- ✅ webpack 4 / vue-cli 4: 走包根 stub 文件, dist 已降级语法 + 无 worker
- ✅ Parcel / Snowpack / esbuild: 同上
- ✅ Node CJS (SSR): 同 webpack 4 路径

### 升级提示 (从 1.3.0 / 1.3.1 升 1.3.2)

无 breaking, `npm i ooxml-excel-editor@1.3.2`. 4 个运行时 lib (`echarts` / `exceljs` / `jspdf` / `hyperformula`) 已改成 `dependencies` 随包自动装, **不用手动 `npm i`**. 之前手动装的不用卸载, npm 会复用现有版本 (前提是 `^range` 满足).

Vue 3 / React 用户**无变化**.

---

## [1.3.1] - 2026-06-08

**Vue 2 子入口扩展兼容到 Vue 2.6.x** — 通过 `@vue/composition-api` package 让一份代码同时支持 Vue 2.6 + 2.7+. 1.3.0 只支持 2.7+ (内置 Composition API), 此版本不破坏 2.7 用户用法.

### 新增

- **Vue 2.6.x 支持**: peerDeps `vue` 从 `^2.7.0 || ^3.4.0` 扩到 `^2.6.0 || ^2.7.0 || ^3.4.0`
- **`@vue/composition-api` 加入 peerDependencies** (optional peer): Vue 2.6 用户必装 + `Vue.use(VueCompositionAPI)`;Vue 2.7+ 用户也装 (它自检测内置 API,plugin 是 noop,无需 Vue.use)

### 实现

- Vue 2 壳代码 (`src/vue2/ExcelViewer.ts` + `use-excel-document.ts`) 的 Composition API import 从 `'vue2'` 改成 `'@vue/composition-api'` (一份代码同时支持 2.6 plugin / 2.7 内置)
- `vite.config.ts` dev alias `@vue/composition-api` → `vue@2.7 dist`(dev 时拿 2.7 内置), build external (消费者自己解析)
- docs/Vue2.md 加 Vue 2.6 vs 2.7 安装步骤对比 + 解释为什么要装 `@vue/composition-api`

### 升级提示 (从 1.3.0 升 1.3.1)

**Vue 2 用户**: 必须额外装 `@vue/composition-api`. 1.3.0 装的用户升级时 `npm install` 会提示缺失 peer, 跟着装即可:

```bash
npm i @vue/composition-api
# Vue 2.6.x 还需 main.js: Vue.use(require('@vue/composition-api').default)
# Vue 2.7+ 不需要 Vue.use (plugin 自动 noop)
```

Vue 3 / React 用户**无变化**.

---

## [1.3.0] - 2026-06-08

**Vue 2 兼容子入口 + 三壳 UI 1:1 复刻 + 独立 dev scripts** — Vue 2.6 / 2.7+ / Vue 3 / React 三个壳视觉与交互完全一致 (Vue 3 SFC 是参考实现 Standard, Vue 2 / React 1:1 复刻).

### 新增

- **`ooxml-excel-editor/vue2` 子入口** (`src/vue2/ExcelViewer.ts`, ~1000 行 render function 版): **Vue 2.6.x + 2.7+** 兼容壳, 跟 Vue 3 壳 1:1 功能对齐. 通过 `@vue/composition-api` package 同时支持 2.6 (装 plugin + `Vue.use`) 和 2.7+ (plugin 自检测 + noop)
  - 全部 28 项 props (src / workbook / jsonOptions / templateFile / fileName / theme / cellStyle / cellImageFit / imageLightbox / openLinks / plugins / toolbar / editable / cellReadOnly / readOnlyRanges / editableTargets / strictDimensions / readOnlyCellStyle / editor / recalc / formulaEngine / contextMenu / exportProgress / transformModel / templateName)
  - 全部 15+ events 跟 Vue 3 同名 (rendered / error / progress / cell-click / cell-dblclick / selection-change / sheet-change / hyperlink-click / cell-change / edit-start / edit-commit / dim-change / dirty-change / image-change / struct-change / permission-denied / before-context-menu / context-menu)
  - 完整命令式 API (80+ 方法) 跟 Vue 3 `viewerApi` / React `ExcelViewerHandle` 对齐
  - 插件系统完整 (events / cellStyle / theme / overlay / toolbar / setup / contextMenu / editor)
- **独立 dev scripts**: `npm run dev:vue3` (port 5300, 默认) / `npm run dev:react` (port 5301) / `npm run dev:vue2` (port 5302), 三个 demo 进程隔离 + 独立端口

### UI 1:1 复刻 (★ 新中心原则)

**Vue 3 SFC 是参考实现 (Standard)**, 任何 UI 变动先 Vue 3 落地再复刻到 Vue 2 / React. 三壳 + 三 demo 视觉/交互对齐:

- **顶部 ViewerToolbar**: 文件名 + " · 模板: <name>" + 灰色 "N 个工作表" + "导出 ▾" 下拉 (PNG / PDF 位图 / PDF 矢量 / 打印 / 导出设置…) + 缩放 [− / select / +]
- **Action ToolBar**: 完整 9 项内置工具 (find / filter / clear-filter / copy / wrap-text / template / image-tools / freeze / export / zoom) + SVG 图标 (`src/components/toolbar-icons.ts`) + 下拉子菜单 + separator 分组
- **状态栏**: 选区范围 (A1:E5) + 计数 / 求和 / 平均 / 最大 / 最小 (调 `renderer.selectionStats`)
- **cell tooltip**: `onTooltip` hook → ref<TooltipState>, 支持 default / comment 黄底批注样式
- **ExportProgressOverlay**: 居中模态 + stage 标签 + 进度条 + 取消按钮 + `chain()` 包装所有 export API + `:export-progress` prop
- **ExportDialog**: 高级导出配置 (范围 / 清晰度 / 内容 / PDF 类型 / 纸张) — 工具栏 "导出设置…" 弹出
- **overlay scoped slot**: `<template v-slot:overlay="{ rectOf, rectOfRange, tick }">` 跨壳同 API
- **loading / error / empty 三态浮层**: progress bar + 错误提示 + 空状态
- **三 demo 入口** (App.vue / vue2-demo / react-demo) **绿色头共享** `src/demo-shared/demo-bar.css`:
  - 共同按钮: 选择 .xlsx / 加载示例 / JSON 示例 / 编辑模式 / PDF(页码+水印) / 数据→JSON / 贴合 select / ↓XLSX / ↓CSV / ↓JSON
  - 编辑模式按钮组: 设置可编辑 (EditTargets 白名单 dialog) / 高亮只读 / B 加粗选区 / 合并 / 拆分 / 背景 / 字体 / 清除填充 / 整表嵌入 / 格→图 / +行 / -行
  - "⋯ 更多" 溢出折叠 (ResizeObserver + 测量行)

### 修复

- **Vue 2 patch 复用 DOM 致 controller stale 引用 (致命空白渲染 bug)**: Vue 2 patch children 算法没 key 时按 tag 匹配, 把 `.ov-render-area` div 复用成 `.ov-toolbar` (改 className + 替换内容), controller.els.renderArea 变 stale (指向 toolbar 元素), 内部 canvas/overlays 都被 Vue 一并清掉. **修复**: 所有 chrome 子节点加唯一 `key` + canvas/overlays/scroller/editor-slot 改 imperative DOM (onMounted createElement + appendChild, Vue 完全不碰)
- **chrome computed 读 renderTick 导致滚动卡顿**: `fbCanEdit` / `formulaBarEditString` / `renderActionToolbar` 读了 `renderTick.value`, 每帧 scroll → root render function 重跑 → 整个 chrome DOM patch. **修复**: 只依赖低频 selVersion/findVersion/filterVersion
- **`updated()` 钩子内改 reactive → 死循环**: `demoRemeasure()` 改 `demoItemWidths` 触发 re-render → 又 `updated()` → CPU 100% 卡死. **修复**: 改用 `watch` 监听具体根本依赖
- **rAF 等浏览器 layout**: chrome 刚加上时 `renderArea.clientHeight` 是中间态 36px (不是最终 537px), `await nextTick()` 后加 `await rAF` 保证 layout 完成再 measure

### 已知限制

- `dist/vue2.js` 424 KB (含内嵌 core) — Vue 2 build pass 独立, 不共享 `chunks/`。后续探索 rollup 多入口共享 chunk
- `dist/vue2.d.ts` 类型声明暂未生成 (vue-tsc 不认 Vue 2 SFC; 本入口是 .ts 后续可补)
- Vue 2 e2e 覆盖待补 (Playwright 多 entry 配置)

### 文档

- 新 [docs/Vue2.md](docs/Vue2.md) — Vue 2 子入口完整使用文档 + 跟 Vue 3 差异表 + Vue 2 壳特殊坑(函数 ref / ResizeObserver / updated 死循环 / template v-for key)
- [README.md](README.md) 加 Vue 2 入口章节 + 四个子入口对照表 + 三壳 UI 1:1 说明
- [ARCHITECTURE.md](ARCHITECTURE.md) 加 demo-shared / 三壳对齐 / 第 7 中心原则段
- [CLAUDE.md](CLAUDE.md) 加第 7 中心原则 (UI 1:1 复刻) + UI 1:1 工作流 + Vue 2 壳特殊坑 + 路线图更新

### Commits (累计 9 个)

`ba6470c` 修空白渲染 + dev scripts → `0f9718a` statusbar/tooltip/ExportProgressOverlay → `d74c9f1` ExportDialog/overlay slot → `e735e8f` Vue 2 1:1 工具栏 → `b075e05` React 1:1 工具栏 → `5462f5e` demo 绿色头 + 演示按钮 → `ee199c3` EditTargets dialog 迁移 → `411e328` Vue 2 ⋯ 更多溢出折叠 → `a312222` 修死循环

---

## [1.2.1] - 2026-06-08

**WPS 风格长文本编辑** — 默认编辑器从 `<input>` 升级 `<textarea>`, 输入长文本自动换行 + 向下撑高, 跟 WPS / Excel 用户习惯一致。

### 修复 + 改进
- **★ 默认编辑器换 textarea + 动态撑高(WPS 风格, 2026-06-08)** — 用户反馈 / 截图: 编辑长文本时, 旧 `<input>` 单行水平溢出被裁切, 看不到全文; WPS 同场景**编辑框向下浮起 + 自动换行**显示完整内容。修复:
  - [default-editor.ts](src/core/edit/default-editor.ts) `<input>` → `<textarea>`, `white-space:pre-wrap` + 行高同步渲染层 `LINE_HEIGHT_FACTOR=1.18`. 短文本 (`rows=1`) 视觉跟之前 `<input>` 一致, 不破坏短输入体验. **Shift+Enter** 插入换行, 普通 **Enter** 提交 (跟 Excel/WPS 一致).
  - 新钩子 `CellEditorReturn.getDesiredHeight(widthPx) → number` + `CellEditorContext.reposition()` — 编辑器 input 事件触发 reposition, host 重测撑高. 用 [text.ts 的 `wrapLines`](src/core/render/text.ts) 复用渲染层换行算法, 保证视觉一致.
  - [editor-host.ts](src/core/edit/editor-host.ts) `position()` 新逻辑: 宽度仍 = 列宽 (跟 WPS 一致, **仅向下溢出**); 高度 = `max(单元格原高, getDesiredHeight(列宽))`, 上限 viewport 一半 (防一格输入 10000 字撑爆屏).
  - [`.editor-slot`](src/components/ExcelViewer.vue) + [`.rxl-editor-slot`](src/react/excel-viewer.css) CSS `overflow: hidden → visible`, 让编辑器可溢出原格. z-index:6 仍最上层, 不影响下方网格 / 冻结窗格 / 滚动交互.
  - **提交后行高保持原样** (跟 WPS 一致, 不持久化撑高): 短期视觉浮起仅在编辑期, 提交后单元格还原. 如需永久撑高, 用户单元格设 `wrapText=true` 走已有 autofit (历史行为不变).
- **★ 公式栏 (fx 内容条) 长文本自动撑高 (2026-06-08)** — 用户反馈: 公式栏内容过长被 `text-overflow:ellipsis` 截断, 看不全公式 / 长文本. 修: Vue + React 壳公式栏 `<input>` → `<textarea>` + `auto-resize` (内容变化时 height = scrollHeight). 跟单元格编辑器一样, 普通 Enter 提交、Shift+Enter 插换行、上限 ~6 行 (max-height: 108px) 超过内部滚动. CSS 同步 — formula-bar 高度自动撑, addr / fx 区垂直居中.
- **★ vAlign 溢出 fallback 顶对齐 (跟 WPS 一致, 2026-06-08)** — 用户反馈: 输入长文本提交后, 单元格"显示文末而不是开头". 根因: [canvas-renderer.ts:1090](src/core/render/canvas-renderer.ts#L1090) 默认 vAlign='bottom' 时, 文本总高超过单元格高 → `startY = y + h - pad - totalH` 变负, 首行画到格外, 用户看到的是最后几行. 修: 当 `totalH > availH` 时, 强制走 `'top'` 分支 (显示文头, 末尾裁切). 跟 WPS / Excel 行为一致.
- **`CellEditorFactory` 返回类型扩展** — 旧 `HTMLElement | { el, destroy? }` → 新增可选 `{ el, destroy?, getDesiredHeight?(w) }`. 旧自定义编辑器 100% 向后兼容 (不返 `getDesiredHeight` = 高度锁单元格高, 老行为).
- **`CellEditorContext.reposition()`** — 给自定义编辑器: 主动通知 host 重测撑高. 内容变化后调一次. 不调 = 不撑 (老行为).

### 测试
- 新单测 [default-editor.test.ts](src/core/edit/__tests__/default-editor.test.ts) 10 项: textarea 类型 / commit-cancel 行为 / Enter 移动 / Shift+Enter 换行 / 失焦提交 / 样式贴合 / `getDesiredHeight` 钩子存在 / 兜底 (依赖 canvas measureText 的具体撑高数测试走 e2e 覆盖)
- 新 e2e [edit-long-text.e2e.ts](e2e/edit-long-text.e2e.ts) Vue + React 双覆盖 12 项: 编辑器 textarea / 短文本不撑大 / 长文本撑高 / 动态变化 / Esc 取消 / Shift+Enter 换行 + Enter 提交
- devDependency 加 `jsdom` (单元格编辑器单测依赖 DOM 环境;**仅 dev**, 不影响 dist)

---

## [1.2.0] - 2026-06-08

**主线**: 只读边界三件套(白名单 / 尺寸多形态 / 视觉钩子 + permission-denied 事件)+
模板语义重设计(样式捐赠者)+ WPS 单元格内嵌图(展示 / 互转 / 导出往返)+ 富粘贴 +
图片放大下载 + 虚拟空行 + 公式栏可编辑 + 背景/字体色 + 编辑 UX 补齐(合并/粘贴/右键菜单)+
性能 + 导出错误可见性 + 1900/1904 日期单测。**全部向后兼容、默认只读零回归**。

测试基线: 306 单测 / 107 e2e。

### 修复 + 新增
- **★ 只读边界三件套(2026-06-08;三次提交)** —— 用户三连问"可编辑是一切改变的基础", 怀疑有路径绕过 isEditable. 三阶段解决:
  - **Phase A 闸门补漏 + `permission-denied` 事件**(commit 57ebdbc):审计找出 4 个绕过点 — `pasteRich` 合并粘贴 / `pasteRich` 图片粘贴 / `mergeCells/unmergeCells` 没逐格检查 / 图片转换族 没检查目标格.全部修复.行为 = 默认 skip(跟 `editRange` 一致)+ 一次操作结束 emit 一次 `permission-denied: { reason, cells[], dims?, message? }`(8 种 reason).Vue `@permission-denied` / React `onPermissionDenied`.右键菜单 disabled 完善:粘贴/清除/换行 = 选区有至少 1 格可编辑;合并/拆分/删除行列 = 全部可编辑.新 helper `partitionByEditable` / `rangeAllEditable` / `collectDeniedInRange`.
  - **Phase B 尺寸 API 多形态 + `strictDimensions`**(commit fac634d):用户问"现 API 是否能设置不相邻一批列宽行高".新类型 `DimTarget = number | number[] | { from, to }` 自动识别.升级 `setColumnWidth(target, width)` / `setRowHeight(target, height)` 接 union,**返回值 boolean → number 成功条数(BREAKING)**;新增 `autoFitColumns/Rows(target?)` / `resetColumnWidth/RowHeight(target)` 公开 API(target 不传 = 整表).多 index 时聚合成单次 `restore-wb` undo.新 prop `:strictDimensions` 默认 `false`:`true` 时该列/行至少 1 格在白名单内才能改尺寸,跟"白名单未覆盖 = 完全只读"严格语义一致.新 helper `canEditDimension` / `normalizeDimTarget`.
  - **Phase C 只读视觉钩子 + cursor**(本次):`CellStyleFn` 签名加可选第 3 入参 `ctx: CellStyleCtx { editable: boolean }`,旧 `(cell, pos) => ...` 完全兼容.新 prop `:readOnlyCellStyle: boolean | CellStyleOverride | CellStyleFn` — `true` 套内置浅灰 `#f5f7fa`,对象 = 固定样式,函数 = 按格自定义,默认 `false` 无视觉差异.canvas-renderer 渲染顺序:base style → cellStyle(传 ctx)→ 只读格再叠 readOnlyCellStyle.控制器注入 `isEditable` 回调给 renderer;鼠标悬停只读格(且编辑模式开)自动 `cursor: not-allowed`.Demo 加「高亮只读」toggle 跟「设置可编辑」配合使用.
  - **测试**:permissions.test.ts 共 27 项(11 项白名单 + 4 项 Phase A helpers + 7 项 Phase B canEditDim/normalizeDim);新单测 `readonly-style.test.ts` 7 项;3 个新 e2e spec — `permission-denied.e2e.ts` 4 项 / `dimension-targets.e2e.ts` 6 项 / `readonly-visual.e2e.ts` 3 项.总 281 单测 + 101 e2e 全绿.
- **★ 可编辑白名单 `editableTargets`(2026-06-08)** —— 现有的 `cellReadOnly` / `readOnlyRanges` 是**黑名单**(默认全可编辑,标只读),新加一个 **白名单**(默认全只读,标可编辑)。用户原话:"设置不相邻的一批单元格可编辑,暴露的 API 应该支持多种参数,可以传区域,也可以传单个单元格,多个单元格,行,列;没有设置的就是只读"。
  - **新类型** `EditableTarget`(`core/edit/types.ts`)联合 4 种形状,**自动识别**带哪些字段:`{row,col}` 单格 / `{row}` 整行 / `{col}` 整列 / `MergeRange` 矩形。单值或数组都接,**允许不相邻**多 target
  - **新 prop** `:editableTargets` —— `undefined` (不传) = 不启用白名单 = 老行为;`[]` (显式空) = 全只读;非空 = 白名单生效。优先级:`editable=false` ► 不在白名单 ► `readOnlyRanges` ► `cellReadOnly` ► 否则可编辑(白名单内仍能被黑名单二次"黑"掉)
  - **新命令式 API** `viewer.setEditableTargets(targets)` / `viewer.getEditableTargets()` —— 运行时改不动 prop,直接覆盖 `editCfg.editableTargets`,立即重绘
  - **demo 加「设置可编辑」按钮 + 弹窗**(编辑模式下出现):12×8 网格化点选,可单击格 / 列标题 / 行号切换;footer 显示已选数;「应用」生效 / 「关闭白名单」恢复默认 / 「取消」放弃
  - **测试**:`permissions.test.ts` 加 11 项白名单分支(单格 / 多格不相邻 / 整行 / 整列 / 矩形 / 空数组 / 混合 4 种 / 与黑名单叠加 / editable=false 时白名单也无效 / 形状识别);新 `e2e/editable-targets.e2e.ts` 3 用例(命令式 setEditableTargets 4 种 target 形状 / Demo 弹窗点选→应用 / 空数组全只读)
  - 不破坏现有 API,完全 additive
- **★ 模板语义彻底重设计(2026-06-08)—— 从"占位符 + 锚点表"改为"样式捐赠者"** —— 之前几轮"占位符 + 锚点表"语义反复改: `discardUnmatched` → `trimUnused` → JSON 自然位置… 都没切中用户实际需求。用户原话:**"模板这个功能,只在 json/csv 这些本身不附带格式的数据源才生效,xlsx 数据源根本不需要"**。新语义:
  - **`:templateFile` = 样式捐赠者** —— 模板贡献 styling(`styles` 池 / `merges` / 列宽 / 行高 / `freeze` / `theme`),JSON / CSV 数据**在 A1 自然位置渲染**,模板的 raw 文字 / 占位符 / 图 / 图表 / 条件格式 / 数据验证 **全部丢弃**(避免幽灵规则)
  - **仅在 `:workbook` (JSON / 模型) 数据源下生效** —— xlsx 数据源(`:src`)自带格式,给 `:templateFile` 会被**忽略并 console.warn**;工具栏「模板 ▾」在 xlsx 模式下**禁用**
  - **API 改动(BREAKING)**:删除 `props.template` (`TemplateFillSpec`)、删除 `viewer.applyTemplate(spec)` 命令式 API、删除 `core/template/fill.ts` (`fillTemplate` / `replacePlaceholders` / `TemplateAnchor` / `TemplateFillSpec`) 整个模块。新增 `core/template/style-overlay.ts` 的 `applyStyleTemplate(dataWb, templateWb)` 纯函数(同步,无 onProgress / signal)
  - **核心契约 9 项单测**(`style-overlay.test.ts`):数据 raw 全部保留在自然位置 / 模板 raw 全部丢弃 / styleId 从模板同位置取 / merges + 列宽 + 行高 + freeze 全部从模板拷贝 / 模板的 images + charts + conditional + dataValidations 不带过来 / 数据 sheet 名 + date1904 + cellImages 透传 / 入参不被原地修改 / 空数据 + 模板 → 干净样式骨架 / 数据 dimension 跟模板列宽行高声明取大
  - **demo 简化**:`App.vue` 删 `templateSpec` 计算属性(placeholders + anchors 整套不要),`loadJsonSample` 只传 `:workbook` 数据数组;`:fileName="'订单数据"`。点工具栏「模板 ▾ → 导入 .xlsx 模板」即套样式,模板的标题/客户/合计/{{占位符}} 全部不见, JSON 5 条数据仍在 A1 起的自然位置
  - **e2e 全重写**(`template-switch.e2e.ts`):3 个用例 —— ① JSON 无模板时数据在 A1 自然位置 ② JSON + 模板时 模板装饰文字全部不见 + JSON 仍在 A1 + 切换/清除模板正常 ③ xlsx 数据源加载后工具栏「模板 ▾」禁用
  - **过渡 / 后续迁移**:用户用旧 `applyTemplate(spec)` 自填数据的场景,改为"前端构造好完整 JSON 后再传 `:workbook`"。后续如需"模板填值"那套(不入命令栈的占位符 / 锚点表写入),应该走"应用层"而非组件 prop,组件只管渲染样式
- **右键菜单全面开放(Plan C)** —— 1.2.0 前右键菜单是内置 hardcoded、无任何对外接口。此版三层开放:
  - **prop `:contextMenu`** —— `false` 关闭内置弹层(事件仍触发,供自渲染);`(ctx, items) => MenuItem[] | undefined` transform 回调,在内置 items 上加 / 减 / 重排
  - **事件 `@before-context-menu` / `@context-menu`**(Vue) / `onBeforeContextMenu` / `onContextMenuShow`(React)—— `payload.preventDefault()` 接管渲染(用 Element Plus / Radix 等自渲染);`@context-menu` 拿到 `{x, y, ctx, items}` 总会触发
  - **命令式 API** `openContextMenu(x, y, items?)` / `closeContextMenu()` —— 键盘 Shift+F10、工具栏触发等
  - **插件贡献** `definePlugin({ contextMenu: (ctx, items) => ... })` —— 多插件按数组顺序串行,组件 prop 最后覆盖
  - 顺序固定:**内置 → 插件 → prop → 事件**;`MenuItem` / `ContextMenuCtx` / `ContextMenuTransform` 全部从 `/core` 导出
  - 顺手修了 Vue 4-pkt 接 prop 的"boolean prop 缺省被 Vue 判成 false"小坑(withDefaults 显式 undefined)
- **内置导出进度遮罩 + 三层覆盖机制(P1.5)** —— P1 已建好 `onProgress` + `AbortSignal` 协议,但壳没接 UI,用户调 `viewer.downloadPdf()` 看不到任何反馈 → 这次补 **Shell 默认 UI** + 用户**逐层覆盖**:
  - ① 新组件 `ExportProgressOverlay`(Vue SFC + React tsx)居中模态:stage 标签 + 进度条 + 取消按钮;同视觉 / 同协议
  - ② 壳自动 wrap 长任务(`downloadPdf`/`exportPdf`/`downloadImage`/`exportImage`/`downloadXlsx`/`exportXlsx`/`print`/`convertImagesInRangeToCell`/`convertCellImagesInRangeToFloat`):建内置 `AbortController` + 接 `onProgress` → 用户调时**默认看见遮罩**,无需任何 prop
  - ③ 用户传 `{ onProgress, signal }` 仍正常**链回调**(并存)
  - ④ 覆盖:`:export-progress="false"` 关闭内置遮罩(纯回调);Vue `#export-progress` 插槽 / React `renderExportProgress` 自渲染(拿到 `{state, busy, cancel}`)
  - ⑤ **修单表 PDF 卡 0% bug**:核心导出对"单表"和"jsPDF/canvasToBlob 黑盒"阶段改 emit `ratio: undefined`,overlay 走 indeterminate 扫动条动画(看着在动);多表仍按 `i/total` 走离散进度
  - ⑥ 1.2.0 起 `convertImagesInRangeToCell` / `convertCellImagesInRangeToFloat` 在壳侧返 `Promise<number>`(为接遮罩),core 内核仍同步
- **JSON 直渲(P3)** —— 新 prop `:workbook` 接 `WorkbookModel | JsonInput`(优先于 `:src`),绕过 parser 直接构造模型渲染。三种 JsonInput shape 自动识别:① 二维数组(首格 A1) ② 对象数组(首行表头 = keys) ③ `{sheets:[{name,rows,...}]}`。类型自动推断:数字字符串 → number、`TRUE`/`FALSE` → boolean、ISO 日期串 → Date(可关 `:jsonOptions="{ autoInfer: false }"`)。新公开导出 `jsonToWorkbook` / `isWorkbookModel` 给"仅引擎"用户。
- ~~**模板填值(P3)** — `props.template` (`TemplateFillSpec`) + `applyTemplate(spec)`~~ — **已废弃**, 改为本版 1.2.0 顶部的"模板样式 overlay"语义. 见上文 ★ 重设计条目.
- **图片浮动 ⇄ 嵌入 选区批量 + 工具栏入口** —— 新 API `convertImagesInRangeToCell(range)` / `convertCellImagesInRangeToFloat(range, size?)` 把"选区内"批量互转,聚合成单次撤销(复用 `convertImagesToCells` 范式 + 新增 `convert-to-floats` exec kind)。Vue 工具栏新内置 `image-tools` 下拉(选区/整表/整列 浮动 ⇄ 嵌入)发现性翻倍;右键菜单多格选区时直接出现「选区浮动图全部嵌入(N 张) / 选区内嵌图全部浮动化(N 张)」。
- **Cell Inspector(单元格全息体检)** —— 新 API `inspectCell(row,col)`,在 `getCellSnapshot` 之上聚合 **合并区**(`merge`/`isMergeAnchor`)、**覆盖到该格的浮动图**(`floatingImages[]`,从 `sheet.images[].from..to` 反推)、**WPS 内嵌图 DISPIMG**(`cellImage`)、**数据验证范围命中**(`dataValidation`)、**条件格式命中**(`conditional[]`,含规则索引 + 等效样式;`ConditionalEngine.inspectHits` 共享 evaluator,跟渲染层同一份)、超链接、批注。框架无关 `src/core/model/inspect.ts`,headless 流程也可调。
- **长任务进度回调 + AbortSignal 取消** —— 所有耗时操作(`exportImage` / `exportPdf` / `print` / `exportXlsx`)统一接 `onProgress?: (p: ExportProgress) => void` + `signal?: AbortSignal`。串行多表导出 + 每页/每表 emit `{stage,sheetIndex,ratio,label}` + `await yieldToEvent()` 让出 UI(防假死)。任意时刻 `ctrl.abort()` → 下一个调度点抛 `DOMException('Aborted','AbortError')`(标准语义),上层用 `e.name === 'AbortError'` 区分取消与真错。新 helper `src/core/export/abort.ts`(`yieldToEvent` / `checkAborted` / `isAbortError`)。
- **自动换行(WPS 风格 toggle)**:工具栏内置 `wrap-text` + 右键菜单「自动换行」+ 命令式 API `toggleWrapTextOnSelection()` / `getSelectionWrapState()`(`'all'/'none'/'mixed'`)。行为对齐 WPS:选区**全已换行 → 全部关掉**;否则**全部打开**;mixed → all。行高自动按内容重撑(失效 autofit 缓存触发),延续"只扩不缩"语义。入命令栈单次撤销 style。两壳同构(Vue 工具栏 `wrap-text` builtin + 右键;React 通过命令式 API + 右键)。
- **HiDPI/系统缩放 canvas 对齐**:Windows 125%/150% 缩放、浏览器 Ctrl+缩放(`devicePixelRatio≠1`)下,canvas 作为"替换元素"默认会按 `width*dpr` 显示,导致整个网格被放大、与 DOM 叠加层(浮动图/图表/HTML 文本框)及鼠标命中错位(越往右下偏得越多)。修复:`render()` 显式把 canvas 的 CSS `width/height` 钉成 view 逻辑尺寸,缓冲仍是 dpr 倍(高清),被浏览器降采样显示 → 像素与逻辑坐标 1:1 对齐。
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
