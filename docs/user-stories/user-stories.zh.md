# PPTX Subagent —— 用户故事

> **项目：** opencode.ai / pptx-subagent
> **日期：** 2025 年 6 月
> **范围：** 4 个 Epic，20 个 Story
> **作者：** 创始人（OpenCode 用户）
> **来源：** `chenyu-user requirement.html`（导出为 `.md`）—— 纯需求文档，已剥离 HTML/CSS/JS 模板代码。

---

## 目录

- [Epic 1：模板抽取与模板化](#epic-1模板抽取与模板化)
- [Epic 2：幻灯片生成](#epic-2幻灯片生成)
- [Epic 3：模板扩展](#epic-3模板扩展)
- [Epic 4：工程基础](#epic-4工程基础)
- [参考：拟定的 JSON Schema](proposed-json-schema.md)

> **Epic 全景图（修订）：** 待办列表由 6 个 Epic 重组为 4 个，按生命周期角色分组（一个 Epic ≈ 一个技能）。旧 → 新：**Epic 1 → 1**（抽取引擎）；**Epic 2 → 并入 1**（页眉/页脚与最佳实践是抽取期的 schema 增补，并非独立的技能/价值流）；**Epic 3 → 并入 1**（模板化技能是抽取能力的产品化）；**Epic 4 → 2**（幻灯片生成）；**Epic 6 → 3**（模板扩展）；**Epic 5 → 4**（工程基础）。
>
> *Story ID（如 `US-2.1`、`US-6.1`）是不可变标识，保留原始前缀作为历史标签 —— 它们不再与所在 Epic 的编号一致。同样的旧→新映射见 GAP-ANALYSIS。*

---

## Epic 1：模板抽取与模板化

生命周期的*摄入/理解*侧：读取任意源 `.pptx`，解析其幻灯片母版/版式/主题，并将这种理解以内嵌 JSON schema（`ppt/template_schema.json`）的形式打包回文件，产出一个自描述、可复用的"已模板化" PPTX。本 Epic 汇聚了抽取引擎（US-1.x）、schema 增补如页眉/页脚与最佳实践检测（US-2.x），以及模板化技能的交互体验（US-3.x）—— 由 `generate-template-skill` 承载的单一能力。*（原散落于旧 Epic 1、2、3。）*

---

### US-1.1 [Epic 1] —— 将幻灯片母版抽取为结构化 JSON `[Must Have]`

**作为**一名以创始人身份工作的 OpenCode 用户，
**我希望**子代理（subagent）能读取我上传的任意 PPTX 文件，解析其幻灯片母版和所有幻灯片版式，并将它们转换为结构化的 JSON 表示，
**以便**我获得一份机器可读的模板蓝图，而不再依赖不可预测的纯 LLM 生成。

**细节：**
系统必须使用脚本化技能（而非原始 LLM）打开 PPTX 压缩包，定位 `ppt/slideMasters/` 和 `ppt/slideLayouts/` 中的 XML 文件，并将 OOXML 解析为拟定的 JSON schema。这确保了确定性、可重复的抽取。

**验收标准：**
- [x] 子代理接受 .pptx 文件输入，且对任何合法的 PPTX 都不会崩溃。
- [x] 幻灯片母版 XML 被解析；母版下的每一个版式都被抽取。
- [x] 输出是一个符合拟定 schema 的合法 JSON 对象。
- [x] 抽取由 Python/Node 脚本执行 —— 而非由 LLM 猜测。

**标签：** extraction, slide-master, ooxml, deterministic

---

### US-1.2 [Epic 1] —— 归一化多边形定位 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**抽取出的 JSON 中每个组件都有一个 `polygon` 字段，定义为恰好 4 个坐标对的数组，按逆时针顺序排列，所有值归一化到相对于幻灯片尺寸的 0.0–1.0 范围，
**以便**每个元素的位置和尺寸都明确、与分辨率无关，并且能映射回任意幻灯片尺寸（例如 16:9、4:3）。

**细节：**
- 坐标顺序：左上 → 右上 → 右下 → 左下（逆时针）。
- 取值为 0.0 到 1.0 之间的浮点数，其中 (0,0) 是幻灯片左上角，(1,1) 是右下角。
- 非矩形形状（如梯形、箭头）必须使用其实际的多边形顶点，同样需归一化。
- 幻灯片的原始宽高比记录在 `template_metadata.slide_dimensions` 中，以便子代理在写回时进行反归一化。

**验收标准：**
- [x] 每个组件都有一个 `polygon` 数组，对于矩形形状恰好包含 4 个 `{x, y}` 对象。
- [x] 所有 x 和 y 值都在 [0.0, 1.0] 范围内。
- [x] 脚本中通过简单的叉积检查验证逆时针绕向。
- [x] 幻灯片尺寸（EMU、英寸和宽高比字符串）记录在元数据中。

**标签：** polygon, normalized-coords, positioning

---

### US-1.3 [Epic 1] —— 组件类型枚举 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**每个组件都有一个取自固定枚举的 `type` 字段，该枚举直接映射到 PowerPoint 元素类型（例如 `textbox`、`image`、`table`、`video`、`shape`、`chart`、`group`、`smartart`、`placeholder`），
**以便**幻灯片生成技能确切知道为每个组件创建哪种 OOXML 元素，避免模糊或错误的元素类型。

**细节：**
- 枚举在 schema 中定义，并在技能的 README 中文档化。
- 含 `<p:txBody>` 的 OOXML `<p:sp>` 元素 → `textbox`。
- `<p:pic>` → `image`。
- 含表格的 `<p:graphicFrame>` → `table`。
- 含图表的 `<p:graphicFrame>` → `chart`。
- 含预设几何但无文本的 `<p:sp>` → `shape`。
- `<p:grpSp>` → `group`（带嵌套 `children` 数组）。
- 若脚本遇到无法识别的元素，回退为 `shape` 并设置 `type_confidence: "low"`。

**验收标准：**
- [x] `type` 字段始终存在，且始终是已定义枚举值之一。
- [x] 没有任何组件的 `type: null` 或 `type: "unknown"`。
- [x] 技能源码中包含从 OOXML 标签到枚举值的映射表。

**标签：** type-enum, ooxml-mapping, component

---

### US-1.4 [Epic 1] —— 字体检测与可用性检查 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**抽取过程为每个含文本的组件捕获字体元数据 —— 包括 `family`、`size_pt`、`weight`、`color`、`alignment` 以及一个 `is_available` 布尔值 —— 并在某字体不属于 PowerPoint 内置默认字体时填充一个顶层 `missing_fonts` 数组，
**以便**我确切知道我的模板依赖哪些字体，并能在生成幻灯片前安装它们，防止字体替换导致的版式错乱。

**细节：**
- 内置 PowerPoint 字体（Calibri、Arial、Times New Roman 等）视为始终可用。
- 自定义字体获得 `is_available: false`，并出现在 `missing_fonts` 中（若已知则附带下载建议）。
- 每个字体条目包含一个 `fallback` 字段，建议最接近的内置替代字体。
- 单个文本框内的多个字体 run（如粗体标题 + 常规副标题）作为 `runs` 数组捕获。

**验收标准：**
- [x] 每个文本框组件都有一个包含所有指定字段的 `font` 对象。
- [x] 当所有字体均为内置时，`missing_fonts` 数组为空。
- [x] 当发现非内置字体时，子代理打印面向用户的警告并列出它们。
- [x] `fallback` 始终是内置字体名。

**标签：** fonts, availability, fallback

---

### US-1.5 [Epic 1] —— JSON 存储于 PPTX 压缩包内 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**生成的 JSON 模板存储在 PPTX 压缩包内、类似 `ppt/template_schema.json` 的路径下 —— 且不破坏 PPTX 文件或导致 PowerPoint 拒绝它，
**以便**JSON 随文件本身携带，幻灯片生成技能可直接从任何我提供的 PPTX 中读取，而无需单独的数据库或文件系统。

**细节：**
- 只要 `[Content_Types].xml` 未被修改，PowerPoint 会忽略压缩包内的未知文件。
- 脚本在不改动压缩包内任何现有条目的前提下追加该 JSON。
- 在嵌入前后用 PowerPoint 打开该文件，行为完全一致。
- JSON 经过压缩（minify）以将文件大小影响降至最低（通常 < 50 KB）。

**验收标准：**
- [x] 嵌入后，PPTX 在 PowerPoint 中打开不会报错或出现修复提示。
- [x] 该 JSON 可通过在已知路径重新读取压缩包取回。
- [x] 现有的幻灯片内容、版式和媒体均未被改动。
- [x] 文件大小增量会记录给用户。

**标签：** zip-embedding, pptx-safe, portability

---

### US-2.1 [Epic 1] —— 页眉与页脚检测 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**子代理检测我的幻灯片母版是否包含页眉和页脚区域 —— 如果没有，在继续之前明确询问我是否要添加它们，
**以便**我生成的幻灯片不会缺少让一份演示文稿看起来未完成或不专业的标准结构元素。

**细节：**
脚本检查幻灯片母版 XML 中是否存在页眉（`<p:hdr>`）和页脚（`<p:ftr>`）元素。若缺失，子代理暂停抽取并向用户返回一个结构化提示。

**验收标准：**
- [x] `template_metadata.header_footer.has_header` 和 `.has_footer` 是反映实际检测结果的布尔值。
- [x] 当两者均为 `false` 时，子代理在继续前输出一个面向用户的提问。
- [x] 如果用户回答"是的，添加页眉"，脚本将一个默认页眉区域注入到 JSON 中（暂不注入 PPTX —— 仅注入 schema）。

**标签：** header, footer, detection, prompt

---

### US-3.1 [Epic 1] —— 端到端模板生成流水线 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**调用一个"生成模板（generate template）"技能，它接收我的 PPTX 文件、运行完整的抽取流水线（Epic 1）并产出 JSON schema，
**以便**我无需手动编辑任何 JSON 即可获得一份标准化、可复用的模板定义。

**细节：**
此技能是入口点。用户说类似"从这个 PPTX 抽取模板"的话，子代理路由到本技能。该技能编排：读取压缩包 → 解析 XML → 构建 JSON → 校验 → 嵌入 → 返回文件。

**验收标准：**
- [x] 技能可通过自然语言意图检测调用（无需特殊命令）。
- [x] 完整流水线端到端运行，无需手动中间步骤。
- [x] 校验错误（如未找到幻灯片母版）会清晰地上报给用户。

**标签：** skill, template-generator, pipeline

---

### US-3.2 [Epic 1] —— 模板命名 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**生成的 JSON 在顶层包含一个 `title` 字段为模板命名（如"Q3 投资人路演 —— 暗色主题"），
**以便**当我有多个已模板化的 PPTX 文件时，可以按名称查询或引用它们，且子代理知道我指的是哪个模板。

**细节：**
标题从 PPTX 文件元数据（文档标题属性）推断，或从第一张幻灯片的标题文本推断，或 —— 若两者都不存在 —— 子代理提示用户命名。

**验收标准：**
- [x] `template_metadata.title` 始终是一个非空字符串。
- [x] 推断顺序：core.xml 标题 → 第 1 张幻灯片标题文本 → 用户提示。
- [x] 抽取后标题会展示给用户确认。

**标签：** metadata, naming, queryable

---

### US-3.3 [Epic 1] —— 返回可下载的已模板化 PPTX `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**技能返回一个新的（内嵌 JSON 的）PPTX 文件供我下载并保存，
**以便**我拥有一个可移植、自描述的模板文件，可跨会话复用或与团队成员共享。

**细节：**
返回的文件是原始 PPTX 加上 `ppt/template_schema.json`。子代理提供一个下载链接和一份抽取摘要（版式数、组件数、字体等）。

**验收标准：**
- [x] 通过 OpenCode 的文件输出机制提供可下载的 PPTX。
- [x] 向用户打印一份人类可读的抽取摘要。
- [x] 文件通过往返测试：在 PowerPoint 中打开、重新上传、重新抽取 → JSON 完全一致。

**标签：** output, download, round-trip

---

### US-3.4 [Epic 1] —— 主题与颜色抽取 `[Should Have]`

**作为**一名 OpenCode 用户，
**我希望**模板 JSON 还能从 PPTX 的 `ppt/theme/theme1.xml` 中捕获颜色主题（primary、secondary、accent、background）和字体调色板，
**以便**当幻灯片生成技能创建新内容时，它可以使用确切的品牌色和排版而无需猜测。

**细节：**
主题颜色从 `<a:clrScheme>` 抽取并映射到语义角色。字体调色板来自 `<a:fontScheme>`。

**验收标准：**
- [x] `theme` 对象包含 `primary_color`、`secondary_color`、`accent_color`、`background_color` 作为十六进制字符串。
- [x] `theme.font_palette` 包含 `heading`、`body`、`accent` 字体名。
- [x] 若 theme1.xml 缺失或格式错误，使用合理的默认值并显示警告。

**标签：** theme, colors, branding

---

## Epic 2：幻灯片生成

*生成*侧：消费一个已模板化的 PPTX，产出全新的、符合品牌调性的幻灯片 —— 读取内嵌 JSON 以选择版式、文本自适应、批量生成、多宽高比输出，以及在无 JSON 时自动链式抽取。由 `generate-slide-skill` 承载。*（原 Epic 4。）*

---

### US-4.1 [Epic 2] —— 读取内嵌 JSON 作为版式参考 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**幻灯片生成技能从我提供的已模板化 PPTX 中读取 `ppt/template_schema.json`，并将其作为权威的版式参考，
**以便**每张生成的幻灯片都把内容放在正确的位置、使用正确的样式，而不是依赖 LLM 对坐标的臆造。

**细节：**
该技能**从 zip 中**读取 JSON（不重新抽取或重新解析 PPTX 的 XML），根据用户意图（如"标题页"、"内容页"、"双栏"）通过 `layout_name` 匹配确定使用哪个幻灯片版式，并**使用幻灯片母版自带的版式**（`add_slide(layout)`）生成新幻灯片——**不是**在多边形坐标处手动放置 OOXML 元素。内嵌 JSON 是模板的忠实、可移植描述（版式名、组件类型、字体、主题、归一化位置），驱动版式选择与一致性；模板版式本身承载实际的定位与继承的样式（项目符号、主题、母版默认值）。归一化的 `polygon` 坐标（US-1.2）仍作为忠实的几何描述保留，可用于可选的一致性/校验检查——它**不是**放置数据源。

**验收标准：**
- [x] 技能从压缩包读取 JSON —— 不重新抽取或重新解析 XML。
- [x] 版式选择基于 `layout_name` 匹配或用户确认。
- [x] 生成的幻灯片使用模板自带的版式（通过 `add_slide`）；内嵌 JSON 驱动版式选择，而不是在多边形坐标处放置元素。（多边形保真度一致性检查为可选、非致命。）

**标签：** slide-generation, layout-matching, embedded-json

---

### US-4.2 [Epic 2] —— 视觉美观的输出与文本适配 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**生成的幻灯片视觉美观 —— 文本适配得当（无溢出）、字体大小适合内容长度、间距一致 —— 同时仍遵守模板定义的区域，
**以便**输出看起来像人工设计的幻灯片，而非把原始数据堆进文本框。

**细节：**
- 脚本实现文本适配逻辑：若文本在模板字号下超出了文本框的多边形区域，则按步长（如 -2pt）缩小字号直到适配，并设最低下限。
- 行距和段距从模板现有的文本 run 推导。
- 项目符号若检测到则使用模板的项目符号样式；否则应用一个干净的默认样式。

**验收标准：**
- [ ] 没有任何生成的幻灯片出现文本溢出其边界框的情况。
- [x] 字号仅在必要时缩小 —— 绝不超过模板定义的尺寸。
- [x] 发生自动调整大小时，组件上会设置一个 `font_size_adjusted` 标志。

**标签：** text-fitting, visual-quality, auto-sizing

---

### US-4.3 [Epic 2] —— 无 JSON 时自动链式抽取 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**能把一个尚未内嵌 JSON 的 PPTX 交给子代理，并说"先从这个抽取模板，然后用它生成幻灯片" —— 全在一次交互中完成，
**以便**当我第一次处理一个新文件时，不必运行两个单独的命令。

**细节：**
子代理检测到 `ppt/template_schema.json` 不存在，自动先链式调用模板生成器技能，然后继续幻灯片生成。用户会被告知这是两步过程，但无需手动触发每一步。

**验收标准：**
- [x] 单个用户提示即可无错误地依次触发两个技能。
- [x] 中间 JSON 被嵌入到输出 PPTX 中。
- [x] 用户看到状态消息，如"未找到模板 —— 先抽取，再生成幻灯片……"。

**标签：** chaining, auto-detect, ux

---

### US-4.5 [Epic 2] —— 多幻灯片批量生成 `[Could Have]`

**作为**一名 OpenCode 用户，
**我希望**幻灯片生成器支持从单个提示批量生成多张幻灯片 —— 例如"创建一份 10 页的投资人路演，包含市场、问题、解决方案、团队和财务等章节"，
**以便**我可以一次性生成一份完整演示文稿，而无需逐张创建幻灯片。

**细节：**
LLM 先规划幻灯片顺序和内容大纲（作为结构化数组），然后脚本遍历该规划，为每张幻灯片选择合适的版式并填充内容。进度指示器显示完成情况（如"第 3/10 页"）。

**验收标准：**
- [x] 单个提示可在一个 PPTX 中产出 2–20+ 张幻灯片。
- [x] 每张幻灯片为其内容类型使用正确的版式。
- [x] 生成过程中向用户报告进度。
- [x] 生成开始前向用户展示 LLM 大纲，并可选择编辑。

**标签：** batch, multi-slide, outline

---

### US-4.6 [Epic 2] —— 多宽高比渲染 `[Should Have]`

**作为**一名 OpenCode 用户，
**我希望**生成与模板不同尺寸或宽高比的演示稿(例如从 16:9 模板渲染出 4:3 的演示稿),且每个元素——文本框、图片、形状——都按比例缩放到新尺寸,
**以便**我能在多种输出格式(16:9、4:3、方形)间复用同一份模板,而不必重新设计模板或得到错位的布局。

**细节：**
当目标幻灯片尺寸与模板原生尺寸不同时,幻灯片生成技能偏离默认的 US-4.1 路径(`add_slide(layout)`,按模板原生尺寸渲染),改走**坐标放置路径**:读取内嵌 JSON 的归一化 `polygon` 坐标(US-1.2,0.0–1.0),按**目标**幻灯片尺寸反归一化,并在所得 EMU 位置创建 OOXML 元素——使每个元素按比例缩放到新尺寸。这之所以可能,是因为 US-1.2 的归一化坐标模型在设计上就是分辨率无关的。技能会向用户询问(或推断)目标宽高比。由于该路径不使用版式占位符,python-pptx 本会继承的样式(字体、主题色、项目符号)需从内嵌 JSON 的 `theme` 与各组件的 `font` 元数据重新应用。

**验收标准：**
- [x] 给定一份 16:9 的已模板化 PPTX,技能能按请求渲染出等效的 4:3 演示稿(反之亦然),走坐标放置路径。
- [x] 每个元素(文本框、图片、形状)都按比例缩放到新尺寸——无裁剪或错位(在 US-4.2 文本适配容差之内)。
- [x] 归一化的 `polygon` 坐标按**目标**幻灯片尺寸反归一化;所得位置与按比例缩放后的原始位置偏差在 1% 以内。
- [x] 字体/主题/项目符号从内嵌 JSON 元数据重新应用,确保绕过版式继承后输出仍符合品牌。
- [x] 当目标尺寸等于模板原生尺寸时,使用默认的 US-4.1 `add_slide(layout)` 路径(此场景下本故事为 no-op)。

**标签：** multi-aspect-ratio, coordinate-placement, proportional-scaling, resolution-independent

---

### US-4.7 [Epic 2] —— 模板选择与渲染前校验 `[Must Have]`

**作为**一名 OpenCode 用户，
**我希望**引擎在我不指定模板时使用默认模板，接受我在对话中给出的任意 `.pptx` 路径，并在所选模板结构上不可用时（给出清晰错误）拒绝渲染，
**以便**生成永远不会因不可用的模板而静默产出一个损坏的演示稿。

**细节：**
- 默认模板为 `template/default.pptx`（仓库根目录），在未提供 `template_path` 时使用；将其放在顶层（与 `output/` 对称）使其易于查找与编辑。
- 用户提供的模板以**路径**形式传入（`template_path` / CLI `--template`），而非覆盖默认模板 —— 取代了早先的「复制覆盖同一路径」工作流。
- 每次加载（默认或用户提供）都运行一次**预检**：严重问题抛出 `TemplateError` 并在渲染循环前中止；次要问题保持非致命警告。

**验收标准：**
- [x] 未指定模板时，引擎针对 `template/default.pptx` 渲染；用户提供的 `.pptx` 路径经 `template_path` 传入（默认模板绝不被覆盖）。
- [x] 严重模板问题 —— 损坏/非 PPTX、无 slide master、零 layout、或无法服务任一 8 种 slide type —— 抛出清晰的 `TemplateError`，而非产出损坏的演示稿。
- [x] 次要问题（缺字体、无页眉/页脚、内容区过小、无内嵌 schema）保持非致命警告；生成照常进行。

**标签：** template, validation, error-handling

---

## Epic 3：模板扩展

*适配*侧：在渲染时，当某张计划幻灯片的类型没有匹配的版式，将一个捐赠版式克隆到派生的 `template_new.pptx`，使幻灯片组仍能渲染 —— 绝不修改用户的原始模板。由 `template-modifier-skill` 承载。*（原 Epic 6。）*

---

### US-6.1 [Epic 3] —— 版式缺失时扩展模板 `[Should Have]`

**作为**一名使用极简或专用自定义模板的 OpenCode 用户，
**我希望**子代理能检测到我的模板缺少某张计划幻灯片所需的版式，并自动将一个克隆版式扩展进一个派生文件（绝不修改我的原文件），
**以便**我的幻灯片组仍能用合适的版式渲染，而不是静默跳过该页或崩溃。

**细节：**
- 作为渲染前的步骤运行：对每张计划幻灯片，检查模板是否提供占位符组成指纹匹配该幻灯片类型的版式。
- 当某幻灯片类型无匹配版式时，将一个捐赠版式（指纹最接近者）经 XML/part 克隆写入派生的 `template_new.pptx`（python-pptx 未公开添加版式的 API），并通过 config override 把该类型钉到克隆版式。
- 基础 `template.pptx` **不可变** —— 绝不写入。克隆只落派生文件。克隆后 reload-verify（新版式须能按名找到）；任何失败均回滚（删除派生文件）并回退基础模板，使幻灯片组仍能渲染。
- 默认情况下，内容超限（正文超出占位符）由 density 模式降档处理，**而非**克隆。为超限内容克隆是可选策略。
- 每当使用了派生模板，强制通知用户所用的模板及原因。

**验收标准：**
- [x] 渲染前，引擎检测到模板中任何版式缺失（无指纹匹配）的幻灯片类型，并将其标记为待扩展。
- [x] 缺失版式被克隆进派生的 `template_new.pptx`；原始 `template.pptx` 绝不被修改。
- [x] 克隆版式经 config override 钉到其幻灯片类型，填充引擎一次性针对活动（基础或派生）模板渲染整套幻灯片组。
- [x] 克隆失败非致命：派生文件被丢弃并改用基础模板；幻灯片组仍能渲染。
- [x] 每当使用派生模板时，用户被通知所用的模板及原因。

**标签：** template-extension, layout-cloning, capability-b, graceful-degradation

---

## Epic 4：工程基础

贯穿三个技能的非功能性基础：CLI 架构与退出码、共享的 JSON-schema 校验契约、以及结构化日志。*（原 Epic 5。）*

---

### US-5.1 [Epic 4] —— 两个独立技能与 CLI 脚本 `[Must Have]`

**作为**一名 OpenCode 用户（以及维护该子代理的开发者），
**我希望**系统恰好暴露 2 个技能 —— "generate-template" 和 "generate-slides" —— 每个都由一个专门的、可独立测试的脚本支撑，
**以便**每个技能都有单一职责，可被单独调试，并遵循 OpenCode 文档化的技能模式。

**细节：**
- **generate-template**：脚本接受 `--input path/to/file.pptx --output path/to/output.pptx`。读取压缩包、抽取、构建 JSON、嵌入、写入新压缩包。
- **generate-slides**：脚本接受 `--template path/to/templated.pptx --prompt "..." --output path/to/deck.pptx`。读取 JSON、生成幻灯片、写入新压缩包。
- 两个脚本都以有意义的退出码退出（0 = 成功，1 = 校验错误，2 = 运行时错误）。

**验收标准：**
- [ ] 每个技能都有各自的目录，包含 `skill.yaml`、脚本和 README。
- [x] 两个脚本均可独立于 LLM 从 CLI 运行。
- [x] 退出码已文档化并被一致使用。

**标签：** architecture, cli, testability, single-responsibility

---

### US-5.2 [Epic 4] —— 用于校验的共享 JSON Schema `[Must Have]`

**作为**一名维护该子代理的开发者，
**我希望**有一个共享的 JSON Schema（JSON Schema draft-07 或 2020-12），两个脚本都用它来校验模板 JSON —— schema 文件随技能一起分发，
**以便**模板生成器的输出能保证可被幻灯片生成器消费，在构建时而非运行时捕获 schema 漂移。

**细节：**
schema 文件（`template_schema.json`）位于一个共享的 `common/` 目录。两个脚本在读或写之前都运行 `validate(json, schema)`。LLM 也在其系统提示中获得该 schema，以便在推理幻灯片内容时理解其结构。

**验收标准：**
- [ ] 存在一个 `.json` schema 文件，并被两个脚本引用。
- [x] 模板生成器在嵌入前校验其输出。
- [ ] 幻灯片生成器在读取前校验 JSON。
- [x] schema 版本在 `template_metadata.schema_version` 中跟踪。

**标签：** json-schema, validation, contract, versioning

---

### US-5.3 [Epic 4] —— 结构化日志 `[Should Have]`

**作为**一名 OpenCode 用户，
**我希望**子代理为每个操作记录结构化输出（JSON lines）—— 包括抽取步骤、版式选择、字体警告和生成进度 —— 输出到一个我可以查看的日志流，
**以便**当出现问题时（如某张幻灯片看起来不对），我能确切追踪脚本做了什么，并分享日志用于调试。

**细节：**
每条日志行是一个 JSON 对象，含 `timestamp`、`level`、`skill`、`action` 和 `details`。日志写入 stderr，以免干扰 stdout 上的文件输出。

**验收标准：**
- [ ] 每个重要操作都发出一条结构化日志行。
- [ ] 日志输出到 stderr；只有文件路径输出到 stdout。
- [ ] 日志级别可通过 `--log-level` 标志控制（debug、info、warn、error）。

**标签：** logging, debugging, observability
