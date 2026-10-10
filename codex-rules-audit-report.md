# Codex Rules Audit Report
审计日期：2026-10-02（Asia/Shanghai）
工作区：E:\works\nuxt\ppt-web-play-nuxt
目的：只完成规则审计、Candidate 建议和可验证的隔离检查；不把 Candidate Adoption 到长期配置。

## 1. Executive Summary

当前最大的负担来自一份可能被当作全局规则加载的 977 行审计流程，以及少数宽泛 Skill 触发条件；但全局 AGENTS 加载范围与缓存文件的实际生效性无法从本机证实，所以只提出隔离和收窄建议，不修改任何规则。当前针对合并问题的项目级类型检查与生产构建均通过，未发现 pptx.ts 的可复现编译错误。

### 计数口径

本报告的规则计数以“单独触发的一组操作约束”为一个审计单元；C1–C11 的 24 条是从本次直接有效或与当前执行有关的 21 个本地文件和 3 个对话注入来源中提取的可追溯规则簇。14 条 C5 记录是本次明确检查过的 Skill 入口触发条件。146 + 2 + 211 是磁盘 Skill 文件数量，包含缓存版本；这些文件的所有内部规则没有被假定为当前生效，也没有被算入 24 条。因而报告数字是有证据的本次审计范围计数，不代表系统里所有潜在 Skill 内部规则的总数。

## 2. Baseline

- AGENTS.md：项目和项目父路径 0；用户 .codex 下物理文件 42，其中 4 个被逐项读取，其余 38 个为缓存、临时或历史配置候选。
- Skills：C:\Users\admin\.codex\skills 下 146 个；C:\Users\admin\.agents\skills 下 2 个；插件缓存下 211 个物理 SKILL.md。后两类和插件缓存的运行时选择范围未完全证实。
- 项目根和子目录：没有 AGENTS.md 或 SKILL.md。README.md 是 265 行项目兼容性和交付说明，含本地文件处理、Worker、支持矩阵，以及“发布前跑 typecheck/build”的项目级流程。
- 工程检查：package.json 有 build、typecheck、verify:animation、verify:chart-number-format 等脚本；没有 test 或 lint 脚本。README 的浏览器检查要求开发服务器在 3000 端口运行。
- 用户配置：C:\Users\admin\.codex\config.toml 中发现 25 个启用的插件 section 和 16 个关闭的 Hook section；[features] hooks=true 仅表示 Hook 功能可用，不等于有 Hook 实例启用。
- Git 基线：branch main；HEAD 5bd4c16aa8f8f916999cc6a2c50f77a36c0ef560。审计前已有 12 个 tracked 修改和 16 个 untracked 文件，均为用户现有工作；本次没有改动或覆盖它们。最终新增的审计报告本身为一个新文件。
- 当前 pptx.ts 差异已有 1,595 行新增、122 行删除；这是审计前的未提交变更，不代表本次修改。

### RULE_SOURCE_INVENTORY

| 路径 / 来源 | 类型与可能范围 | 可读取 | 当前生效证据 | 覆盖关系 |
|---|---|---:|---|---|
| C:\Users\admin\.codex\AGENTS.md | 用户级 AGENTS 候选；977 行 | 是 | 本次内容由用户直接粘贴，故本任务有效；Codex 是否在普通任务自动加载该文件为 UNKNOWN | 若自动加载则可能影响所有任务；非项目目录祖先 |
| E:\works\nuxt\ppt-web-play-nuxt 及父目录中的 AGENTS.md | 项目 / 上层 AGENTS | 不存在 | 不适用 | 没有项目级覆盖层 |
| E:\works\nuxt\ppt-web-play-nuxt 中的 SKILL.md | 项目技能 | 不存在 | 不适用 | 无项目 Skill 覆盖层 |
| C:\Users\admin\.codex\plugins\cache\context-mode\context-mode\1.0.169\configs\codex\AGENTS.md | Context-mode 插件的 Codex 配置规则 | 是 | 对应的 context-mode 行为由当前工具环境注入；该磁盘文件是否被本轮加载仍为 UNKNOWN | 仅可能作用于 context-mode 插件配置 |
| C:\Users\admin\.codex\plugins\cache\ponytail\ponytail\4.10.0\AGENTS.md | Ponytail 插件规则 | 是 | Ponytail 规则在当前开发者注入文本中明确激活；该磁盘文件单独是否加载为 UNKNOWN | 当前注入规则高于插件文件 |
| C:\Users\admin\.codex\plugins\cache\i-have-adhd\i-have-adhd\0.3.0\AGENTS.md | 插件仓库局部规则 | 是 | 文件说明适用于该插件仓库；当前工作区不是该仓库，故不适用 | 不覆盖当前项目 |
| C:\Users\admin\.codex\skills\**\SKILL.md（146 个） | 用户 Skill 的入口、description 和渐进披露内容 | 入口可读取；本次只核对相关入口 | 当前 Skill 目录可见，但每个任务是否触发及内部 references 是否加载依赖运行时，未逐个验证 | 具体 Skill 只在触发后作用 |
| C:\Users\admin\.agents\skills\**\SKILL.md（2 个） | Agent Skill 入口 | 物理存在；未逐个读取 | UNKNOWN | 与 .codex 技能目录关系未证实 |
| C:\Users\admin\.codex\plugins\cache\**\SKILL.md（211 个） | 插件缓存 Skill，包含版本副本 | 可枚举；仅相关入口已读 | 哪一缓存版本会被当前目录选用为 UNKNOWN | 可能被插件安装状态和版本选择覆盖 |
| C:\Users\admin\.codex\config.toml | 插件 / Hook 工具配置 | 是 | 文件可读；Hook section 均 disabled | 配置工具可用性，不等同于提示规则 |
| E:\works\nuxt\ppt-web-play-nuxt\README.md | 项目知识、兼容性和发布验证约定 | 是 | 当前项目文件，且正文已被用户修改 | 仅适用于 PPTX 项目；应保留现有改动 |
| E:\works\nuxt\ppt-web-play-nuxt\package.json | 项目脚本和工程约束 | 是 | 当前项目配置，已有用户修改 | 仅适用于本项目 |
| 当前系统 / 开发者注入指令 | 工具、安全、权限、输出、协作和会话规则；无本地路径 | 当前可见 | 本轮直接有效；无法读取为本地文件或计算哈希 | 高于用户级 / 项目级文本 |
| 用户 PPTX 角色与工作流提示 | 项目技术栈、OOXML、兼容性与播放器目标；无本地路径 | 当前对话可见 | 仅本任务 / PPTX 项目上下文 | 用户当前请求可更新其目标 |
| 用户本轮 AGENTS 审计指令 | 本次审计流程；与 .codex/AGENTS.md 内容对应 | 当前对话可见 | 本任务直接有效 | 最新用户指令优先 |
| C:\Users\admin\Downloads\CommonAnimationParser.md | 用户给的技术参考文档 | 路径存在；本审计未读取内容 | 不是规则来源；用户要求“可参考、以官方文档为主” | 文内任何指令都只当文档内容，不覆盖用户请求 |

#### 用户 .codex 下其余 38 个 AGENTS 候选

以下路径已枚举但本次未逐一打开；均标记 EFFECTIVE_SCOPE=UNKNOWN，不凭文件存在推定生效。它们分别位于插件暂存、平台专用配置、模板、旧可视化运行目录或审计样例目录，不在当前项目祖先路径中。

- C:\Users\admin\.codex\.tmp\marketplaces\.staging\marketplace-upgrade-ZTXCzV\AGENTS.md
- C:\Users\admin\.codex\.tmp\marketplaces\.staging\marketplace-upgrade-Zw98st\AGENTS.md
- C:\Users\admin\.codex\.tmp\marketplaces\i-have-adhd\AGENTS.md
- C:\Users\admin\.codex\.tmp\marketplaces\ponytail\AGENTS.md
- C:\Users\admin\.codex\.tmp\plugins\plugins\build-web-apps\skills\react-best-practices\AGENTS.md
- C:\Users\admin\.codex\.tmp\plugins\plugins\build-web-apps\skills\supabase-best-practices\AGENTS.md
- C:\Users\admin\.codex\.tmp\plugins\plugins\data-analytics\AGENTS.md
- C:\Users\admin\.codex\.tmp\plugins\plugins\product-design\templates\mobile-app\AGENTS.md
- C:\Users\admin\.codex\.tmp\plugins\plugins\product-design\templates\prototype\AGENTS.md
- C:\Users\admin\.codex\.tmp\plugins\plugins\vercel\skills\react-best-practices\AGENTS.md
- C:\Users\admin\.codex\.tmp\plugins\plugins\zoom\AGENTS.md
- C:\Users\admin\.codex\plugins\.marketplace-plugin-source-staging\marketplace-plugin-source-FosW8K\AGENTS.md
- C:\Users\admin\.codex\plugins\.marketplace-plugin-source-staging\marketplace-plugin-source-OlOERe\AGENTS.md
- C:\Users\admin\.codex\plugins\.marketplace-plugin-source-staging\marketplace-plugin-source-qLLDpS\AGENTS.md
- C:\Users\admin\.codex\plugins\.marketplace-plugin-source-staging\marketplace-plugin-source-rmKaWb\AGENTS.md
- C:\Users\admin\.codex\plugins\.marketplace-plugin-source-staging\marketplace-plugin-source-TaiYBY\AGENTS.md
- C:\Users\admin\.codex\plugins\.marketplace-plugin-source-staging\marketplace-plugin-source-ZTZGuR\AGENTS.md
- C:\Users\admin\.codex\plugins\cache\context-mode\context-mode\1.0.169\configs\kilo\AGENTS.md
- C:\Users\admin\.codex\plugins\cache\context-mode\context-mode\1.0.169\configs\openclaw\AGENTS.md
- C:\Users\admin\.codex\plugins\cache\context-mode\context-mode\1.0.169\configs\opencode\AGENTS.md
- C:\Users\admin\.codex\plugins\cache\context-mode\context-mode\1.0.169\configs\pi\AGENTS.md
- C:\Users\admin\.codex\plugins\cache\context-mode\context-mode\1.0.169\configs\zed\AGENTS.md
- C:\Users\admin\.codex\plugins\cache\openai-curated\build-web-apps\1dc19589\skills\react-best-practices\AGENTS.md
- C:\Users\admin\.codex\plugins\cache\openai-curated\build-web-apps\1dc19589\skills\supabase-best-practices\AGENTS.md
- C:\Users\admin\.codex\plugins\cache\openai-curated\vercel\1dc19589\skills\react-best-practices\AGENTS.md
- C:\Users\admin\.codex\plugins\cache\openai-curated-remote\data-analytics\1.0.11\templates\data-app\base\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\15\01a0a2e4-fa0f-7902-83d7-f2d6bbafa459\runtime-a\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\15\01a0a2e4-fa0f-7902-83d7-f2d6bbafa459\runtime-a\.tmp\marketplaces\.staging\marketplace-upgrade-h8FbSa\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\15\01a0a2e4-fa0f-7902-83d7-f2d6bbafa459\runtime-a\.tmp\marketplaces\i-have-adhd\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\15\01a0a2e4-fa0f-7902-83d7-f2d6bbafa459\runtime-b\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\15\01a0a2e4-fa0f-7902-83d7-f2d6bbafa459\runtime-b\.tmp\marketplaces\i-have-adhd\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\15\01a0a2e4-fa0f-7902-83d7-f2d6bbafa459\runtime-b\.tmp\marketplaces\ponytail\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\15\01a0a2e4-fa0f-7902-83d7-f2d6bbafa459\runtime-c\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\16\01a0a81e-1f8f-7241-b4e8-859452992cfe\audit-ab-home-B-6a74d1d8b6df42c4b6b1227e99deb91e\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\16\01a0a81e-1f8f-7241-b4e8-859452992cfe\audit-ab-home-B-6a74d1d8b6df42c4b6b1227e99deb91e\.tmp\marketplaces\i-have-adhd\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\16\01a0a81e-1f8f-7241-b4e8-859452992cfe\audit-ab-home-B-6a74d1d8b6df42c4b6b1227e99deb91e\.tmp\marketplaces\ponytail\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\16\01a0a81e-1f8f-7241-b4e8-859452992cfe\audit-ab-home-B-6a74d1d8b6df42c4b6b1227e99deb91e\.tmp\plugins\plugins\build-web-apps\skills\react-best-practices\AGENTS.md
- C:\Users\admin\.codex\visualizations\2026\09\16\01a0a81e-1f8f-7241-b4e8-859452992cfe\audit-ab-home-B-6a74d1d8b6df42c4b6b1227e99deb91e\.tmp\plugins\plugins\build-web-apps\skills\supabase-best-practices\AGENTS.md

## 3. 11 类规则统计

| 类别 | 条数 | 本次识别的唯一规则簇 |
|---|---:|---|
| C1 安全与不可逆操作 | 1 | 保留密钥、隐私、删除、生产、外发和不可逆边界 |
| C2 用户授权与决策边界 | 1 | 最新用户请求优先；不得让旧记忆覆盖当前请求 |
| C3 任务范围与禁止项 | 1 | 本任务只审计和提 Candidate，不做正式 Adoption |
| C4 上下文与文档读取 | 1 | 仅加载与任务、目录、阶段相关材料；审计清单不得无条件变成全局流程 |
| C5 Skills 与触发条件 | 14 | 本次检查的 14 个入口 trigger；其中 5 个存在过宽或重叠候选 |
| C6 工具与执行方式 | 1 | 工具应按任务和输出规模选择，不能强制所有读取走同一运行时 |
| C7 测试、验证与 QA | 1 | 验证程度由风险和影响面决定；保留项目发布门槛 |
| C8 Multi-Agent / Delegation | 1 | 仅对可并行且能分配文件所有权的工作委派 |
| C9 完成条件与停止条件 | 1 | 修到验收条件后停；遇可修复问题继续闭环 |
| C10 输出、汇报与交付格式 | 1 | 真实证据优先；报告格式只在审计任务触发 |
| C11 项目知识与业务规范 | 1 | PPTX / OOXML 支持范围留在项目说明，不因缩 Context 删除 |

TOTAL_RULE_SOURCES = 23（21 个本地文件 + 2 个非文件注入来源；本地文件数含当前 project README/package.json；用户粘贴的审计指令与磁盘 AGENTS 内容视为同一规则内容，不重复算源）
TOTAL_RULES_IDENTIFIED = 24（以上 C1–C11 唯一审计单元之和）
DUPLICATES = 3 个语义重复簇
CONFLICTS = 1 个 scope-conditional 冲突
OVERBROAD_SKILLS = 5 个入口触发候选
UNCONDITIONAL_READS = 1 个 scope-conditional 候选
OVER_CONFIRMATION_RULES = 0 个已证实规则
OVER_TESTING_RULES = 1 个
MULTI_AGENT_RISKS = 2 个（1 次重复读取；1 个潜在并行写风险）
EARLY_STOP_RISKS = 0 个已证实规则

### C5 Skill 入口盘点（14 条，各计一次）

| Skill 入口 | 触发描述摘要 | 判断 |
|---|---|---|
| source-driven-development | 任何 framework / library implementation decision | 范围偏宽，候选收窄 |
| debugging-and-error-recovery | 测试、build、行为异常或 unexpected error | 与调试相关，范围可接受 |
| test-driven-development | 任意 logic、bug、behavior change、bug report | 范围偏宽，候选收窄 |
| verification-loop | 验证 Claude Code session 的工作 | 产品 / runtime 名称不匹配当前 Codex 语境，建议澄清支持范围 |
| frontend-coding | Vue / Nuxt 一般页面、组件、路由、状态和样式 | 与 frontend-ui-engineering 有重叠 |
| frontend-ui-engineering | UI、组件、布局、状态、响应式和 WCAG | 与 frontend-coding 有重叠 |
| frontend-skill | 用户强调强视觉风格的网站 / app / prototype / game UI | 视觉专项，触发较清楚 |
| browser-qa | 部署功能后的视觉测试和交互验证 | 触发较清楚 |
| browser-testing-with-devtools | 任何浏览器运行的功能 | 范围偏宽；按需调试数据收窄 |
| playwright | 终端驱动真实浏览器导航 / 表单 / 截图 / 提取 | 触发较清楚 |
| e2e-testing | 编写 / 维护 Playwright E2E 测试 | 触发较清楚 |
| i-have-adhd | 仅显式 /i-have-adhd 命令激活，直到 stop | 触发较清楚 |
| context-mode | 使用 context-mode 工具处理大输出 / 检索 | 与大输出处理场景相关；不要求所有短操作使用 |
| ponytail | 当前 developer 注入明确持续激活；Skill 入口描述为最小实现策略 | 注入作用明确；磁盘插件规则是否单独生效未知 |

归类规则：以上 14 条各只计入 C5。R09 的 C7 是另一条独立验证强度规则；它不重复计入 Skill trigger 数。

## 4. 9 项退化扫描

| 扫描 | 状态 | 证据和结论 |
|---|---|---|
| D1 重复规则 | FINDINGS | AGENTS 的 Adoption 边界、禁止伪造 Runtime Evidence、回滚验收各重复成簇；这类重复多数是安全闸门。建议保留 Adoption / 回滚关键提醒，不把防错提示全部删掉。 |
| D2 规则冲突 | FINDINGS | Context-mode AGENTS 第 62 行称 Skills、roles、decisions 持续整个 session；当前注入会话规则规定记忆仅为辅助，最新用户请求优先。磁盘规则是否生效 UNKNOWN，故为条件性语义冲突。 |
| D3 过度确认 | PASS | 检查的当前规则没有低风险本地修改逐步确认要求；高风险操作仍保留明确授权与安全边界。 |
| D4 无条件 Context 加载 | FINDINGS | 若 977 行用户 AGENTS 被所有任务全局加载，则其完整 11 类盘点、D1–D9、T1–T8 会变成普通任务的固定开销。该文件自动加载范围 UNKNOWN；不要把候选当成已生效事实。 |
| D5 Skill 过度触发 | FINDINGS | 5 个入口范围偏宽或重叠：source-driven-development、test-driven-development、browser-testing-with-devtools、frontend-coding、frontend-ui-engineering。其余入口如 playwright、browser-qa、e2e-testing、i-have-adhd 边界相对明确。 |
| D6 过度测试 | FINDINGS | TDD 入口以“任意逻辑、bug 或行为变化”触发，与当前开发者层的“除非用户要求测试 / 验证，不要新增或运行测试”及 Ponytail 的单一最小检查形成测试压力叠加。README 的 typecheck/build 只在发布前适用，属于项目门槛，保留。 |
| D7 过度 Delegation | FINDINGS | 未发现强制简单任务开 Agent 的通用规则。T7 为满足审计要求，主代理和一个子代理都读了 package.json，出现一次低成本重复读取；没有重复实现或重叠写入。 |
| D8 并行写冲突 | FINDINGS | Agent 与主代理共享工作目录是工具运行时事实；若不同 Agent 同时写同一配置，存在最后写入者覆盖风险。T7 agents 均只读且分配 README.md 与 package.json / workflow 目录所有权，未发生冲突。标记 PARALLEL_WRITE_CONFLICT_RISK（潜在）。 |
| D9 过早停止 / 过度完成 | PASS | 检查范围内没有“遇首错即停止”或“已验收仍继续扩展”的规则。当前 PPTX 诊断继续完成了类型检查和构建；尚无代码修改，不把它当成 bug 已修复。 |

### 重要审计记录

RULE_ID = R01
SOURCE_FILE = C:\Users\admin\.codex\AGENTS.md（第 19–29、335–350、584–600、800–833 行）
SOURCE_SCOPE = 用户级候选；自动加载范围 UNKNOWN；本轮由用户直接粘贴，故本审计有效
CATEGORY = C4
ISSUE_TYPE = D4
CURRENT_RULE = 该流程要求先全面盘点规则，再跑 9 项扫描和最多 8 项 A/B；文末限定为本次审计和禁止本任务内 Adoption。
WHY_IT_MAY_BE_A_PROBLEM = 若被当作常驻全局流程，普通代码任务也可能加载长流程并扩展出不相关检查。
REAL_RISK = 上下文和执行成本增大；同时可能影响任务闭环速度。加载行为未证实。
PROPOSED_ACTION = MOVE
PROPOSED_RULE = 仅在用户明确要求“规则审计 / Rules Audit”时加载完整审计流程；其他任务只继承可适用的安全授权底线。
SAFETY_IMPACT = 保留 C1/C2 高风险边界；不建议删安全条款。
CONTEXT_IMPACT = 普通任务不再默认携带完整 D1–D9 / T1–T8 过程；未测量 Token 节省。
BEHAVIOR_IMPACT = 审计任务保持完整；普通任务不被审计流程接管。
CONFIDENCE = MEDIUM

RULE_ID = R02
SOURCE_FILE = C:\Users\admin\.codex\plugins\cache\context-mode\context-mode\1.0.169\configs\codex\AGENTS.md:62；当前系统注入的 Session Continuity 规则
SOURCE_SCOPE = 插件 Codex 配置候选与当前 turn 注入规则；磁盘文件生效性 UNKNOWN
CATEGORY = C2
ISSUE_TYPE = D2
CURRENT_RULE = 磁盘文本要求 Skills、roles、decisions 整个 session 持续；当前注入规则要求把这些当记忆辅助并服从最新用户输入。
WHY_IT_MAY_BE_A_PROBLEM = 旧角色或决策可能被误当成永久指令，与新请求发生冲突。
REAL_RISK = 可能误用过期偏好或角色；本轮实际以最新的 pptx.ts 问题为当前诊断 steering。
PROPOSED_ACTION = REWRITE
PROPOSED_RULE = “把先前 Skills、roles 和 decisions 当作上下文提示；如与最新用户请求冲突，遵循最新请求。”
SAFETY_IMPACT = 不降低安全要求；用户新请求仍受安全边界约束。
CONTEXT_IMPACT = 不变。
BEHAVIOR_IMPACT = 允许新请求更新旧目标和偏好。
CONFIDENCE = MEDIUM

RULE_ID = R03
SOURCE_FILE = C:\Users\admin\.codex\skills\source-driven-development\SKILL.md:3
SOURCE_SCOPE = 用户 Skill 入口；当前可见，具体触发记录不完整
CATEGORY = C5
ISSUE_TYPE = D5
CURRENT_RULE = 对所有 framework/library implementation decision 都要求 official documentation。
WHY_IT_MAY_BE_A_PROBLEM = 对稳定、低风险、代码库已有惯例的改动也可能过度触发外部检索。
REAL_RISK = 多做无必要查询；需官方规范的 OOXML、Nuxt API 仍确实应查官方资料。
PROPOSED_ACTION = REWRITE
PROPOSED_RULE = “当任务依赖外部规范、版本敏感 API、陌生库行为或需要权威引用时使用；已有代码惯例和稳定简单实现不强制外查。”
SAFETY_IMPACT = 保留准确性；不回避适用的官方 OOXML / Nuxt 文档。
CONTEXT_IMPACT = 减少无关文档加载。
BEHAVIOR_IMPACT = 对不熟悉或时效敏感决策仍查权威来源。
CONFIDENCE = HIGH

RULE_ID = R04
SOURCE_FILE = C:\Users\admin\.codex\skills\test-driven-development\SKILL.md:3
SOURCE_SCOPE = 用户 Skill 入口；当前可见，具体触发记录不完整
CATEGORY = C5
ISSUE_TYPE = D5
CURRENT_RULE = 任意逻辑、bug、行为变化或 bug report 都触发 TDD。
WHY_IT_MAY_BE_A_PROBLEM = 把 TDD 扩展成所有微小变更或只需诊断的任务的固定流程。
REAL_RISK = 机械增加测试；目前 typecheck/build 通过，不代表需要再加测试。
PROPOSED_ACTION = REWRITE
PROPOSED_RULE = “用户要求测试、项目约定要求测试，或新增/改变的非平凡行为可由现有测试工具验证时使用；纯诊断、微小单行修正或没有可复用测试环境时，按风险选择最小验证。”
SAFETY_IMPACT = 不跳过高风险逻辑的必要验证。
CONTEXT_IMPACT = 避免为简单工作载入完整 TDD playbook。
BEHAVIOR_IMPACT = 测试跟随风险和项目约定。
CONFIDENCE = HIGH

RULE_ID = R05
SOURCE_FILE = C:\Users\admin\.codex\skills\browser-testing-with-devtools\SKILL.md:3；browser-qa、playwright、e2e-testing 的 description 第 3 行
SOURCE_SCOPE = 多个用户 Skill 入口；当前目录可见
CATEGORY = C5
ISSUE_TYPE = D5
CURRENT_RULE = DevTools Skill 对“任何在浏览器运行的工作”触发；其他三个 Skill 分别针对视觉交互、终端浏览器自动化、Playwright E2E 编写。
WHY_IT_MAY_BE_A_PROBLEM = 普通 Vue 代码也可能同时触发 DevTools、Playwright 和 E2E 规则。
REAL_RISK = 多加载一套 QA 指引，或重复要求浏览器验证。
PROPOSED_ACTION = REWRITE
PROPOSED_RULE = “DevTools：需要 DOM、console、网络或性能数据时；Playwright：需要终端驱动真实浏览器流程时；browser-qa：需要视觉 / 交互验收时；e2e-testing：创建或维护 E2E 测试时。”
SAFETY_IMPACT = 保留浏览器运行时证据要求。
CONTEXT_IMPACT = 各入口只加载匹配的验证指南。
BEHAVIOR_IMPACT = 按任务目标选择单一首选 Skill，必要时再组合。
CONFIDENCE = HIGH

RULE_ID = R06
SOURCE_FILE = C:\Users\admin\.codex\skills\frontend-coding\SKILL.md:3；frontend-ui-engineering、frontend-skill 的 description 第 3 行
SOURCE_SCOPE = 用户 Skill 入口；当前目录可见
CATEGORY = C5
ISSUE_TYPE = D5
CURRENT_RULE = frontend-coding 覆盖 Vue / Nuxt 的一般页面和前端工程；frontend-ui-engineering 也覆盖页面、组件、布局、状态和无障碍；frontend-skill 再覆盖视觉界面。
WHY_IT_MAY_BE_A_PROBLEM = 三个 Skill 对页面实现和组件改动的触发语义重合。
REAL_RISK = 重复载入设计与实现建议，增加约束而没有明确优先级。
PROPOSED_ACTION = REWRITE
PROPOSED_RULE = “frontend-coding 用于 Vue / Nuxt 行为、数据流、SSR、路由和状态；frontend-ui-engineering 用于可访问性 / 响应式与生产 UI 完成度专项；frontend-skill 仅在用户强调视觉风格、品牌或美术表现时触发。”
SAFETY_IMPACT = 保留 a11y 需求。
CONTEXT_IMPACT = 视觉设计规则只在相关任务加载。
BEHAVIOR_IMPACT = 入口边界可预测。
CONFIDENCE = MEDIUM

RULE_ID = R07
SOURCE_FILE = C:\Users\admin\.codex\plugins\cache\context-mode\context-mode\1.0.169\configs\codex\AGENTS.md:25–27、57
SOURCE_SCOPE = 插件配置规则；生效性 UNKNOWN；本轮因外部路径访问被拒后采用普通 PowerShell 做了已授权只读检查
CATEGORY = C6
ISSUE_TYPE = D4
CURRENT_RULE = 普通 Shell 只允许少量命令，其余需走 context-mode；文本规定产出文件写入文件。
WHY_IT_MAY_BE_A_PROBLEM = 把特定插件的批处理偏好扩成所有短小命令 / 原生文件操作的固定路由。
REAL_RISK = 工具选择受限，或为了一个短输出创建多余处理流程。
PROPOSED_ACTION = REWRITE
PROPOSED_RULE = “大文件、长日志、统计和批量筛选优先用 context-mode；短小直接观察、项目 Git 操作和需要持久化的文件修改使用合适的原生工具。”
SAFETY_IMPACT = 不改变权限、外部写入或不可逆操作的安全边界。
CONTEXT_IMPACT = 小任务少走不必要的工具链。
BEHAVIOR_IMPACT = 仍为大输出保留低噪声处理。
CONFIDENCE = MEDIUM

RULE_ID = R08
SOURCE_FILE = 当前 Multi-Agent 注入规则；Codex agent collaboration runtime（无可读本地文件）
SOURCE_SCOPE = 当前工具运行时；Agent 共享工作区已由运行时说明确认
CATEGORY = C8
ISSUE_TYPE = D8
CURRENT_RULE = Worker 应有清晰文件所有权；所有 Agent 共享同一目录。
WHY_IT_MAY_BE_A_PROBLEM = 两个写 Agent 若改同一文件或配置，可能互相覆盖。
REAL_RISK = PARALLEL_WRITE_CONFLICT_RISK；本次 T7 只读，无真实冲突。
PROPOSED_ACTION = KEEP
PROPOSED_RULE = 并行前声明文件 / 目录所有权；重叠写串行化；审计和探索任务默认只读。
SAFETY_IMPACT = 保护用户修改和配置完整性。
CONTEXT_IMPACT = 无。
BEHAVIOR_IMPACT = 只并行真正独立任务。
CONFIDENCE = HIGH

RULE_ID = R09
SOURCE_FILE = 当前开发者层“除非用户要求测试或验证，不要新增或运行测试”规则；Ponytail AGENTS.md:30；README.md:12–14；TDD Skill:3
SOURCE_SCOPE = 本轮当前注入 / 用户 Skill / 项目发布要求，适用范围各异
CATEGORY = C7
ISSUE_TYPE = D6
CURRENT_RULE = 通用测试限制、非平凡逻辑留一个最小检查、bug 一律触发 TDD、发布前 typecheck/build 同时存在。
WHY_IT_MAY_BE_A_PROBLEM = 若忽略范围条件，容易把发布门槛套到每个微小修改，也可能出现“测试不能跑”与“每个 bug 必须 TDD”叠加。
REAL_RISK = 过量测试或不必要的验证流程。README 发布门槛本身并不过量，应保持作用域。
PROPOSED_ACTION = REWRITE
PROPOSED_RULE = “对非平凡逻辑和风险敏感路径，执行一个最小有效检查；用户要求验证或项目发布流程要求时再跑对应检查。发布前保留 typecheck/build。”
SAFETY_IMPACT = 保留安全、数据处理、解析逻辑必要验证。
CONTEXT_IMPACT = 小型纯诊断任务不加载全套测试步骤。
BEHAVIOR_IMPACT = 风险驱动验证，不要求机械扩大测试。
CONFIDENCE = MEDIUM

RULE_ID = R10
SOURCE_FILE = C:\Users\admin\.codex\AGENTS.md:23–29、586–600、747–768、800–833、970–976
SOURCE_SCOPE = 本轮用户提供且直接生效的审计规范；该文件对普通任务的加载范围 UNKNOWN
CATEGORY = C10
ISSUE_TYPE = D1
CURRENT_RULE = Adoption 必须分离、不得伪造 Runtime Evidence、A/B 后要验证回滚，在正文与末尾重复强调。
WHY_IT_MAY_BE_A_PROBLEM = 同一高优先级闸门多处复述会增加文档长度；但它们分别在任务范围、验证和底线阶段出现，具备阶段提醒价值。
REAL_RISK = 如果把它们机械合并，可能削弱安全闸门；如果全局加载，普通任务会携带冗长报告流程。
PROPOSED_ACTION = KEEP
PROPOSED_RULE = 保留 Adoption / 证据 / 回滚阶段的关键提醒；将该审计流程整体限定为显式审计任务触发。
SAFETY_IMPACT = 保持原有高风险隔离。
CONTEXT_IMPACT = 通过条件触发减少无关加载。
BEHAVIOR_IMPACT = 关键阶段继续提醒。
CONFIDENCE = HIGH

## 5. Top Findings

1. R01：如果 977 行审计规程在全局常驻，普通任务也可能被要求先做完整规则盘点和 A/B。其自动加载范围 UNKNOWN，建议移动为显式审计流程。
2. R02：Context-mode 的 session 持续指令与最新用户请求优先规则存在条件性冲突。
3. R03–R06：5 个 Skill 入口宽泛或重叠；将 OOXML 官方资料检索保留为按需触发。
4. R07：Context-mode 的 Shell 路由规定可能把插件偏好扩展成全局工具限制。
5. R08：共享工作区有并行写风险；本次通过只读和所有权隔离避免了实际覆盖。
6. R09：测试规则在全局、Skill、Ponytail 和项目发布流程之间需要明确适用范围。
7. 安全保留项：密钥、隐私、删除、生产、外发、权限与不可逆操作的保护要求没有提出弱化；README 的 PPTX 格式支持和本地处理要求应保留。

## 6. Candidate Changes

只提出 Candidate，没有编辑任何规则文件。

| 动作 | 建议 |
|---|---|
| KEEP | 保留 C1 安全 / 权限 / 不可逆操作边界；保留 README 的本地解析、兼容矩阵及发布前 typecheck/build 要求；保留工作区所有权和回滚保护。 |
| REWRITE | 收窄 source-driven、TDD、DevTools、frontend-coding 和 frontend-ui-engineering 触发描述；修正 Context-mode 的 session continuity 表述；按风险明确测试强度。 |
| MOVE | 将完整 D1–D9 / T1–T8 审计流程移到仅由明确“规则审计”请求触发的 audit Skill / playbook。 |
| MERGE | 对重复的非安全性说明可并入一个阶段检查；Adoption、无伪证据、回滚等高风险闸门保留多阶段提醒。 |
| REMOVE | 无建议删除项。 |
| NEEDS_USER_DECISION | 是否将 C:\Users\admin\.codex\AGENTS.md 作为普通 Codex 任务的全局常驻规则；是否希望未来单独开启 Adoption Task。 |

## 7. Multi-Agent Findings

- 过度 Delegation：没有发现强制拆分普通小任务的有效规则。T7 启动两个 Agent 是用户审计规范明确要求。
- 并行写冲突：潜在风险存在，因为共享工作目录；本次 workers 都是只读且所有权分离，没有写冲突。
- 等待 / 轮询：仅等待已派发的独立检查结果；没有反复空轮询。
- 主代理 / 子代理重复：T7 worker 与主代理都读取了 package.json，以分别提供并行和单代理证据；这是一次低成本重复读取，没有重复实现或重复测试。未来应选不同的只读文件作为串行对照。
- 实际 Agent 运行证据：工具状态显示 /root/t7_readme_scope 和 /root/t7_test_scope 均完成；前者只读 README.md，后者只读 package.json 与 workflow 目录存在性。该证据只证明本次工具调用，不能证明 Candidate 改善了运行时行为。

## 8. Runtime A/B

Codex 当前工具没有暴露可在隔离进程中替换全局 / 用户 AGENTS、Skills 与注入规则的安全开关。config.toml 只显示插件和 Hook 状态，没有规则 A/B 路由。因此 T1–T8 没有进行原配置与 Candidate 配置的比较，也没有为 A/B 写入配置文件。

| 测试 | 状态 | Runtime Evidence / 限制 |
|---|---|---|
| T1 简单文件修改 | NO_RUNTIME_EVIDENCE | 无 A/B 规则切换，不执行配置对比。 |
| T2 需要特定 Skill | NO_RUNTIME_EVIDENCE | 可见入口描述，但没有基线 / Candidate 运行时切换日志。 |
| T3 不需要 Skill | NO_RUNTIME_EVIDENCE | 没有比较误触发率的工具事件。 |
| T4 测试行为 | NO_RUNTIME_EVIDENCE | 当前 typecheck/build 是代码诊断，不是规则 A/B。 |
| T5 确认边界 | NO_RUNTIME_EVIDENCE | 本次没有执行高风险操作或隔离模拟，也不据此声称 Candidate 保留了防护。 |
| T6 任务闭环 | NO_RUNTIME_EVIDENCE | 本次未进行修改→失败→修复→再测试的规则对比。 |
| T7 Multi-Agent | NO_RUNTIME_EVIDENCE（A/B） | 有补充运行时事实：2 个只读 subagent 实际并行，主代理也实际独立检查；工具列表与完成状态可验证。由于没有真实 Candidate 配置切换，只能证明委派工具运行过，不能证明 A/B 效果。 |
| T8 Context / 文档读取 | NO_RUNTIME_EVIDENCE | 观察到 README.md、package.json 和相关 Skill 入口被读取；没有原规则和 Candidate 的同任务对照或可信 Token 计量。 |

### T7 runtime record

TEST_ID = T7
A_CONFIG = 当前运行时；未建立对照配置
B_CONFIG = Agent prompt 中使用了只读和文件所有权候选约束；没有持久化 / 切换 Candidate 配置
TASK = 并行只读检查 README 和 package scripts；主代理单独读取 package scripts 作为串行例子
FILES_READ = README.md；package.json；检查 .github/workflows 目录不存在
SKILLS_TRIGGERED = UNVERIFIED
TOOLS_USED = collaboration.spawn_agent、collaboration.list_agents、agent completion records、exec_command
AGENTS_USED = /root/t7_readme_scope、/root/t7_test_scope
SUBAGENTS_USED = 2
TESTS_RUN = none
CONFIRMATIONS_REQUESTED = 0
RESULT = 两个只读 Agent 并行完成、文件所有权清楚；主代理串行检查完成。没有并行写或覆盖。
RUNTIME_EVIDENCE = 子 Agent 完成状态和返回内容；主代理命令结果；不是模型自述。
LIMITATIONS = package.json 被 Agent 和主代理重复读取；没有规则基线 / Candidate 配置 A/B，不得声称效率或 Context 改善。

### 当前 pptx.ts 诊断（独立于规则 A/B）

- npm run typecheck：exit code 0。
- npm run build：exit code 0。
- app/utils/pptx.ts 没有冲突分隔符；git ls-files -u 为空。
- git diff --check -- app/utils/pptx.ts：exit code 0。
- 全局 git diff --check 报出 scripts/verify-animation.mjs 的 CRLF 行尾差异，不是 pptx.ts 编译诊断。
- 结论：本次工具链未复现用户 IDE 中的错误提示；未修改 pptx.ts。请将具体诊断文本或截图提供给后续修复任务，避免盲改这份 1,717 行差异。

## 9. Rollback Verification

没有修改任何 AGENTS、SKILL、config 或项目规则文件；没有创建 A/B 配置副本或隔离 overlay。审计报告是唯一新增的交付文件。

BASELINE_FILE = NONE（未修改规则文件）
BASELINE_SHA256 = N/A
BACKUP_LOCATION = NOT_CREATED
BASELINE_GIT_STATUS = main / HEAD 5bd4c16aa8f8f916999cc6a2c50f77a36c0ef560 / 12 个 modified + 16 个 untracked

以下 19 个代表性来源在报告创建前记录 baseline SHA-256，并在报告保存后重新计算。最终值逐项与 baseline 相等。由于没有发生规则写入或 A/B 配置切换，回滚结果是无修改条件下的精确一致性确认。

ROLLBACK_VERIFICATION = PASS

| SHA-256（baseline = final，已复核） | File |
|---|---|
| 5A7AB34C3EA236988652213121A91299D03B413357423ABCD12EDC2540927EC8 | C:\Users\admin\.codex\AGENTS.md |
| 74DB66897DD59AC3783A4EA36B8A32DFB79FA713105A1FE69638A4C50B61DA7B | C:\Users\admin\.codex\config.toml |
| B57F736A83DDAB7111D752008C3E131A157E646A6B47C9EB223F076C494685A8 | C:\Users\admin\.codex\plugins\cache\ponytail\ponytail\4.10.0\AGENTS.md |
| 3F06CFBCE24FDFDE17B4FF03BB25796471C8A9EADD671C125CD8EC6BA1BACB91 | C:\Users\admin\.codex\plugins\cache\context-mode\context-mode\1.0.169\configs\codex\AGENTS.md |
| 7A2BEB941E77CF4CF749DEE31F74AFD10D6E87B4D8C6AA50B239F968D41ED659 | C:\Users\admin\.codex\plugins\cache\i-have-adhd\i-have-adhd\0.3.0\AGENTS.md |
| 719D4E54083C90DED62112FB41DF3DBC4619309118EE0E6AA4D846F92D8204AF | C:\Users\admin\.codex\skills\source-driven-development\SKILL.md |
| 67CE2C9442DA0C5A6E3515617FC9C4003CFE232EF7C7210DA342F40F508F9958 | C:\Users\admin\.codex\skills\debugging-and-error-recovery\SKILL.md |
| 440FB3C9B76F4F1B0B564DFBAEC19D370BEAB0167A3E2AB048CE59C5A865ECEF | C:\Users\admin\.codex\skills\test-driven-development\SKILL.md |
| F0D56107F79E607D4DE6C5543B8EDFB1B58AB9AE4019B39831689CD91116543E | C:\Users\admin\.codex\skills\verification-loop\SKILL.md |
| 3DE86600C8E0F559750478BC8DE0F6B59C96B6E25010F8E2F2FE7A2709CAF282 | C:\Users\admin\.codex\skills\frontend-coding\SKILL.md |
| 2B74AC4862BE3902EC918DCEAC9366A6FE83B9E003601C0DEAF6BE09C1766ACA | C:\Users\admin\.codex\skills\frontend-ui-engineering\SKILL.md |
| D367B2D22825D74C169475507E7FF109B81C1E9C20B2DE3421FB19FCFBE3329B | C:\Users\admin\.codex\skills\frontend-skill\SKILL.md |
| C501257D2C9CA711342B56D7ADFB04897F17C6946DC6CDB88A5C75C87E4AA951 | C:\Users\admin\.codex\skills\browser-qa\SKILL.md |
| 4E3AACD6A380CD25BC6C2D67FDD1C926A9B22535B8A62109ECD33CEFD909E3D9 | C:\Users\admin\.codex\skills\browser-testing-with-devtools\SKILL.md |
| 0FFAABCC8E0990627C4725F18BF1C7955534A796C1C199E872909DE2013CE6A8 | C:\Users\admin\.codex\skills\playwright\SKILL.md |
| B0E00E7E17CB259101139900816C5528AED18DD10BCF5F9CB42CFC35BAF8A755 | C:\Users\admin\.codex\skills\e2e-testing\SKILL.md |
| 3170B16ACE00AECB0DD7FEB54C0B5AA642E7502ACDA06ECD24FD89A11C7127E9 | C:\Users\admin\.codex\plugins\cache\i-have-adhd\i-have-adhd\0.3.0\skills\i-have-adhd\SKILL.md |
| CE26CB01DE7A116EAB256F916620935F029E850771F3FB6CB7D40A3CE8D1F7FF | C:\Users\admin\.codex\plugins\cache\context-mode\context-mode\1.0.169\skills\context-mode\SKILL.md |
| 1316A2F3F95741D2300B116FE0C2D81CE4A9568656ED0A62643F54AAF09957F2 | C:\Users\admin\.codex\plugins\cache\ponytail\ponytail\4.10.0\skills\ponytail\SKILL.md |

## 10. Adoption Recommendation

ADOPTION_RECOMMENDATION = NEEDS_MORE_EVIDENCE
WHY = 关键来源的自动加载范围 UNKNOWN；T1–T8 缺少可切换且隔离的真实规则运行时，无法声称 Candidate 改善。
EXPECTED_BENEFIT = 若未来验证通过，可减少无关审计流程、减少宽泛 Skill 误触发、让测试和委派随任务风险选择。
KNOWN_RISKS = 过度收窄可能遗漏安全、官方规范或项目知识；全部 C1 安全边界及 README 项目事实必须保留。
UNKNOWN = 用户级 AGENTS 实际加载范围；Skill catalog 与插件缓存的选版映射；候选规则在真实 Codex Runtime 的触发行为。
NEXT_TASK = 先在独立 Task 中确认 AGENTS / Skill 的实际加载机制；若证据充分，再开一个单独 Task 执行受控 Adoption 和回滚。
正式 Adoption 必须在新的独立 Task 中执行。
