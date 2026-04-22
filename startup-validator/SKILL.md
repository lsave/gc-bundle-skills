---
name: startup-validator
description: "创业点子可行性验证。帮你联网搜索竞品、分析市场规模、评估技术可行性，生成完整可行性报告和MVP路线图。This skill should be used when the user asks about validating a startup idea, analyzing market feasibility, competitor analysis, MVP planning, or business model evaluation. Keywords: 验证想法, 创业分析, 可行性, 竞品分析, MVP, startup validation, market analysis, idea validation."
---

# 创业点子可行性验证

> 有想法不敢下手？帮你快速验证可行性，零成本少踩坑。

## 前置依赖

```bash
pip install pandas requests
```

## 核心能力

### 能力一：竞品搜索与市场数据采集

对用户提出的创业想法，自动进行全网竞品扫描：

| 维度 | 具体动作 |
|------|---------|
| 竞品搜索 | 用 `web_search` 搜索 "[想法关键词] 竞品"、"[关键词] alternative" 等，找到至少5个竞品 |
| 竞品详情 | 用 `web_fetch` 抓取每个竞品官网，提取产品定位、核心功能、定价策略 |
| 市场规模 | 用 `web_search` 搜索 "[行业] 市场规模 2024"、"[industry] market size TAM" |
| 融资信息 | 用 `web_search` 搜索竞品的融资轮次和金额 |

### 能力二：市场规模与竞争格局分析

基于采集到的数据，运行分析脚本：

```bash
python3 scripts/startup_validator_tool.py analyze --idea "用户的创业想法" --competitors "竞品1,竞品2,竞品3"
```

脚本会生成结构化的竞品对比矩阵（JSON），AI 再据此撰写分析。

### 能力三：技术可行性评估与 MVP 路线图

基于分析结果，生成 MVP 功能清单和实施路线图，输出为 Markdown 文件。

## 使用流程

当用户说"帮我验证一个创业想法"或"分析这个项目的可行性"时，按以下步骤执行：

### 步骤 1：收集创业想法信息

向用户询问以下信息（如果用户未主动提供）：
- 创业想法的一句话描述
- 目标用户群体
- 解决的核心痛点
- 预期的商业模式（可选）

### 步骤 2：竞品搜索（至少找到5个竞品）

依次执行以下搜索：

```
web_search("用户想法关键词 竞品分析")
web_search("用户想法关键词 alternative competitor")
web_search("用户想法关键词 市场规模 行业报告")
```

对每个找到的竞品，用 `web_fetch` 抓取其官网首页，提取：
- 产品名称和定位
- 核心功能列表
- 定价模式
- 目标用户

### 步骤 3：运行分析脚本生成竞品矩阵

```bash
python3 scripts/startup_validator_tool.py analyze \
  --idea "用户的创业想法描述" \
  --competitors "竞品1:定位1,竞品2:定位2,竞品3:定位3,竞品4:定位4,竞品5:定位5" \
  --market-keywords "市场关键词1,市场关键词2"
```

读取脚本输出的 JSON 结果，包含竞品对比矩阵和差异化分析框架。

### 步骤 4：搜索市场规模数据

```
web_search("行业名称 market size 2024 2025")
web_search("行业名称 TAM SAM SOM")
web_search("行业名称 增长率 CAGR")
```

将搜索到的数据补充到分析中。

### 步骤 5：生成可行性报告和 MVP 路线图

运行脚本生成报告框架：

```bash
python3 scripts/startup_validator_tool.py report \
  --idea "创业想法" \
  --output "/path/to/feasibility_report.md"
```

然后用 `write_to_file` 将完整报告写入文件，包含以下章节：

1. **执行摘要** — 一段话总结可行性结论
2. **竞品分析** — 竞品对比表格（功能/定价/优劣势）
3. **市场分析** — TAM/SAM/SOM + 增长趋势
4. **技术可行性** — 技术栈建议、开发周期估算
5. **MVP功能清单** — 核心功能 vs 可延后功能
6. **实施路线图** — 按月/季度的里程碑
7. **风险评估** — 主要风险及应对策略
8. **结论与建议** — Go/No-Go 建议

## 输出格式

最终输出两个文件：

### 文件1：可行性分析报告（Markdown）

```markdown
# 📊 创业可行性分析报告：[想法名称]

**生成时间**: YYYY-MM-DD HH:MM
**分析师**: AI Startup Validator

## 一、执行摘要
[基于真实搜索数据的一段话结论]

## 二、竞品分析（≥5个竞品）
| 竞品 | 定位 | 核心功能 | 定价 | 融资情况 | 优势 | 劣势 |
|------|------|---------|------|---------|------|------|
| [真实竞品] | [真实定位] | [真实功能] | [真实定价] | [真实融资] | [...] | [...] |

## 三、市场分析
- **TAM**: [数据来源+数值]
- **SAM**: [数据来源+数值]
- **SOM**: [估算+依据]
- **CAGR**: [增长率+来源]

## 四、技术可行性评估
[具体技术栈、开发难度、开发周期]

## 五、MVP 功能清单
| 优先级 | 功能 | 实现复杂度 | 开发周期 |
|--------|------|----------|---------|
| P0 | [核心功能] | [高/中/低] | [X周] |

## 六、实施路线图
[按季度的里程碑计划]

## 七、风险评估
| 风险 | 概率 | 影响 | 应对策略 |
|------|------|------|---------|

## 八、结论
🟢/🟡/🔴 **[建议Go/谨慎/不建议]**
[具体理由]

---
⚠️ 本报告基于公开数据分析，仅供参考，不构成投资建议。
```

### 文件2：MVP 路线图（Markdown）

单独生成一个 `mvp_roadmap.md` 文件，包含详细的功能拆解和时间规划。

## 验收标准

- ✅ 竞品搜索≥5个
- ✅ 市场规模有依据
- ✅ MVP清单可执行
- ✅ 报告结构完整

## 参考资料

- [YC 产品市场契合度指南](https://www.ycombinator.com/library/5z-the-real-product-market-fit)
- [Lean Startup 方法论](https://theleanstartup.com/principles)

## 注意事项

- 所有竞品数据必须来自 `web_search` + `web_fetch` 的真实搜索结果，**严禁编造**
- 市场规模数据必须标注来源，找不到时标注"数据不可用"
- 竞品数量不少于 5 个
- 报告必须保存为文件（`write_to_file`），不能只在对话中输出
