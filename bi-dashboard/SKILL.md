---
name: bi-dashboard
description: "业务指标可视化管理。业务数据看板生成器 This skill should be used when the user asks about 业务指标可视化管理. Keywords: 数据看板, KPI监控, 可视化."
---

# 业务指标可视化管理

> 业务数据看板生成器

## 前置依赖

```bash
pip install pandas openpyxl
```

## 核心能力

### 能力1：读取业务数据文件（read_file，CSV/Excel）

用 `web_fetch` 抓取页面内容 / 用 `read_file` 读取文件。

### 能力2：提取关键 KPI 指标

用 `read_file` 读取数据文件，用 `execute_command` 运行Python（matplotlib/plotly）生成图表，用 `write_to_file` 生成HTML仪表盘。

### 能力3：用 Python（execute_command）生成可视化图表

用 `write_to_file` 生成文件。

### 能力4：设置指标阈值告警规则

用 `read_file` 读取数据文件，用 `execute_command` 运行Python（matplotlib/plotly）生成图表，用 `write_to_file` 生成HTML仪表盘。

### 能力5：生成数据看板报告（HTML 或 Markdown）

用 `write_to_file` 生成文件。

## 使用流程

### 步骤 1：收集用户需求

向用户确认以下信息（如果未主动提供）：
- 数据源文件路径（CSV/Excel/数据库连接）
- 需要哪些指标的可视化？
- 仪表盘受众（管理层/运营/技术）
- 更新频率（一次性/每日/每周）

### 步骤 2：运行脚本处理数据

```bash
python3 scripts/bi_dashboard_tool.py run \
  --input "用户提供的输入" \
  --output "/path/to/output_file"
```

读取脚本输出的结果，确认数据处理成功。

### 步骤 3：生成最终产出

基于脚本输出和搜索到的资源，用 `write_to_file` 生成以下文件：

- **数据看板报告（HTML/Markdown）**
- **图表文件**

输出格式要求：HTML/Markdown 看板 + 图表 + 告警规则

### 步骤 4：汇总交付

向用户展示：
1. 生成的文件路径和内容摘要
2. 搜集到的资源链接列表
3. 关键发现和建议

## 输出格式

```markdown
# 📋 业务指标可视化管理 — 执行报告

**生成时间**: YYYY-MM-DD HH:MM
**目标用户**: 运营总监、产品负责人、数据分析师

## 执行摘要
[基于实际执行结果的一段话摘要]

## 详细结果

### 📊 生成的文件
| 文件名 | 类型 | 说明 |
|--------|------|------|
| [文件名] | [类型] | [说明] |

### 🔗 资源链接
| 名称 | 链接 | 说明 |
|------|------|------|
| [资源] | [URL] | [说明] |

## 行动建议
[具体的下一步建议]
```

## 验收标准

- ✅ KPI 提取准确
- ✅ 图表清晰可读
- ✅ 告警规则合理
- ✅ 看板可交付

## 场景化适配

根据业务类型（电商/SaaS/内容）调整KPI定义


## 依赖 Skills

本 Skill 参考以下已有 Skill 的能力进行增强：
- **analytics-dashboard**

## 注意事项

- 所有数据必须来自 `web_search` / `web_fetch` 的真实搜索结果，**严禁编造数据**
- 数据缺失时标注"数据不可用"而非猜测
- 报告必须保存为文件（`write_to_file`），不能只在对话中输出
- 建议结合人工判断使用，AI 分析仅供参考
