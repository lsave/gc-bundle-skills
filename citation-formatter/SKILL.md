---
name: citation-formatter
description: "参考文献格式规范。写论文引用格式老出错？自动管理参考文献，一键切换APA/MLA格式 This skill should be used when the user asks about 参考文献格式规范. Keywords: 参考文献, 引用格式, APA格式."
---

# 参考文献格式规范

> 写论文引用格式老出错？自动管理参考文献，一键切换APA/MLA格式

## 前置依赖

```bash
pip install pandas openpyxl
```

## 核心能力

### 能力1：通过 DOI/URL 导入文献（web_fetch）

用 `read_file` 读取原始引用数据，用 `execute_command` 运行格式转换脚本，用 `write_to_file` 输出标准化引用。

### 能力2：自动提取文献元数据（标题/作者/年份/期刊）

用 `read_file` 读取原始引用数据，用 `execute_command` 运行格式转换脚本，用 `write_to_file` 输出标准化引用。

### 能力3：生成指定格式引用（APA/MLA/GB-T

用 `write_to_file` 生成文件。

### 能力4：自动去重排序

用 `read_file` 读取原始引用数据，用 `execute_command` 运行格式转换脚本，用 `write_to_file` 输出标准化引用。

### 能力5：输出完整参考文献列表

用 `write_to_file` 生成文件。

## 使用流程

### 步骤 1：收集用户需求

向用户确认以下信息（如果未主动提供）：
- 需要格式化的参考文献（提供原始文本或BibTeX文件路径）
- 目标引用格式（APA/MLA/Chicago/GB/T 7714/IEEE）
- 输出方式（文本/BibTeX/Word兼容）

### 步骤 2：检索外部信息

执行以下搜索获取真实数据：

```
web_search("[用户主题] paper arXiv")
web_search("[用户主题] 研究综述")
```

对搜索结果中的重要链接，用 `web_fetch` 抓取页面详细内容，提取关键信息。

确保获取到以下资源：
- 文献元数据
- 格式化引用模板

### 步骤 3：运行脚本处理数据

```bash
python3 scripts/citation_formatter_tool.py run \
  --input "用户提供的输入" \
  --output "/path/to/output_file"
```

读取脚本输出的结果，确认数据处理成功。

### 步骤 4：生成最终产出

基于脚本输出和搜索到的资源，用 `write_to_file` 生成以下文件：

- **参考文献列表文档**
- **BibTeX 文件**

输出格式要求：格式化参考文献列表 + BibTeX 文件

### 步骤 5：汇总交付

向用户展示：
1. 生成的文件路径和内容摘要
2. 搜集到的资源链接列表
3. 关键发现和建议

## 输出格式

```markdown
# 📋 参考文献格式规范 — 执行报告

**生成时间**: YYYY-MM-DD HH:MM
**目标用户**: 论文写作者、学术研究人员

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

- ✅ 文献导入成功
- ✅ 格式符合规范
- ✅ 去重排序正确
- ✅ 支持一键切换格式

## 场景化适配

根据目标格式（APA/MLA/GB-T7714）调整输出


## 依赖 Skills

本 Skill 参考以下已有 Skill 的能力进行增强：
- **citation-manager**

## 注意事项

- 所有数据必须来自 `web_search` / `web_fetch` 的真实搜索结果，**严禁编造数据**
- 数据缺失时标注"数据不可用"而非猜测
- 报告必须保存为文件（`write_to_file`），不能只在对话中输出
- 建议结合人工判断使用，AI 分析仅供参考
