---
name: knowledge-framework
description: "知识点框架梳理。看书越看越乱？帮你把知识点理成一张清晰骨架图 This skill should be used when the user asks about 知识点框架梳理. Keywords: 知识框架, 知识点梳理."
---

# 知识点框架梳理

> 看书越看越乱？帮你把知识点理成一张清晰骨架图

## 前置依赖

```bash
pip install pandas
```

## 核心能力

### 能力1：权威教材和大纲检索

用 `web_search` 搜索目标主题的教材、大纲、知识点列表。

### 能力2：知识框架图生成

运行脚本将知识点结构化，生成Markdown格式的知识框架图。

### 能力3：重点标注和记忆技巧

对关键知识点标注重要程度，生成记忆口诀。

## 命令列表

| 命令 | 说明 | 用法 |
|------|------|------|
| `build` | 构建知识框架 | `python3 scripts/knowledge_framework_tool.py build --topic TOPIC --chapters CHAPTERS --output PATH` |
| `export` | 导出为其他格式 | `python3 scripts/knowledge_framework_tool.py export --input PATH --format md|json` |

## 使用流程

### 步骤 1：明确主题

询问用户要梳理的知识主题，如「高等数学-微积分」「Python基础」「经济学原理」

### 步骤 2：检索权威资料

执行搜索获取知识结构：
```
web_search("[主题] 知识点大纲 目录")
web_search("[主题] 教材推荐 核心概念")
web_search("[主题] 思维导图 知识框架")
```
用 `web_fetch` 抓取教材目录页面，提取章节结构。

### 步骤 3：运行脚本结构化知识点

```bash
python3 scripts/knowledge_framework_tool.py build \
  --topic "主题名称" \
  --chapters "第1章:子主题1|子主题2,第2章:子主题3|子主题4" \
  --output "/path/to/knowledge_framework.md"
```

### 步骤 4：标注重点和记忆技巧

在生成的框架基础上：
1. 用 `web_search` 搜索「[主题] 重点难点 易错点」
2. 在框架中用 ⭐ 标注高频考点
3. 为难点添加记忆口诀或类比说明
4. 用 `write_to_file` 保存最终文件

## 输出格式

输出格式：Markdown 思维导图 + 标注文档

## 验收标准

- ✅ 生成框架图
- ✅ 标注重点
- ✅ 提供记忆技巧
- ✅ 可打印

## 场景化适配

根据学习阶段调整深度

## 注意事项

- 所有数据必须来自真实搜索结果或用户提供的文件，**严禁编造数据**
- 数据缺失时标注"数据不可用"而非猜测
- 输出必须保存为文件（`write_to_file`），不能只在对话中输出
- 建议结合人工判断使用，AI 分析仅供参考
