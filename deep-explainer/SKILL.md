---
name: deep-explainer
description: "深度内容讲解。遇到难题卡壳想放弃？帮你拆解卡点，找到破局思路 This skill should be used when the user asks about 深度内容讲解. Keywords: 讲解 XXX, 什么是 XXX."
---

# 深度内容讲解

> 遇到难题卡壳想放弃？帮你拆解卡点，找到破局思路

## 前置依赖

```bash
pip install 
```

## 核心能力

### 能力1：深度内容检索

用 `web_search` 搜索主题相关的权威解释、科普文章、视频教程。

### 能力2：讲解文档生成

运行脚本生成结构化的讲解文档模板。

### 能力3：延伸阅读资源

用 `web_search` 搜索延伸阅读材料并整理成列表。

## 命令列表

| 命令 | 说明 | 用法 |
|------|------|------|
| `explain` | 生成讲解文档框架 | `python3 scripts/deep_explainer_tool.py explain --topic TOPIC --level beginner|intermediate|advanced --output PATH` |

## 使用流程

### 步骤 1：了解讲解需求

询问用户：① 要讲解什么概念/主题 ② 当前理解水平 ③ 应用场景（考试/工作/兴趣）

### 步骤 2：搜索优质讲解资源

```
web_search("[概念] 通俗解释 入门")
web_search("[概念] 生活类比 举例说明")
web_search("[概念] 科普视频 推荐")
web_search("[概念] 延伸阅读 深入理解")
```
提取搜索结果中最好的解释和类比。

### 步骤 3：生成讲解文档

运行脚本生成文档框架：
```bash
python3 scripts/deep_explainer_tool.py explain --topic "概念名" --output "/path/to/explanation.md"
```
然后用 `write_to_file` 在框架基础上填充：
1. 用大白话解释概念（避免专业术语）
2. 至少2个生活类比
3. 3-5个具体例子（从简单到复杂）
4. 常见误区说明

### 步骤 4：附上延伸阅读

在文档末尾附上：
- 科普视频链接（从web_search结果中筛选）
- 入门文章链接
- 进阶教材推荐

## 输出格式

输出格式：讲解文档 + 资源链接

## 验收标准

- ✅ 浅显语言
- ✅ 使用类比
- ✅ 3-5 个例子
- ✅ 延伸阅读资源

## 场景化适配

根据年龄/背景调整讲解深度

## 注意事项

- 所有数据必须来自真实搜索结果或用户提供的文件，**严禁编造数据**
- 数据缺失时标注"数据不可用"而非猜测
- 输出必须保存为文件（`write_to_file`），不能只在对话中输出
- 建议结合人工判断使用，AI 分析仅供参考
