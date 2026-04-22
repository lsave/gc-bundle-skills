---
name: study-planner
description: "制定轻量学习规划。一到备考就头大？把你的学习计划烂摊子全甩给我吧 This skill should be used when the user asks about 制定轻量学习规划. Keywords: 制定学习计划, 备考计划."
---

# 制定轻量学习规划

> 一到备考就头大？把你的学习计划烂摊子全甩给我吧

## 前置依赖

```bash
pip install pandas openpyxl
```

## 核心能力

### 能力1：考试大纲和真题检索

用 `web_search` 搜索目标考试的大纲、真题资源，用 `web_fetch` 抓取内容。

### 能力2：Excel学习计划表生成

运行脚本生成包含每日任务、时间安排、材料链接的Excel计划表。

### 能力3：定时提醒创建

用 `automation_update` 创建每日学习提醒定时任务。

## 命令列表

| 命令 | 说明 | 用法 |
|------|------|------|
| `generate` | 生成Excel学习计划表 | `python3 scripts/study_planner_tool.py generate --subject SUBJECT --days N --hours-per-day N --output PATH` |
| `track` | 更新学习进度 | `python3 scripts/study_planner_tool.py track --date YYYY-MM-DD --completed ITEMS --file PATH` |
| `report` | 生成学习周报 | `python3 scripts/study_planner_tool.py report --file PATH --output REPORT_PATH` |

## 使用流程

### 步骤 1：收集学习需求

向用户询问：① 要学什么（考试/技能）② 目标时间（几天/几周）③ 每天可学几小时 ④ 当前水平

### 步骤 2：检索学习资料

执行以下搜索获取真实资料：
```
web_search("[考试名] 历年真题 下载")
web_search("[考试名] 备考APP推荐 2024")
web_search("[考试名] 考试大纲 重点章节")
```
用 `web_fetch` 抓取搜索结果中的资料页面，提取真题链接、APP下载链接、大纲内容。确保至少获取：
- 3套以上真题链接
- 3个以上备考APP推荐
- 考试大纲/重点章节列表

### 步骤 3：运行脚本生成Excel计划表

```bash
python3 scripts/study_planner_tool.py generate \
  --subject "考试/技能名" \
  --days 7 \
  --hours-per-day 3 \
  --output "/path/to/study_plan.xlsx"
```
脚本会生成包含以下sheet的Excel文件：
- 「每日计划」: 日期、时间段、学习内容、材料链接、完成状态
- 「资料清单」: 资料名称、链接、类型、优先级
- 「进度追踪」: 每日完成度统计图表数据

读取脚本输出确认文件已生成。

### 步骤 4：设置每日提醒

用 `automation_update` 创建每日学习提醒：
```
automation_update: rrule=FREQ=DAILY;BYHOUR=8;BYMINUTE=0
prompt="检查学习计划进度，提醒今日学习任务"
```

### 步骤 5：汇总输出给用户

将以下内容整理后展示给用户：
1. Excel计划表文件路径
2. 搜集到的真题链接列表（带超链接）
3. APP推荐列表（带下载链接）
4. 定时提醒已创建的确认信息

最后用 `write_to_file` 生成一份Markdown版学习指南作为备份。

## 输出格式

输出格式：文字 + Excel 文件 + 日程确认 + 资源链接

## 验收标准

- ✅ Excel 计划表可访问
- ✅ 至少 3 套真题链接
- ✅ 至少 3 个 APP 推荐
- ✅ 创建了日程

## 场景化适配

根据年级、教材、考试日期调整

## 依赖 Skills

- **proactive-agent**

## 注意事项

- 所有数据必须来自真实搜索结果或用户提供的文件，**严禁编造数据**
- 数据缺失时标注"数据不可用"而非猜测
- 输出必须保存为文件（`write_to_file`），不能只在对话中输出
- 建议结合人工判断使用，AI 分析仅供参考
