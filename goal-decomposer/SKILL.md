---
name: goal-decomposer
description: "目标拆解与打卡。立的flag总倒？把大目标拆成小步骤，每天打卡追踪进度 This skill should be used when the user asks about 目标拆解与打卡. Keywords: 目标管理, 拆解目标, 打卡."
---

# 目标拆解与打卡

> 立的flag总倒？把大目标拆成小步骤，每天打卡追踪进度

## 前置依赖

```bash
pip install pandas openpyxl
```

## 核心能力

### 能力1：用 SMART 原则拆解大目标为可执行子任务

用 `web_search` 搜索该目标的常见实现路径和最佳实践，用 `automation_update` 创建打卡提醒。

### 能力2：生成月度/周度里程碑计划（write_to_file）

用 `write_to_file` 生成文件。

### 能力3：每日打卡记录进度

用 `web_search` 搜索该目标的常见实现路径和最佳实践，用 `automation_update` 创建打卡提醒。

### 能力4：用 automation_update 创建定期回顾提醒

用 `automation_update` 创建定时任务。

### 能力5：生成进度追踪报告

用 `write_to_file` 生成文件。

## 使用流程

### 步骤 1：收集用户需求

向用户确认以下信息（如果未主动提供）：
- 你的目标是什么？（具体描述）
- 期望完成时间
- 当前进度/起点
- 每天/每周能投入多少时间？

### 步骤 2：运行脚本处理数据

```bash
python3 scripts/goal_decomposer_tool.py run \
  --input "用户提供的输入" \
  --output "/path/to/output_file"
```

读取脚本输出的结果，确认数据处理成功。

### 步骤 3：生成最终产出

基于脚本输出和搜索到的资源，用 `write_to_file` 生成以下文件：

- **目标计划文件**
- **进度追踪表**
- **周/月回顾报告**

输出格式要求：计划文件 + 进度表 + 定时任务确认

### 步骤 4：设置定时任务

用 `automation_update` 创建定时任务：
```
automation_update:
  name="目标拆解与打卡"
  rrule=FREQ=DAILY;BYHOUR=9;BYMINUTE=0
  prompt="执行目标拆解与打卡skill的定时任务"
```

### 步骤 5：汇总交付

向用户展示：
1. 生成的文件路径和内容摘要
2. 搜集到的资源链接列表
3. 关键发现和建议
4. 定时任务创建确认

## 输出格式

```markdown
# 📋 目标拆解与打卡 — 执行报告

**生成时间**: YYYY-MM-DD HH:MM
**目标用户**: 有目标但执行力不足的人、自我管理者

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

- ✅ 目标符合 SMART 原则
- ✅ 子任务可执行
- ✅ 定期回顾已设置
- ✅ 有进度追踪机制

## 场景化适配

根据目标类型（学习/健康/职业）调整拆解方式


## 依赖 Skills

本 Skill 参考以下已有 Skill 的能力进行增强：
- **goal-tracker**

## 注意事项

- 所有数据必须来自 `web_search` / `web_fetch` 的真实搜索结果，**严禁编造数据**
- 数据缺失时标注"数据不可用"而非猜测
- 报告必须保存为文件（`write_to_file`），不能只在对话中输出
- 建议结合人工判断使用，AI 分析仅供参考
