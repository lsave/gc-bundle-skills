#!/usr/bin/env python3
"""
制定轻量学习规划 — 工具脚本
一到备考就头大？把你的学习计划烂摊子全甩给我吧

目标用户: 备考学生
输出产物: Excel 计划表、进度追踪表
"""

import sys, json, os, argparse
from datetime import datetime
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")


def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)


import pandas as pd
from datetime import datetime, timedelta
import json

def cmd_generate(args):
    """生成Excel学习计划表"""
    subject = args.subject
    days = int(args.days)
    hours = float(args.hours_per_day)
    output = args.output or f"study_plan_{subject}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    
    # Generate daily plan
    daily_plan = []
    start_date = datetime.now().date()
    for d in range(days):
        date = start_date + timedelta(days=d)
        # Split hours into morning/afternoon/evening sessions
        sessions = []
        remaining = hours
        if remaining >= 1.5:
            sessions.append({"时间段": "上午 9:00-10:30", "时长": 1.5, "内容": f"Day{d+1} - 核心知识学习", "材料": "[待补充搜索结果]", "状态": "⬜ 未完成"})
            remaining -= 1.5
        if remaining >= 1:
            sessions.append({"时间段": "下午 14:00-15:00", "时长": 1.0, "内容": f"Day{d+1} - 真题练习", "材料": "[待补充搜索结果]", "状态": "⬜ 未完成"})
            remaining -= 1
        if remaining >= 0.5:
            sessions.append({"时间段": "晚上 20:00-20:30", "时长": 0.5, "内容": f"Day{d+1} - 复习巩固", "材料": "当日笔记", "状态": "⬜ 未完成"})
        for s in sessions:
            daily_plan.append({"日期": date.strftime("%Y-%m-%d"), "星期": ["周一","周二","周三","周四","周五","周六","周日"][date.weekday()], **s})
    
    df_plan = pd.DataFrame(daily_plan)
    
    # Resource list template
    df_resources = pd.DataFrame({
        "序号": range(1, 11),
        "资料名称": [f"[资料{i}]" for i in range(1, 11)],
        "链接": ["[web_search获取]"] * 10,
        "类型": ["真题","真题","真题","APP","APP","APP","教材","视频","笔记","其他"],
        "优先级": ["P0","P0","P0","P1","P1","P1","P0","P1","P2","P2"],
    })
    
    # Progress tracking
    df_progress = pd.DataFrame({
        "日期": [(start_date + timedelta(days=d)).strftime("%Y-%m-%d") for d in range(days)],
        "计划任务数": [len([s for s in daily_plan if s["日期"] == (start_date + timedelta(days=d)).strftime("%Y-%m-%d")]) for d in range(days)],
        "完成任务数": [0] * days,
        "完成率": ["0%"] * days,
        "备注": [""] * days,
    })
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_plan.to_excel(writer, sheet_name="每日计划", index=False)
        df_resources.to_excel(writer, sheet_name="资料清单", index=False)
        df_progress.to_excel(writer, sheet_name="进度追踪", index=False)
    
    result = {
        "status": "success",
        "output_file": output,
        "subject": subject,
        "total_days": days,
        "total_sessions": len(daily_plan),
        "hours_per_day": hours,
        "message": f"学习计划已生成: {output}",
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0

def cmd_track(args):
    """更新学习进度"""
    import openpyxl
    file_path = args.file
    date = args.date
    completed = int(args.completed)
    
    wb = openpyxl.load_workbook(file_path)
    ws = wb["进度追踪"]
    for row in ws.iter_rows(min_row=2, values_only=False):
        if row[0].value == date:
            row[2].value = completed
            planned = row[1].value or 1
            row[3].value = f"{completed/planned*100:.0f}%"
    wb.save(file_path)
    print(json.dumps({"status": "success", "message": f"已更新 {date} 进度: {completed}项完成"}, ensure_ascii=False))
    return 0

def cmd_report(args):
    """生成学习周报"""
    df = pd.read_excel(args.file, sheet_name="进度追踪")
    total_planned = df["计划任务数"].sum()
    total_done = df["完成任务数"].sum()
    rate = total_done / total_planned * 100 if total_planned > 0 else 0
    
    report = f"""# 📚 学习周报

**统计周期**: {df["日期"].iloc[0]} ~ {df["日期"].iloc[-1]}
**总完成率**: {rate:.1f}% ({total_done}/{total_planned})

## 每日完成情况
| 日期 | 计划 | 完成 | 完成率 |
|------|------|------|--------|
"""
    for _, row in df.iterrows():
        report += f"| {row['日期']} | {row['计划任务数']} | {row['完成任务数']} | {row['完成率']} |\n"
    
    output = args.output or "study_weekly_report.md"
    with open(output, "w", encoding="utf-8") as f:
        f.write(report)
    print(json.dumps({"status": "success", "output_file": output, "completion_rate": f"{rate:.1f}%"}, ensure_ascii=False))
    return 0


def cmd_status(args):
    """查看当前状态"""
    data_files = []
    if os.path.exists(DATA_DIR):
        data_files = [f for f in os.listdir(DATA_DIR) if not f.startswith(".")]
    result = {
        "skill": "study-planner",
        "scene": "制定轻量学习规划",
        "data_dir": DATA_DIR,
        "data_files": data_files,
        "file_count": len(data_files),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def cmd_export(args):
    """导出结果"""
    fmt = getattr(args, "format", "json") or "json"
    data_files = []
    if os.path.exists(DATA_DIR):
        data_files = [os.path.join(DATA_DIR, f) for f in os.listdir(DATA_DIR) if not f.startswith(".")]
    
    if fmt == "json":
        output = json.dumps({"files": data_files, "count": len(data_files)}, ensure_ascii=False, indent=2)
    else:
        output = "\n".join(data_files)
    
    print(output)
    return 0


def main():
    parser = argparse.ArgumentParser(description="制定轻量学习规划")
    subparsers = parser.add_subparsers(dest="command", help="可用命令")
    
    p_generate = subparsers.add_parser("generate", help="生成Excel学习计划表")
    p_generate.add_argument("--subject", help="SUBJECT")
    p_generate.add_argument("--days", help="N")
    p_generate.add_argument("--hours-per-day", help="N")
    p_generate.add_argument("--output", help="PATH")

    p_track = subparsers.add_parser("track", help="更新学习进度")
    p_track.add_argument("--date", help="YYYY-MM-DD")
    p_track.add_argument("--completed", help="ITEMS")
    p_track.add_argument("--file", help="PATH")

    p_report = subparsers.add_parser("report", help="生成学习周报")
    p_report.add_argument("--file", help="PATH")
    p_report.add_argument("--output", help="REPORT_PATH")

    subparsers.add_parser("status", help="查看状态")
    p_export = subparsers.add_parser("export", help="导出结果")
    p_export.add_argument("format", nargs="?", default="json", help="导出格式")

    args = parser.parse_args()

    if args.command == "generate":
        return cmd_generate(args)
    if args.command == "track":
        return cmd_track(args)
    if args.command == "report":
        return cmd_report(args)
    elif args.command == "status":
        return cmd_status(args)
    elif args.command == "export":
        return cmd_export(args)
    else:
        parser.print_help()
        return 1


if __name__ == "__main__":
    sys.exit(main())
