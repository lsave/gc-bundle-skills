#!/usr/bin/env python3
"""
拆解学习目标与督促打卡 — 工具脚本
单词背了就忘？每天帮你盯进度，到点就催你学

目标用户: 英语学习者
输出产物: 学习记录表、周报文件
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

def cmd_init(args):
    """初始化学习记录文件"""
    goal = args.goal
    target = args.daily_target
    output = args.output or f"learning_log_{datetime.now().strftime('%Y%m%d')}.xlsx"
    
    # Create 30-day plan
    days = 30
    start = datetime.now().date()
    records = []
    for d in range(days):
        date = start + timedelta(days=d)
        records.append({
            "日期": date.strftime("%Y-%m-%d"),
            "星期": ["周一","周二","周三","周四","周五","周六","周日"][date.weekday()],
            "今日目标": target,
            "完成内容": "",
            "完成度": "0%",
            "学习时长(分钟)": 0,
            "心得笔记": "",
        })
    
    df = pd.DataFrame(records)
    summary = pd.DataFrame({
        "指标": ["学习目标", "每日目标", "开始日期", "计划天数", "累计完成天数", "总学习时长(小时)"],
        "值": [goal, target, start.strftime("%Y-%m-%d"), days, 0, 0],
    })
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="概况", index=False)
        df.to_excel(writer, sheet_name="每日记录", index=False)
    
    print(json.dumps({"status": "success", "output_file": output, "message": f"学习记录已初始化，共{days}天计划"}, ensure_ascii=False, indent=2))
    return 0

def cmd_log(args):
    """记录每日学习"""
    import openpyxl
    file_path = args.file
    date = args.date if args.date != "today" else datetime.now().strftime("%Y-%m-%d")
    completed = args.completed
    
    wb = openpyxl.load_workbook(file_path)
    ws = wb["每日记录"]
    updated = False
    for row in ws.iter_rows(min_row=2, values_only=False):
        if row[0].value == date:
            row[3].value = completed
            row[4].value = "100%"
            row[5].value = 30  # default 30min
            updated = True
    wb.save(file_path)
    msg = f"已记录 {date} 学习: {completed}" if updated else f"未找到日期 {date}"
    print(json.dumps({"status": "success" if updated else "warning", "message": msg}, ensure_ascii=False))
    return 0

def cmd_report(args):
    """生成学习周报"""
    df = pd.read_excel(args.file, sheet_name="每日记录")
    completed_days = len(df[df["完成度"] != "0%"])
    total_minutes = df["学习时长(分钟)"].sum()
    
    report = f"""# 📖 学习周报

**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}
**已坚持**: {completed_days} 天
**总学习时长**: {total_minutes/60:.1f} 小时

## 每日记录
| 日期 | 完成内容 | 完成度 | 学习时长 |
|------|---------|--------|---------|
"""
    for _, row in df.iterrows():
        if row["完成内容"]:
            report += f"| {row['日期']} | {row['完成内容']} | {row['完成度']} | {row['学习时长(分钟)']}min |\n"
    
    output = args.output or "weekly_report.md"
    with open(output, "w", encoding="utf-8") as f:
        f.write(report)
    print(json.dumps({"status": "success", "output_file": output}, ensure_ascii=False))
    return 0


def cmd_status(args):
    """查看当前状态"""
    data_files = []
    if os.path.exists(DATA_DIR):
        data_files = [f for f in os.listdir(DATA_DIR) if not f.startswith(".")]
    result = {
        "skill": "learning-tracker",
        "scene": "拆解学习目标与督促打卡",
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
    parser = argparse.ArgumentParser(description="拆解学习目标与督促打卡")
    subparsers = parser.add_subparsers(dest="command", help="可用命令")
    
    p_init = subparsers.add_parser("init", help="初始化学习记录")
    p_init.add_argument("--goal", help="GOAL")
    p_init.add_argument("--daily-target", help="TARGET")
    p_init.add_argument("--output", help="PATH")

    p_log = subparsers.add_parser("log", help="记录每日学习")
    p_log.add_argument("--file", help="PATH")
    p_log.add_argument("--date", help="DATE")
    p_log.add_argument("--completed", help="ITEMS")

    p_report = subparsers.add_parser("report", help="生成周报")
    p_report.add_argument("--file", help="PATH")
    p_report.add_argument("--output", help="REPORT_PATH")

    subparsers.add_parser("status", help="查看状态")
    p_export = subparsers.add_parser("export", help="导出结果")
    p_export.add_argument("format", nargs="?", default="json", help="导出格式")

    args = parser.parse_args()

    if args.command == "init":
        return cmd_init(args)
    if args.command == "log":
        return cmd_log(args)
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
