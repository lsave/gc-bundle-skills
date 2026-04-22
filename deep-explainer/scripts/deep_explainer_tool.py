#!/usr/bin/env python3
"""
深度内容讲解 — 工具脚本
遇到难题卡壳想放弃？帮你拆解卡点，找到破局思路

目标用户: 初学者
输出产物: 讲解文档、延伸阅读清单
"""

import sys, json, os, argparse
from datetime import datetime

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")


def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)


import json
from datetime import datetime

def cmd_explain(args):
    """生成讲解文档框架"""
    topic = args.topic
    level = getattr(args, "level", "beginner") or "beginner"
    output = args.output or f"explain_{datetime.now().strftime('%Y%m%d')}.md"
    
    level_map = {"beginner": "入门", "intermediate": "进阶", "advanced": "高级"}
    
    md = f"""# 📖 深度讲解：{topic}

**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}
**难度级别**: {level_map.get(level, level)}

## 一句话解释

> [用最简单的一句话说明{topic}是什么]

## 通俗理解

[用大白话解释，不用任何专业术语，假设读者是完全的外行人]

## 生活类比

### 类比1：[类比名]
[用一个日常生活的场景来类比{topic}]

### 类比2：[类比名]  
[用另一个角度的类比]

## 具体例子

### 例子1（最简单）
[一个最基础的例子]

### 例子2（稍复杂）
[一个稍微复杂的应用场景]

### 例子3（实际应用）
[一个真实世界的应用案例]

### 例子4（进阶）
[一个需要更深理解的例子]

### 例子5（综合）
[一个综合性的例子]

## ⚠️ 常见误区

| 误区 | 正确理解 |
|------|---------|
| [常见错误理解1] | [正确说法] |
| [常见错误理解2] | [正确说法] |

## 📚 延伸阅读

| 类型 | 名称 | 链接 | 推荐理由 |
|------|------|------|---------|
| 视频 | [待搜索] | [web_search获取] | |
| 文章 | [待搜索] | [web_search获取] | |
| 教材 | [待搜索] | [web_search获取] | |
"""
    
    with open(output, "w", encoding="utf-8") as f:
        f.write(md)
    
    print(json.dumps({"status": "success", "output_file": output, "topic": topic}, ensure_ascii=False, indent=2))
    return 0


def cmd_status(args):
    """查看当前状态"""
    data_files = []
    if os.path.exists(DATA_DIR):
        data_files = [f for f in os.listdir(DATA_DIR) if not f.startswith(".")]
    result = {
        "skill": "deep-explainer",
        "scene": "深度内容讲解",
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
    parser = argparse.ArgumentParser(description="深度内容讲解")
    subparsers = parser.add_subparsers(dest="command", help="可用命令")
    
    p_explain = subparsers.add_parser("explain", help="生成讲解文档框架")
    p_explain.add_argument("--topic", help="TOPIC")
    p_explain.add_argument("--level", help="beginner|intermediate|advanced")
    p_explain.add_argument("--output", help="PATH")

    subparsers.add_parser("status", help="查看状态")
    p_export = subparsers.add_parser("export", help="导出结果")
    p_export.add_argument("format", nargs="?", default="json", help="导出格式")

    args = parser.parse_args()

    if args.command == "explain":
        return cmd_explain(args)
    elif args.command == "status":
        return cmd_status(args)
    elif args.command == "export":
        return cmd_export(args)
    else:
        parser.print_help()
        return 1


if __name__ == "__main__":
    sys.exit(main())
