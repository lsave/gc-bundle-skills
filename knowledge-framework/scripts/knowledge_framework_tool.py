#!/usr/bin/env python3
"""
知识点框架梳理 — 工具脚本
看书越看越乱？帮你把知识点理成一张清晰骨架图

目标用户: 学生
输出产物: 知识框架图、重点标注文档
"""

import sys, json, os, argparse
from datetime import datetime
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")


def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)


import json
from datetime import datetime

def cmd_build(args):
    """构建知识框架Markdown文件"""
    topic = args.topic
    chapters_raw = args.chapters or ""
    output = args.output or f"knowledge_framework_{datetime.now().strftime('%Y%m%d')}.md"
    
    # Parse chapters: "第1章:子主题1|子主题2,第2章:子主题3"
    chapters = []
    for ch in chapters_raw.split(","):
        ch = ch.strip()
        if ":" in ch:
            name, subs = ch.split(":", 1)
            subtopics = [s.strip() for s in subs.split("|") if s.strip()]
            chapters.append({"name": name.strip(), "subtopics": subtopics})
        elif ch:
            chapters.append({"name": ch, "subtopics": []})
    
    # Generate Markdown framework
    md = f"""# 🧠 {topic} — 知识框架图

**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}
**主题**: {topic}
**章节数**: {len(chapters)}

---

"""
    for i, ch in enumerate(chapters, 1):
        md += f"## {i}. {ch['name']}\n\n"
        for j, sub in enumerate(ch["subtopics"], 1):
            md += f"### {i}.{j} {sub}\n\n"
            md += f"**核心概念**: [待补充]\n\n"
            md += f"**重要程度**: ⬜普通 / ⭐重点 / ⭐⭐高频考点\n\n"
            md += f"**记忆技巧**: [待补充]\n\n"
            md += f"**易错点**: [待补充]\n\n"
            md += "---\n\n"
    
    md += """## 📝 总结

### 核心公式/概念速查表
| 编号 | 公式/概念 | 应用场景 | 重要程度 |
|------|----------|---------|---------|
| 1 | [待补充] | [待补充] | ⭐ |

### 记忆口诀
[待补充]

### 推荐教材与资料
| 资料名称 | 链接 | 类型 |
|---------|------|------|
| [待补充] | [web_search获取] | 教材/视频/习题 |
"""
    
    with open(output, "w", encoding="utf-8") as f:
        f.write(md)
    
    print(json.dumps({"status": "success", "output_file": output, "chapters": len(chapters), "message": f"知识框架已生成: {output}"}, ensure_ascii=False, indent=2))
    return 0

def cmd_export(args):
    """导出为其他格式"""
    with open(args.input, "r", encoding="utf-8") as f:
        content = f.read()
    fmt = args.format or "json"
    if fmt == "json":
        # Simple extraction of headers
        import re
        headers = re.findall(r"^##+ (.+)$", content, re.MULTILINE)
        result = {"topic": headers[0] if headers else "", "sections": headers}
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(content)
    return 0


def cmd_status(args):
    """查看当前状态"""
    data_files = []
    if os.path.exists(DATA_DIR):
        data_files = [f for f in os.listdir(DATA_DIR) if not f.startswith(".")]
    result = {
        "skill": "knowledge-framework",
        "scene": "知识点框架梳理",
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
    parser = argparse.ArgumentParser(description="知识点框架梳理")
    subparsers = parser.add_subparsers(dest="command", help="可用命令")
    
    p_build = subparsers.add_parser("build", help="构建知识框架")
    p_build.add_argument("--topic", help="TOPIC")
    p_build.add_argument("--chapters", help="CHAPTERS")
    p_build.add_argument("--output", help="PATH")

    p_export = subparsers.add_parser("export", help="导出为其他格式")
    p_export.add_argument("--input", help="PATH")
    p_export.add_argument("--format", help="md|json")

    subparsers.add_parser("status", help="查看状态")

    args = parser.parse_args()

    if args.command == "build":
        return cmd_build(args)
    if args.command == "export":
        return cmd_export(args)
    elif args.command == "status":
        return cmd_status(args)
    else:
        parser.print_help()
        return 1


if __name__ == "__main__":
    sys.exit(main())
