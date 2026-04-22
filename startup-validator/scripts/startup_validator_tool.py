#!/usr/bin/env python3
"""
创业点子可行性验证 — 辅助工具脚本

功能:
  analyze: 生成竞品对比矩阵和差异化分析框架
  report:  生成可行性报告的结构化框架（JSON）
  
用法:
  python3 startup_validator_tool.py analyze --idea "..." --competitors "A:定位1,B:定位2" --market-keywords "kw1,kw2"
  python3 startup_validator_tool.py report --idea "..." --output "report.md"
"""

import sys
import json
import os
import argparse
from datetime import datetime


def cmd_analyze(args):
    """生成竞品对比矩阵和分析框架"""
    idea = args.idea
    competitors_raw = args.competitors or ""
    market_keywords = args.market_keywords or ""
    
    # Parse competitors: "竞品1:定位1,竞品2:定位2"
    competitors = []
    for item in competitors_raw.split(","):
        item = item.strip()
        if ":" in item:
            name, position = item.split(":", 1)
            competitors.append({"name": name.strip(), "position": position.strip()})
        elif item:
            competitors.append({"name": item.strip(), "position": "待补充"})
    
    # Generate analysis framework
    analysis = {
        "meta": {
            "idea": idea,
            "analysis_time": datetime.now().isoformat(),
            "competitor_count": len(competitors),
        },
        "competitor_matrix": {
            "headers": ["竞品名称", "产品定位", "核心功能", "定价模式", "融资情况", "优势", "劣势", "差异化机会"],
            "rows": [
                {
                    "name": c["name"],
                    "position": c["position"],
                    "core_features": "[需web_fetch竞品官网补充]",
                    "pricing": "[需web_fetch补充]",
                    "funding": "[需web_search补充]",
                    "strengths": "[待分析]",
                    "weaknesses": "[待分析]",
                    "differentiation_opportunity": "[待分析]",
                }
                for c in competitors
            ],
        },
        "market_analysis_framework": {
            "tam_keywords": [f"{kw} market size" for kw in market_keywords.split(",") if kw.strip()],
            "sam_keywords": [f"{kw} addressable market" for kw in market_keywords.split(",") if kw.strip()],
            "growth_keywords": [f"{kw} CAGR growth rate" for kw in market_keywords.split(",") if kw.strip()],
        },
        "feasibility_dimensions": [
            {"dimension": "市场需求", "weight": 0.25, "score": None, "evidence": ""},
            {"dimension": "竞争格局", "weight": 0.20, "score": None, "evidence": ""},
            {"dimension": "技术可行性", "weight": 0.20, "score": None, "evidence": ""},
            {"dimension": "商业模式", "weight": 0.15, "score": None, "evidence": ""},
            {"dimension": "团队匹配度", "weight": 0.10, "score": None, "evidence": ""},
            {"dimension": "时机窗口", "weight": 0.10, "score": None, "evidence": ""},
        ],
        "mvp_template": {
            "p0_features": [],
            "p1_features": [],
            "p2_features": [],
            "estimated_dev_weeks": None,
            "tech_stack_suggestion": "",
        },
    }
    
    print(json.dumps(analysis, ensure_ascii=False, indent=2))
    return 0


def cmd_report(args):
    """生成可行性报告的Markdown框架"""
    idea = args.idea
    output_path = args.output or f"feasibility_report_{datetime.now().strftime('%Y%m%d')}.md"
    
    report_template = f"""# 📊 创业可行性分析报告：{idea}

**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}
**分析师**: AI Startup Validator

## 一、执行摘要

[请基于以下分析结果撰写1-2段总结]

## 二、竞品分析

| 竞品 | 定位 | 核心功能 | 定价 | 融资情况 | 优势 | 劣势 |
|------|------|---------|------|---------|------|------|
| [待填充] | | | | | | |

> 💡 使用 web_search + web_fetch 获取真实数据后填充此表

## 三、市场分析

- **TAM (总可用市场)**: [数据来源 + 数值]
- **SAM (可服务市场)**: [数据来源 + 数值]  
- **SOM (可获得市场)**: [估算 + 依据]
- **CAGR (复合增长率)**: [数据来源 + 数值]

## 四、技术可行性评估

| 维度 | 评估 | 说明 |
|------|------|------|
| 技术栈 | | |
| 开发复杂度 | 高/中/低 | |
| 核心技术壁垒 | | |
| 预估开发周期 | X个月 | |

## 五、MVP 功能清单

| 优先级 | 功能模块 | 功能描述 | 实现复杂度 | 预估开发周期 |
|--------|---------|---------|----------|------------|
| P0 | | | | |
| P1 | | | | |
| P2 | | | | |

## 六、实施路线图

| 阶段 | 时间 | 里程碑 | 关键交付物 |
|------|------|--------|----------|
| Phase 1 | 第1-4周 | MVP | |
| Phase 2 | 第5-8周 | Beta | |
| Phase 3 | 第9-12周 | Launch | |

## 七、风险评估

| 风险类型 | 具体风险 | 概率 | 影响程度 | 应对策略 |
|---------|---------|------|---------|---------|
| 市场风险 | | 高/中/低 | 高/中/低 | |
| 技术风险 | | | | |
| 竞争风险 | | | | |
| 资金风险 | | | | |

## 八、综合评分与结论

| 维度 | 评分(1-10) | 权重 | 加权分 |
|------|-----------|------|--------|
| 市场需求 | | 25% | |
| 竞争格局 | | 20% | |
| 技术可行性 | | 20% | |
| 商业模式 | | 15% | |
| 团队匹配度 | | 10% | |
| 时机窗口 | | 10% | |
| **综合评分** | | **100%** | |

### 最终建议

🟢/🟡/🔴 **[建议Go / 谨慎推进 / 不建议]**

[具体理由和下一步建议]

---
⚠️ 免责声明：本报告基于公开数据分析，仅供参考，不构成投资建议。
"""
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(report_template)
    
    result = {
        "status": "success",
        "output_file": output_path,
        "idea": idea,
        "message": f"报告框架已生成到 {output_path}，请用真实数据填充后完善。"
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def main():
    parser = argparse.ArgumentParser(description="创业点子可行性验证工具")
    subparsers = parser.add_subparsers(dest='command', help='可用命令')
    
    # analyze command
    analyze_parser = subparsers.add_parser('analyze', help='生成竞品对比矩阵和分析框架')
    analyze_parser.add_argument('--idea', '-i', required=True, help='创业想法描述')
    analyze_parser.add_argument('--competitors', '-c', help='竞品列表，格式: "竞品1:定位1,竞品2:定位2"')
    analyze_parser.add_argument('--market-keywords', '-m', help='市场搜索关键词，逗号分隔')
    
    # report command
    report_parser = subparsers.add_parser('report', help='生成可行性报告框架')
    report_parser.add_argument('--idea', '-i', required=True, help='创业想法描述')
    report_parser.add_argument('--output', '-o', help='输出文件路径')
    
    args = parser.parse_args()
    
    if args.command == 'analyze':
        return cmd_analyze(args)
    elif args.command == 'report':
        return cmd_report(args)
    else:
        parser.print_help()
        return 1


if __name__ == '__main__':
    sys.exit(main())
