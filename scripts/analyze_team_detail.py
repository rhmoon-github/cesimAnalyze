#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
专门针对单个队伍的详细分析脚本
生成"做大做强队"的深度分析报告
"""

import sys
from pathlib import Path
from datetime import datetime

# 添加utils目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent / 'utils'))

from utils_data_analysis import (
    read_excel_data, get_metric_value, find_metric
)

def get_metric_with_priority(metrics_dict, metric_name, team):
    """使用优先级列表获取指标值"""
    metric_priorities = {
        '销售额': ['销售额合计', '本地销售额', '当地销售额', '销售额'],
        '净利润': ['本回合利润', '税后利润', '净利润'],
        '现金': ['现金及等价物', '现金 31.12.', '现金 1.1.', '现金'],
        '短期贷款': ['短期贷款（无计划）', '短期贷款'],
        '长期贷款': ['长期贷款'],
    }
    priority_list = metric_priorities.get(metric_name, [metric_name])
    return get_metric_value(metrics_dict, priority_list, team)

def analyze_team_detailed(team_name, input_dir, output_dir):
    """生成单个队伍的详细分析报告"""
    
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 读取数据文件
    all_rounds_data = {}
    
    # ir00
    ir00_path = input_dir / 'results-ir00.xls'
    if ir00_path.exists():
        metrics_dict, teams = read_excel_data(str(ir00_path))
        all_rounds_data['ir00'] = metrics_dict
    
    # pr01 (r01)
    r01_path = input_dir / 'results-r01.xls'
    if not r01_path.exists():
        r01_path = input_dir / 'results-pr01.xls'
    if r01_path.exists():
        metrics_dict, teams = read_excel_data(str(r01_path))
        all_rounds_data['pr01'] = metrics_dict
    
    if team_name not in teams:
        print(f"错误: 未找到队伍 '{team_name}'")
        print(f"可用队伍: {', '.join(teams)}")
        return
    
    # 生成报告
    report = []
    report.append(f"# {team_name} 详细分析报告\n")
    report.append(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    report.append("=" * 80 + "\n")
    
    # 一、关键指标对比
    report.append("\n## 一、关键指标多回合对比\n")
    
    rounds_order = ['ir00', 'pr01']
    available_rounds = [r for r in rounds_order if r in all_rounds_data]
    
    report.append("### 1.1 财务核心指标\n")
    report.append("| 指标 | " + " | ".join([r.upper() for r in available_rounds]) + " | 变化 |")
    report.append("|------|" + "|".join(["------" for _ in available_rounds]) + "|------|")
    
    metrics_to_analyze = [
        ('销售额', '销售额'),
        ('净利润', '净利润'),
        ('现金', '现金'),
        ('权益合计', '权益合计'),
        ('总资产', '总资产'),
        ('短期贷款', '短期贷款'),
        ('长期贷款', '长期贷款'),
        ('负债合计', ['负债合计', '负债总计']),
    ]
    
    for metric_display, metric_name in metrics_to_analyze:
        values = []
        for rnd in available_rounds:
            metrics_dict = all_rounds_data[rnd]
            if isinstance(metric_name, list):
                val = get_metric_value(metrics_dict, metric_name, team_name)
            elif metric_display in ['销售额', '净利润', '现金']:
                val = get_metric_with_priority(metrics_dict, metric_display, team_name)
            else:
                val = get_metric_value(metrics_dict, metric_name, team_name)
            
            if val is not None:
                if metric_display == '现金':
                    values.append(f"${val/1000:.0f}k")
                elif metric_display in ['销售额', '净利润', '权益合计', '总资产', '短期贷款', '长期贷款', '负债合计']:
                    values.append(f"{val/1000:.0f}k")
                else:
                    values.append(f"{val:.2f}")
            else:
                values.append("N/A")
        
        # 计算变化
        if len(available_rounds) >= 2 and values[0] != "N/A" and values[1] != "N/A":
            try:
                val0 = float(values[0].replace('$', '').replace('k', '').replace(',', ''))
                val1 = float(values[1].replace('$', '').replace('k', '').replace(',', ''))
                if val0 != 0:
                    change = ((val1 - val0) / abs(val0)) * 100
                    change_str = f"{change:+.1f}%"
                else:
                    change_str = "N/A"
            except:
                change_str = "-"
        else:
            change_str = "-"
        
        report.append(f"| {metric_display} | " + " | ".join(values) + f" | {change_str} |")
    
    # 二、财务健康度分析
    report.append("\n\n## 二、财务健康度深度分析\n")
    
    if 'pr01' in all_rounds_data:
        metrics_dict = all_rounds_data['pr01']
        
        # 现金储备
        cash = get_metric_with_priority(metrics_dict, '现金', team_name) or 0
        report.append(f"### 2.1 现金储备分析\n")
        report.append(f"- **当前现金**: ${cash/1000:.0f}k\n")
        
        if cash < 100000:
            status = "🔴 危险（<$100k）"
        elif cash < 300000:
            status = "🟡 预警（<$300k）"
        else:
            status = "🟢 安全（≥$300k）"
        report.append(f"- **状态**: {status}\n")
        
        # 净债务/权益比
        equity = get_metric_value(metrics_dict, '权益合计', team_name) or 0
        short_debt = get_metric_value(metrics_dict, '短期贷款', team_name) or 0
        long_debt = get_metric_value(metrics_dict, '长期贷款', team_name) or 0
        
        if equity > 0:
            net_debt = (short_debt + long_debt) - cash
            debt_equity_ratio = (net_debt / equity) * 100
            report.append(f"\n### 2.2 债务结构分析\n")
            report.append(f"- **权益合计**: ${equity/1000:.0f}k\n")
            report.append(f"- **短期贷款**: ${short_debt/1000:.0f}k\n")
            report.append(f"- **长期贷款**: ${long_debt/1000:.0f}k\n")
            report.append(f"- **净债务**: ${net_debt/1000:.0f}k\n")
            report.append(f"- **净债务/权益比**: {debt_equity_ratio:.1f}%\n")
            
            if debt_equity_ratio < 30:
                debt_status = "🟢 安全（<30%）"
            elif debt_equity_ratio <= 70:
                debt_status = "🟡 预警（30-70%）"
            else:
                debt_status = "🔴 危险（>70%）"
            report.append(f"- **状态**: {debt_status}\n")
        
        # EBITDA率
        ebitda = get_metric_value(metrics_dict, 'EBITDA', team_name)
        if ebitda is None:
            ebitda = get_metric_value(metrics_dict, '息税折旧及摊销前利润', team_name) or 0
        else:
            ebitda = ebitda or 0
        
        sales = get_metric_with_priority(metrics_dict, '销售额', team_name) or 0
        profit = get_metric_with_priority(metrics_dict, '净利润', team_name) or 0
        
        report.append(f"\n### 2.3 盈利能力分析\n")
        report.append(f"- **销售额**: ${sales/1000:.0f}k\n")
        report.append(f"- **净利润**: ${profit/1000:.0f}k\n")
        
        if sales > 0:
            profit_margin = (profit / sales) * 100
            report.append(f"- **净利润率**: {profit_margin:.2f}%\n")
            
            ebitda_rate = (ebitda / sales) * 100
            report.append(f"- **EBITDA率**: {ebitda_rate:.4f}%\n")
            
            if ebitda_rate > 20:
                ebitda_status = "🟢 优秀（>20%）"
            elif ebitda_rate >= 5:
                ebitda_status = "🟡 一般（5-20%）"
            else:
                ebitda_status = "🔴 危险（<5%）"
            report.append(f"- **EBITDA状态**: {ebitda_status}\n")
        
        # 权益比率
        assets = get_metric_value(metrics_dict, '总资产', team_name) or 0
        if assets > 0 and equity > 0:
            equity_ratio = (equity / assets) * 100
            report.append(f"\n### 2.4 资本结构分析\n")
            report.append(f"- **总资产**: ${assets/1000:.0f}k\n")
            report.append(f"- **权益比率**: {equity_ratio:.1f}%\n")
            
            if equity_ratio > 100:
                equity_status = "🟢 安全（>100%）"
            elif equity_ratio >= 50:
                equity_status = "🟡 预警（50-100%）"
            else:
                equity_status = "🔴 危险（<50%）"
            report.append(f"- **状态**: {equity_status}\n")
    
    # 三、行业对比分析
    report.append("\n\n## 三、行业对比分析\n")
    
    if 'pr01' in all_rounds_data:
        metrics_dict = all_rounds_data['pr01']
        
        # 收集所有队伍的数据进行对比
        all_teams_sales = {}
        all_teams_profit = {}
        all_teams_cash = {}
        
        for team in teams:
            sales_val = get_metric_with_priority(metrics_dict, '销售额', team)
            profit_val = get_metric_with_priority(metrics_dict, '净利润', team)
            cash_val = get_metric_with_priority(metrics_dict, '现金', team)
            
            if sales_val is not None:
                all_teams_sales[team] = sales_val
            if profit_val is not None:
                all_teams_profit[team] = profit_val
            if cash_val is not None:
                all_teams_cash[team] = cash_val
        
        # 销售额排名
        if all_teams_sales:
            sorted_sales = sorted(all_teams_sales.items(), key=lambda x: x[1], reverse=True)
            sales_rank = next((i+1 for i, (t, _) in enumerate(sorted_sales) if t == team_name), None)
            sales_rank_total = len(sorted_sales)
            
            report.append(f"### 3.1 销售额排名\n")
            report.append(f"- **当前排名**: 第{sales_rank}位 / 共{sales_rank_total}支队伍\n")
            if sales_rank:
                team_sales = all_teams_sales[team_name]
                if sales_rank > 1:
                    prev_team, prev_sales = sorted_sales[sales_rank - 2]
                    gap = prev_sales - team_sales
                    report.append(f"- **距离上一名差距**: ${gap/1000:.0f}k ({prev_team})\n")
                if sales_rank < sales_rank_total:
                    next_team, next_sales = sorted_sales[sales_rank]
                    gap = team_sales - next_sales
                    report.append(f"- **领先下一名优势**: ${gap/1000:.0f}k ({next_team})\n")
        
        # 净利润排名
        if all_teams_profit:
            sorted_profit = sorted(all_teams_profit.items(), key=lambda x: x[1], reverse=True)
            profit_rank = next((i+1 for i, (t, _) in enumerate(sorted_profit) if t == team_name), None)
            
            report.append(f"\n### 3.2 净利润排名\n")
            report.append(f"- **当前排名**: 第{profit_rank}位 / 共{len(sorted_profit)}支队伍\n")
        
        # 现金排名
        if all_teams_cash:
            sorted_cash = sorted(all_teams_cash.items(), key=lambda x: x[1], reverse=True)
            cash_rank = next((i+1 for i, (t, _) in enumerate(sorted_cash) if t == team_name), None)
            
            report.append(f"\n### 3.3 现金储备排名\n")
            report.append(f"- **当前排名**: 第{cash_rank}位 / 共{len(sorted_cash)}支队伍\n")
    
    # 四、策略建议
    report.append("\n\n## 四、策略建议与行动方案\n")
    
    if 'pr01' in all_rounds_data:
        metrics_dict = all_rounds_data['pr01']
        
        cash = get_metric_with_priority(metrics_dict, '现金', team_name) or 0
        
        report.append("### 4.1 当前状况评估\n")
        
        if cash < 100000:
            report.append("🔴 **高风险状态** - 需要立即采取行动\n")
            report.append("- 现金储备严重不足，面临流动性危机\n")
            report.append("- 建议进入生存模式\n")
        elif cash < 300000:
            report.append("🟡 **中等风险状态** - 需要谨慎规划\n")
            report.append("- 现金储备低于安全线，需要保留缓冲\n")
            report.append("- 建议维持模式\n")
        else:
            report.append("🟢 **相对安全状态** - 可以考虑扩张\n")
            report.append("- 现金储备充足，有扩张空间\n")
            report.append("- 建议进攻模式\n")
        
        report.append("\n### 4.2 具体行动建议\n")
        
        if cash < 100000:
            report.append("1. **立即停止所有非必要投资**\n")
            report.append("2. **出售闲置产能或资产**\n")
            report.append("3. **削减广告和研发支出**\n")
            report.append("4. **优先偿还高利率债务**\n")
            report.append("5. **寻求融资或合并机会**\n")
        elif cash < 300000:
            report.append("1. **保留现金缓冲（至少70%）**\n")
            report.append("2. **仅进行必要广告投入（20%）**\n")
            report.append("3. **维持现有产能，不扩张**\n")
            report.append("4. **监控竞争对手动态**\n")
            report.append("5. **等待更好的扩张时机**\n")
        else:
            report.append("1. **可以考虑适度扩张产能**\n")
            report.append("2. **增加广告投入抢占市场份额**\n")
            report.append("3. **考虑研发投入提升竞争力**\n")
            report.append("4. **保留20-30%现金作为风险缓冲**\n")
            report.append("5. **评估区域市场进入机会**\n")
    
    # 保存报告
    report_text = "\n".join(report)
    output_file = output_dir / f'{team_name}详细分析报告.md'
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(report_text)
    
    print(f"报告已保存到: {output_file}")
    return output_file

if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='生成单个队伍的详细分析报告')
    parser.add_argument('--team', '-t', type=str, default='做大做强队', help='队伍名称')
    parser.add_argument('--input-dir', '-i', type=str, required=True, help='数据输入目录')
    parser.add_argument('--output-dir', '-o', type=str, required=True, help='报告输出目录')
    
    args = parser.parse_args()
    
    analyze_team_detailed(args.team, args.input_dir, args.output_dir)

