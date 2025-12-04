#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生成指定队伍与其他队伍的差距对比分析报告
"""

import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent.parent / 'utils'))
from utils_data_analysis import read_excel_data, get_metric_value

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

def get_all_rounds_data(input_dir):
    """读取所有回合的数据"""
    input_dir = Path(input_dir)
    all_rounds_data = {}
    
    # ir00
    ir00_path = input_dir / 'results-ir00.xls'
    if ir00_path.exists():
        metrics_dict, teams = read_excel_data(str(ir00_path))
        all_rounds_data['ir00'] = {'metrics': metrics_dict, 'teams': teams}
    
    # pr01/pr02
    for i in range(1, 10):
        r_path = input_dir / f'results-r{i:02d}.xls'
        if not r_path.exists():
            r_path = input_dir / f'results-pr{i:02d}.xls'
        if r_path.exists():
            metrics_dict, teams = read_excel_data(str(r_path))
            all_rounds_data[f'pr{i:02d}'] = {'metrics': metrics_dict, 'teams': teams}
    
    return all_rounds_data

def calculate_metrics(metrics_dict, team):
    """计算关键指标"""
    cash = get_metric_with_priority(metrics_dict, '现金', team) or 0
    sales = get_metric_with_priority(metrics_dict, '销售额', team) or 0
    profit = get_metric_with_priority(metrics_dict, '净利润', team) or 0
    equity = get_metric_value(metrics_dict, '权益合计', team) or 0
    assets = get_metric_value(metrics_dict, '总资产', team) or 0
    short_debt = get_metric_value(metrics_dict, '短期贷款', team) or 0
    long_debt = get_metric_value(metrics_dict, '长期贷款', team) or 0
    
    # EBITDA
    ebitda = get_metric_value(metrics_dict, ['息税折旧及摊销前利润(EBITDA)', '息税折旧及摊销前利润', 'EBITDA'], team)
    # 如果EBITDA值太小（可能是百分比值），设为None
    if ebitda is not None and abs(ebitda) < 100:
        ebitda = None
    if ebitda is None:
        ebitda = 0
    
    # 计算衍生指标
    net_debt = (short_debt + long_debt) - cash
    debt_equity_ratio = (net_debt / equity * 100) if equity > 0 else None
    ebitda_rate = (ebitda / sales * 100) if sales > 0 else None
    profit_margin = (profit / sales * 100) if sales > 0 else None
    equity_ratio = (equity / assets * 100) if assets > 0 else None
    
    return {
        '现金': cash,
        '销售额': sales,
        '净利润': profit,
        '权益合计': equity,
        '总资产': assets,
        '短期贷款': short_debt,
        '长期贷款': long_debt,
        'EBITDA': ebitda,
        '净债务': net_debt,
        '净债务权益比': debt_equity_ratio,
        'EBITDA率': ebitda_rate,
        '净利润率': profit_margin,
        '权益比率': equity_ratio,
    }

def generate_gap_analysis(target_team, input_dir, output_dir):
    """生成差距对比分析报告"""
    
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 读取所有回合数据
    all_rounds_data = get_all_rounds_data(input_dir)
    
    if not all_rounds_data:
        print(f"错误: 未找到数据文件")
        return
    
    # 获取最新回合
    latest_round = None
    for rnd in ['pr09', 'pr08', 'pr07', 'pr06', 'pr05', 'pr04', 'pr03', 'pr02', 'pr01', 'ir00']:
        if rnd in all_rounds_data:
            latest_round = rnd
            break
    
    if latest_round is None:
        latest_round = list(all_rounds_data.keys())[-1]
    
    metrics_dict = all_rounds_data[latest_round]['metrics']
    teams = all_rounds_data[latest_round]['teams']
    
    if target_team not in teams:
        print(f"错误: 未找到队伍 '{target_team}'")
        print(f"可用队伍: {', '.join(teams)}")
        return
    
    # 计算所有队伍的关键指标
    all_teams_metrics = {}
    for team in teams:
        all_teams_metrics[team] = calculate_metrics(metrics_dict, team)
    
    target_metrics = all_teams_metrics[target_team]
    
    # 生成报告
    report = []
    report.append(f"# {target_team} 与其他队伍差距对比分析报告\n")
    report.append(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    report.append(f"分析回合：{latest_round.upper()}\n")
    report.append("=" * 80 + "\n")
    
    # 一、整体排名对比
    report.append("\n## 一、整体排名对比\n")
    
    # 销售额排名
    sales_ranking = sorted(teams, key=lambda t: all_teams_metrics[t]['销售额'], reverse=True)
    sales_rank = sales_ranking.index(target_team) + 1
    report.append(f"### 1.1 销售额排名\n")
    report.append(f"- **{target_team}排名**：第{sales_rank}位 / 共{len(teams)}支队伍\n")
    report.append(f"- **销售额**：${target_metrics['销售额']/1000:.0f}k\n")
    
    if sales_rank > 1:
        prev_team = sales_ranking[sales_rank - 2]
        gap = all_teams_metrics[prev_team]['销售额'] - target_metrics['销售额']
        report.append(f"- **距离上一名差距**：${gap/1000:.0f}k ({prev_team})\n")
    
    if sales_rank < len(teams):
        next_team = sales_ranking[sales_rank]
        gap = target_metrics['销售额'] - all_teams_metrics[next_team]['销售额']
        report.append(f"- **领先下一名优势**：${gap/1000:.0f}k ({next_team})\n")
    
    # 净利润排名
    profit_ranking = sorted(teams, key=lambda t: all_teams_metrics[t]['净利润'], reverse=True)
    profit_rank = profit_ranking.index(target_team) + 1
    report.append(f"\n### 1.2 净利润排名\n")
    report.append(f"- **{target_team}排名**：第{profit_rank}位 / 共{len(teams)}支队伍\n")
    report.append(f"- **净利润**：${target_metrics['净利润']/1000:.0f}k\n")
    
    # 现金排名
    cash_ranking = sorted(teams, key=lambda t: all_teams_metrics[t]['现金'], reverse=True)
    cash_rank = cash_ranking.index(target_team) + 1
    report.append(f"\n### 1.3 现金储备排名\n")
    report.append(f"- **{target_team}排名**：第{cash_rank}位 / 共{len(teams)}支队伍\n")
    report.append(f"- **现金**：${target_metrics['现金']/1000:.0f}k\n")
    
    # 二、与TOP3队伍详细对比
    report.append("\n## 二、与TOP3队伍详细对比\n")
    
    top3_teams = sales_ranking[:3]
    report.append("| 指标 | " + " | ".join([f"{target_team}"] + top3_teams) + " |")
    report.append("|------|" + "|".join(["------" for _ in range(4)]) + "|")
    
    # 关键指标对比
    key_metrics = [
        ('销售额', '销售额', 'k'),
        ('净利润', '净利润', 'k'),
        ('现金', '现金', 'k'),
        ('权益合计', '权益合计', 'k'),
        ('EBITDA', 'EBITDA', 'k'),
        ('EBITDA率', 'EBITDA率', '%'),
        ('净利润率', '净利润率', '%'),
        ('净债务权益比', '净债务权益比', '%'),
        ('权益比率', '权益比率', '%'),
    ]
    
    for metric_name, metric_key, unit in key_metrics:
        # 格式化目标队伍的值
        target_val = target_metrics[metric_key]
        if target_val is not None:
            if unit == 'k':
                values = [f"${target_val/1000:.0f}{unit}"]
            else:
                values = [f"{target_val:.1f}{unit}"]
        else:
            values = ["N/A"]
        
        # 格式化其他队伍的值
        for team in top3_teams:
            val = all_teams_metrics[team][metric_key]
            if val is not None:
                if unit == 'k':
                    values.append(f"${val/1000:.0f}{unit}")
                else:
                    values.append(f"{val:.1f}{unit}")
            else:
                values.append("N/A")
        report.append(f"| {metric_name} | " + " | ".join(values) + " |")
    
    # 三、差距分析
    report.append("\n## 三、关键差距分析\n")
    
    # 与第1名对比
    top1_team = top3_teams[0]
    top1_metrics = all_teams_metrics[top1_team]
    
    report.append(f"### 3.1 与第1名（{top1_team}）的差距\n")
    
    sales_gap = top1_metrics['销售额'] - target_metrics['销售额']
    sales_gap_pct = (sales_gap / top1_metrics['销售额'] * 100) if top1_metrics['销售额'] > 0 else 0
    report.append(f"- **销售额差距**：${sales_gap/1000:.0f}k（差距{sales_gap_pct:.1f}%）\n")
    
    profit_gap = top1_metrics['净利润'] - target_metrics['净利润']
    # 计算差距百分比：如果第1名净利润为正，使用第1名作为基准；否则使用目标队伍作为基准
    if top1_metrics['净利润'] > 0:
        profit_gap_pct = (profit_gap / top1_metrics['净利润'] * 100)
    elif target_metrics['净利润'] > 0:
        profit_gap_pct = (profit_gap / target_metrics['净利润'] * 100)
    else:
        profit_gap_pct = 0
    report.append(f"- **净利润差距**：${profit_gap/1000:.0f}k（差距{profit_gap_pct:.1f}%）\n")
    
    cash_gap = top1_metrics['现金'] - target_metrics['现金']
    cash_gap_pct = (cash_gap / top1_metrics['现金'] * 100) if top1_metrics['现金'] > 0 else 0
    report.append(f"- **现金差距**：${cash_gap/1000:.0f}k（差距{cash_gap_pct:.1f}%）\n")
    
    # 与行业均值对比
    report.append(f"\n### 3.2 与行业均值对比\n")
    
    import numpy as np
    avg_sales = np.mean([all_teams_metrics[t]['销售额'] for t in teams])
    avg_profit = np.mean([all_teams_metrics[t]['净利润'] for t in teams])
    avg_cash = np.mean([all_teams_metrics[t]['现金'] for t in teams])
    
    sales_vs_avg = ((target_metrics['销售额'] - avg_sales) / avg_sales * 100) if avg_sales > 0 else 0
    # 净利润对比：如果行业均值为正，使用均值作为基准；否则使用目标队伍作为基准
    if avg_profit > 0:
        profit_vs_avg = ((target_metrics['净利润'] - avg_profit) / avg_profit * 100)
    elif target_metrics['净利润'] > 0:
        profit_vs_avg = ((target_metrics['净利润'] - avg_profit) / target_metrics['净利润'] * 100)
    else:
        profit_vs_avg = 0
    cash_vs_avg = ((target_metrics['现金'] - avg_cash) / avg_cash * 100) if avg_cash > 0 else 0
    
    report.append(f"- **销售额**：${target_metrics['销售额']/1000:.0f}k（行业均值：${avg_sales/1000:.0f}k，{sales_vs_avg:+.1f}%）\n")
    report.append(f"- **净利润**：${target_metrics['净利润']/1000:.0f}k（行业均值：${avg_profit/1000:.0f}k，{profit_vs_avg:+.1f}%）\n")
    report.append(f"- **现金**：${target_metrics['现金']/1000:.0f}k（行业均值：${avg_cash/1000:.0f}k，{cash_vs_avg:+.1f}%）\n")
    
    # 四、多回合趋势对比
    report.append("\n## 四、多回合趋势对比\n")
    
    rounds_order = ['ir00', 'pr01', 'pr02', 'pr03', 'pr04', 'pr05']
    available_rounds = [r for r in rounds_order if r in all_rounds_data]
    
    if len(available_rounds) > 1:
        report.append("### 4.1 销售额趋势对比\n")
        report.append("| 队伍 | " + " | ".join([r.upper() for r in available_rounds]) + " |")
        report.append("|------|" + "|".join(["------" for _ in available_rounds]) + "|")
        
        # 显示目标队伍和TOP3
        display_teams = [target_team] + top3_teams
        for team in display_teams:
            values = []
            for rnd in available_rounds:
                if rnd in all_rounds_data:
                    metrics = all_rounds_data[rnd]['metrics']
                    sales = get_metric_with_priority(metrics, '销售额', team) or 0
                    values.append(f"${sales/1000:.0f}k")
                else:
                    values.append("N/A")
            report.append(f"| {team} | " + " | ".join(values) + " |")
    
    # 五、改进建议
    report.append("\n## 五、改进建议\n")
    
    report.append("### 5.1 关键改进方向\n")
    
    # 基于差距分析给出建议
    if sales_rank > 3:
        report.append(f"1. **提升销售额**：当前排名第{sales_rank}位，需要提升${sales_gap/1000:.0f}k才能追上第1名\n")
    
    if target_metrics['EBITDA率'] and target_metrics['EBITDA率'] < 20:
        report.append("2. **提升盈利能力**：EBITDA率较低，需要优化成本结构或提升定价\n")
    
    if target_metrics['现金'] < 300000:
        report.append("3. **增加现金储备**：现金储备不足，建议保留更多现金缓冲\n")
    
    if target_metrics['净债务权益比'] and target_metrics['净债务权益比'] > 30:
        report.append("4. **优化债务结构**：净债务/权益比较高，建议降低负债或增加权益\n")
    
    report.append("\n### 5.2 学习对象\n")
    report.append(f"- **销售额标杆**：{top1_team}（${top1_metrics['销售额']/1000:.0f}k）\n")
    
    # 找出盈利能力最强的队伍
    profit_leader = max(teams, key=lambda t: all_teams_metrics[t]['净利润'])
    report.append(f"- **盈利能力标杆**：{profit_leader}（净利润${all_teams_metrics[profit_leader]['净利润']/1000:.0f}k）\n")
    
    # 找出现金最充足的队伍
    cash_leader = max(teams, key=lambda t: all_teams_metrics[t]['现金'])
    report.append(f"- **现金管理标杆**：{cash_leader}（现金${all_teams_metrics[cash_leader]['现金']/1000:.0f}k）\n")
    
    # 保存报告
    report_text = "\n".join(report)
    output_file = output_dir / f'{target_team}差距分析报告.md'
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(report_text)
    
    print(f"报告已保存到: {output_file}")
    return output_file

if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='生成队伍差距对比分析报告')
    parser.add_argument('--team', '-t', type=str, default='做大做强队', help='目标队伍名称')
    parser.add_argument('--input-dir', '-i', type=str, required=True, help='数据输入目录')
    parser.add_argument('--output-dir', '-o', type=str, required=True, help='报告输出目录')
    
    args = parser.parse_args()
    
    generate_gap_analysis(args.team, args.input_dir, args.output_dir)

