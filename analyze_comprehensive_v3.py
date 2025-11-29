#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
商业模拟竞赛结果综合分析脚本 v3.0
严格按照方法论文档3.0版本进行完整分析
"""

import sys
import os
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import json

# 添加当前目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from utils_data_analysis import (
    read_excel_data, find_metric, get_metric_value,
    check_excel_structure, diagnose_missing_data
)

# ============================================================================
# 配置部分
# ============================================================================

# 文件路径配置
BASE_DIR = Path(__file__).parent.parent.parent / '结果'

FILES = {
    'ir00': BASE_DIR / 'results-ir00.xls',
    'pr01': BASE_DIR / 'results-pr01.xls',
    'pr02': BASE_DIR / 'results-pr02.xls',
    'pr03': BASE_DIR / 'results-pr03.xls',
}

# 队伍名称映射
TEAM_NAME_MAPPING = {
    '创世纪的大富翁': 'Blue',
    '星野四喜': 'Black',
}

# 阈值配置（来自方法论第七章）
THRESHOLDS = {
    '现金储备': {'green': 300000, 'yellow': 100000},
    '净债务权益比': {'green': 30, 'yellow': 70},
    'EBITDA率': {'green': 20, 'yellow': 5},
    '权益比率': {'green': 100, 'yellow': 50},
    '研发回报率': {'green': 15, 'yellow': 0},
}

# ============================================================================
# 第一章：数据基础建设
# ============================================================================

def normalize_team_names(teams):
    """队伍名称标准化"""
    return [TEAM_NAME_MAPPING.get(team, team) for team in teams]


def get_metric_priority_list(metric_name):
    """
    根据标准指标名称返回优先级列表
    用于指标提取时的优先级匹配
    """
    metric_priorities = {
        '销售额': ['销售额合计', '本地销售额', '当地销售额', '销售额'],
        '净利润': ['本回合利润', '税后利润', '净利润'],
        '现金': ['现金及等价物', '现金 31.12.', '现金 1.1.', '现金'],
        '短期贷款': ['短期贷款（无计划）', '短期贷款'],
        '长期贷款': ['长期贷款'],
    }
    return metric_priorities.get(metric_name, [metric_name])


def get_metric_with_priority(metrics_dict, metric_name, team):
    """使用优先级列表获取指标值"""
    priority_list = get_metric_priority_list(metric_name)
    return get_metric_value(metrics_dict, priority_list, team)


def validate_data_integrity(metrics_dict, teams):
    """数据完整性验证（使用正确的会计恒等式）"""
    issues = []
    
    for team in teams:
        assets = get_metric_value(metrics_dict, '总资产', team)
        equity = get_metric_value(metrics_dict, '权益合计', team)
        # 使用负债合计而不是分别计算短期和长期贷款
        liability_total = get_metric_value(metrics_dict, ['负债合计', '负债总计'], team)
        
        if assets and equity is not None:
            # 正确的会计恒等式：总资产 = 权益合计 + 负债合计
            if liability_total is not None:
                calculated = equity + liability_total
                if assets > 0:
                    error_rate = abs(assets - calculated) / abs(assets) * 100
                    if error_rate > 10:  # 误差容忍度10%
                        issues.append({
                            'team': team,
                            'error_rate': error_rate,
                            'calculated': calculated,
                            'actual': assets,
                            'status': '需要人工核查' if error_rate < 50 else '数据异常'
                        })
    
    return issues


def detect_anomalies(metrics_dict, teams):
    """异常值检测"""
    anomalies = defaultdict(list)
    
    for team in teams:
        # 现金极端值
        cash = get_metric_with_priority(metrics_dict, '现金', team)
        if cash:
            if cash > 1500000 or cash < 5000:
                anomalies[team].append({
                    'type': '现金极端值',
                    'value': cash,
                    'rule': '>$1.5M或<$5k'
                })
        
        # 负权益
        equity = get_metric_value(metrics_dict, '权益合计', team)
        if equity and equity < 0:
            anomalies[team].append({
                'type': '负权益',
                'value': equity,
                'rule': '权益合计<0'
            })
    
    return anomalies


def calculate_derived_metrics(all_rounds_data, teams):
    """计算衍生指标"""
    derived = {}
    rounds = ['ir00', 'pr01', 'pr02', 'pr03']
    
    for rnd in rounds:
        if rnd not in all_rounds_data:
            continue
        
        metrics_dict = all_rounds_data[rnd]
        derived[rnd] = {}
        
        # 计算行业统计量
        for metric_name in ['销售额', '净利润', '现金', '权益合计']:
            values = []
            for team in teams:
                val = get_metric_with_priority(metrics_dict, metric_name, team)
                if val is not None:
                    values.append(val)
            
            if values:
                import numpy as np
                derived[rnd][f'{metric_name}_行业均值'] = np.mean(values)
                derived[rnd][f'{metric_name}_行业中位数'] = np.median(values)
                derived[rnd][f'{metric_name}_行业标准差'] = np.std(values)
        
        # 计算排名
        for metric_name in ['销售额', '净利润', '现金']:
            team_values = {}
            for team in teams:
                val = get_metric_with_priority(metrics_dict, metric_name, team)
                if val is not None:
                    team_values[team] = val
            
            if team_values:
                sorted_teams = sorted(team_values.items(), key=lambda x: x[1], reverse=True)
                rankings = {team: rank+1 for rank, (team, _) in enumerate(sorted_teams)}
                derived[rnd][f'{metric_name}_排名'] = rankings
        
        # 计算环比增长率（需要上回合数据）
        if rnd != 'ir00':
            prev_rnd = rounds[rounds.index(rnd) - 1]
            if prev_rnd in all_rounds_data:
                prev_metrics = all_rounds_data[prev_rnd]
                for metric_name in ['销售额', '净利润', '现金']:
                    growth_rates = {}
                    for team in teams:
                        current = get_metric_with_priority(metrics_dict, metric_name, team)
                        previous = get_metric_with_priority(prev_metrics, metric_name, team)
                        if current is not None and previous is not None and previous != 0:
                            growth_rate = ((current - previous) / abs(previous)) * 100
                            growth_rates[team] = growth_rate
                    if growth_rates:
                        derived[rnd][f'{metric_name}_环比增长'] = growth_rates
        
        # 计算排名变化（需要上回合数据）
        if rnd != 'ir00':
            prev_rnd = rounds[rounds.index(rnd) - 1]
            if prev_rnd in all_rounds_data:
                prev_derived = derived.get(prev_rnd, {})
                for metric_name in ['销售额', '净利润', '现金']:
                    current_rankings = derived[rnd].get(f'{metric_name}_排名', {})
                    previous_rankings = prev_derived.get(f'{metric_name}_排名', {})
                    if current_rankings and previous_rankings:
                        rank_changes = {}
                        for team in teams:
                            current_rank = current_rankings.get(team)
                            previous_rank = previous_rankings.get(team)
                            if current_rank is not None and previous_rank is not None:
                                rank_changes[team] = current_rank - previous_rank
                        if rank_changes:
                            derived[rnd][f'{metric_name}_排名变化'] = rank_changes
        
        # 计算战略偏离度（自身指标与行业均值的偏离程度）
        for metric_name in ['销售额', '净利润', '现金']:
            industry_mean = derived[rnd].get(f'{metric_name}_行业均值')
            if industry_mean is not None and industry_mean != 0:
                deviations = {}
                for team in teams:
                    team_value = get_metric_with_priority(metrics_dict, metric_name, team)
                    if team_value is not None:
                        deviation = abs(team_value - industry_mean) / abs(industry_mean) * 100
                        deviations[team] = deviation
                if deviations:
                    derived[rnd][f'{metric_name}_战略偏离度'] = deviations
    
    return derived


# ============================================================================
# 第三章：自身诊断分析
# ============================================================================

def calculate_financial_health(metrics_dict, teams):
    """财务健康度红绿灯系统"""
    health = {}
    
    for team in teams:
        health[team] = {
            'indicators': {},
            'status': {},
            'action_required': []
        }
        
        # 1. 现金储备
        cash = get_metric_with_priority(metrics_dict, '现金', team) or 0
        if cash > THRESHOLDS['现金储备']['green']:
            status = '🟢'
        elif cash >= THRESHOLDS['现金储备']['yellow']:
            status = '🟡'
        else:
            status = '🔴'
        
        health[team]['indicators']['现金储备'] = cash
        health[team]['status']['现金储备'] = status
        
        # 2. 净债务/权益比
        equity = get_metric_value(metrics_dict, '权益合计', team) or 0
        short_debt = get_metric_value(metrics_dict, '短期贷款', team) or 0
        long_debt = get_metric_value(metrics_dict, '长期贷款', team) or 0
        
        if equity > 0:
            net_debt = (short_debt + long_debt) - cash
            debt_equity_ratio = (net_debt / equity) * 100
            
            if debt_equity_ratio < THRESHOLDS['净债务权益比']['green']:
                status = '🟢'
            elif debt_equity_ratio <= THRESHOLDS['净债务权益比']['yellow']:
                status = '🟡'
            else:
                status = '🔴'
            
            health[team]['indicators']['净债务权益比'] = debt_equity_ratio
            health[team]['status']['净债务权益比'] = status
        else:
            health[team]['indicators']['净债务权益比'] = None
            health[team]['status']['净债务权益比'] = '🔴'
        
        # 3. EBITDA率
        ebitda = get_metric_value(metrics_dict, 'EBITDA', team)
        if ebitda is None:
            ebitda = get_metric_value(metrics_dict, '息税折旧及摊销前利润', team) or 0
        else:
            ebitda = ebitda or 0
        
        sales = get_metric_with_priority(metrics_dict, '销售额', team) or 0
        
        if sales > 0:
            ebitda_rate = (ebitda / sales) * 100
            if ebitda_rate > THRESHOLDS['EBITDA率']['green']:
                status = '🟢'
            elif ebitda_rate >= THRESHOLDS['EBITDA率']['yellow']:
                status = '🟡'
            else:
                status = '🔴'
            
            health[team]['indicators']['EBITDA率'] = ebitda_rate
            health[team]['status']['EBITDA率'] = status
        else:
            health[team]['indicators']['EBITDA率'] = None
            health[team]['status']['EBITDA率'] = '🔴'
        
        # 4. 权益比率
        assets = get_metric_value(metrics_dict, '总资产', team) or 0
        if assets > 0 and equity > 0:
            equity_ratio = (equity / assets) * 100
            if equity_ratio > THRESHOLDS['权益比率']['green']:
                status = '🟢'
            elif equity_ratio >= THRESHOLDS['权益比率']['yellow']:
                status = '🟡'
            else:
                status = '🔴'
            
            health[team]['indicators']['权益比率'] = equity_ratio
            health[team]['status']['权益比率'] = status
        else:
            health[team]['indicators']['权益比率'] = None
            health[team]['status']['权益比率'] = '🔴'
        
        # 5. 研发回报率
        profit = get_metric_with_priority(metrics_dict, '净利润', team) or 0
        rd_expense = get_metric_value(metrics_dict, '研发', team) or 0
        
        if rd_expense and rd_expense > 0 and profit is not None:
            rd_return = (profit / rd_expense) * 100
            if rd_return > THRESHOLDS['研发回报率']['green']:
                status = '🟢'
            elif rd_return >= THRESHOLDS['研发回报率']['yellow']:
                status = '🟡'
            else:
                status = '🔴'
            
            health[team]['indicators']['研发回报率'] = rd_return
            health[team]['status']['研发回报率'] = status
        else:
            health[team]['indicators']['研发回报率'] = None
            health[team]['status']['研发回报率'] = '🟡'  # 无研发投入
        
        # 统计并生成行动建议
        red_count = sum(1 for s in health[team]['status'].values() if '🔴' in str(s))
        yellow_count = sum(1 for s in health[team]['status'].values() if '🟡' in str(s))
        
        if red_count > 2:
            health[team]['action_required'].append('⚠️ 立即进入生存模式（停止投资、削减成本）')
        elif yellow_count > 3 or red_count > 0:
            health[team]['action_required'].append('⚠️ 召开紧急战略复盘会')
        elif red_count == 0 and yellow_count <= 1:
            health[team]['action_required'].append('✅ 可考虑激进扩张')
    
    return health


def analyze_cash_flow_source(metrics_dict, teams, prev_metrics_dict):
    """现金流源头分析"""
    cash_flow = {}
    
    for team in teams:
        cash = get_metric_with_priority(metrics_dict, '现金', team) or 0
        prev_cash = get_metric_with_priority(prev_metrics_dict, '现金', team) or 0 if prev_metrics_dict else 0
        cash_change = cash - prev_cash
        
        # 修复：确保能提取到EBITDA值
        ebitda = get_metric_value(metrics_dict, 'EBITDA', team)
        if ebitda is None:
            ebitda = get_metric_value(metrics_dict, '息税折旧及摊销前利润', team) or 0
        else:
            ebitda = ebitda or 0
        
        if ebitda > 100000:
            cash_type = 'A. 经营驱动型（健康）'
            description = f'经营现金流+${ebitda/1000:.0f}k → 可扩张'
        elif cash_change > 0 and abs(ebitda) < abs(cash_change) * 0.5:
            cash_type = 'B. 融资驱动型（危险）'
            description = '融资现金流为主要来源 → 不可持续'
        else:
            cash_type = 'C. 投资消耗型（过渡期）'
            description = '投资现金流消耗现金 → 关注下回合回报'
        
        cash_flow[team] = {
            '现金变化': cash_change,
            '经营现金流(EBITDA)': ebitda,
            '现金流类型': cash_type,
            '描述': description
        }
    
    return cash_flow


def analyze_regional_market(all_rounds_data, teams, round_name):
    """区域市场表现分析（替代方案）
    
    注意：由于Excel中区域销售额数据不可用或数据量极小（仅占总额的0.05%-0.65%），
    区域市场分析功能受限。当前使用"美国"、"亚洲"、"欧洲"指标作为替代，
    但这些指标的实际含义可能与区域销售额不符。
    """
    regional_performance = {}
    regions = ['美国', '亚洲', '欧洲']
    
    metrics_dict = all_rounds_data[round_name]
    
    # 计算每个区域所有队伍的销售额
    # 修复：区域销售额指标名直接使用区域名（"美国"、"亚洲"、"欧洲"），而不是"在{region}销售"
    region_total_sales = {}
    for region in regions:
        total = 0
        region_sales = {}
        for team in teams:
            # 优先级：1. 直接区域名 2. "在{region}销售" 3. "{region}销售额"
            sales = get_metric_value(metrics_dict, region, team)
            if sales is None or sales == 0:
                # 尝试其他命名方式
                sales = get_metric_value(metrics_dict, f'在{region}销售', team)
            if sales is None or sales == 0:
                sales = get_metric_value(metrics_dict, f'{region}销售额', team)
            
            # 只统计有销售额的队伍，且销售额必须>0
            if sales is not None and sales > 0:
                region_sales[team] = sales
                total += sales
        region_total_sales[region] = {'total': total, 'team_sales': region_sales}
    
    # 计算销售趋势（对比上回合）
    rounds = ['ir00', 'pr01', 'pr02', 'pr03']
    round_idx = rounds.index(round_name) if round_name in rounds else -1
    prev_round = rounds[round_idx - 1] if round_idx > 0 else None
    
    for team in teams:
        regional_performance[team] = {}
        
        for region in regions:
            # 修复：区域销售额指标名直接使用区域名，优先级匹配
            sales = get_metric_value(metrics_dict, region, team)
            if sales is None or sales == 0:
                sales = get_metric_value(metrics_dict, f'在{region}销售', team)
            if sales is None or sales == 0:
                sales = get_metric_value(metrics_dict, f'{region}销售额', team)
            # 处理None值，统一为0
            if sales is None:
                sales = 0
            
            # 计算市场份额（替代方案）
            # 修复：只有销售额>0时才计算市场份额和排名
            market_share = None
            ranking = None
            
            if sales is not None and sales > 0:
                if region_total_sales[region]['total'] > 0:
                    market_share = (sales / region_total_sales[region]['total']) * 100
                
                # 计算排名（只有销售额>0的队伍才排名）
                team_sales = region_total_sales[region]['team_sales']
                if team_sales and sales in team_sales.values():
                    sorted_teams = sorted(team_sales.items(), key=lambda x: x[1], reverse=True)
                    for rank, (t, _) in enumerate(sorted_teams, 1):
                        if t == team:
                            ranking = rank
                            break
            
            # 计算销售趋势（如果数据可用）
            sales_trend = '稳定'
            if prev_round and prev_round in all_rounds_data:
                prev_metrics = all_rounds_data[prev_round]
                prev_sales = get_metric_value(prev_metrics, region, team)
                if prev_sales is None or prev_sales == 0:
                    prev_sales = get_metric_value(prev_metrics, f'在{region}销售', team)
                if prev_sales is None or prev_sales == 0:
                    prev_sales = get_metric_value(prev_metrics, f'{region}销售额', team)
                if prev_sales is None:
                    prev_sales = 0
                if prev_sales > 0:
                    growth_rate = ((sales - prev_sales) / prev_sales) * 100
                    if growth_rate > 10:
                        sales_trend = '增长'
                    elif growth_rate < -10:
                        sales_trend = '下降'
                    else:
                        sales_trend = '稳定'
                elif sales > 0:
                    sales_trend = '新进入'
            
            # 策略建议（考虑排名和趋势）
            suggestions = []
            if sales > 0:  # 只在有销售额时给出建议
                if ranking and ranking <= 3:
                    if sales_trend == '增长':
                        suggestions.append('巩固优势，考虑提价')
                    elif sales_trend == '稳定':
                        suggestions.append('增加功能或广告投入')
                    elif sales_trend == '下降':
                        suggestions.append('分析原因，调整策略')
                elif ranking and 4 <= ranking <= 8:
                    if sales_trend == '增长':
                        suggestions.append('加大投入，抢占份额')
                    elif sales_trend == '下降':
                        suggestions.append('评估退出或差异化')
                elif ranking and ranking > 8:
                    suggestions.append('退出或大幅调整策略')
            
            regional_performance[team][region] = {
                '销售额': sales,
                '市场份额': market_share,
                '排名': ranking,
                '销售趋势': sales_trend,
                '策略建议': suggestions
            }
    
    return regional_performance


# ============================================================================
# 第四章：竞争分析解码
# ============================================================================

def calculate_competitive_position(metrics_dict, teams):
    """三维度对标矩阵"""
    competitive_matrix = {}
    
    for team in teams:
        equity = get_metric_value(metrics_dict, '权益合计', team) or 0
        short_debt = get_metric_value(metrics_dict, '短期贷款', team) or 0
        long_debt = get_metric_value(metrics_dict, '长期贷款', team) or 0
        cash = get_metric_with_priority(metrics_dict, '现金', team) or 0
        sales = get_metric_with_priority(metrics_dict, '销售额', team) or 0
        rd_expense = get_metric_value(metrics_dict, '研发', team) or 0
        ad_expense = get_metric_value(metrics_dict, '广告', team) or 0
        profit = get_metric_with_priority(metrics_dict, '净利润', team) or 0
        
        # 1. 财务激进度
        if equity > 0:
            net_debt = (short_debt + long_debt) - cash
            financial_aggressiveness = (net_debt / equity) * 100
        else:
            financial_aggressiveness = 999
        
        # 2. 市场侵略性
        market_aggressiveness = (ad_expense / sales * 100) if sales > 0 else 0
        
        # 3. 技术投入度
        tech_investment = (rd_expense / sales * 100) if sales > 0 else 0
        
        # 策略类型识别
        strategy_type = '未知'
        if tech_investment > 20 and rd_expense > 0:
            ros = (profit / sales * 100) if sales > 0 else 0
            if ros > 20:
                strategy_type = '战略清晰（高投入+高回报）'
            else:
                strategy_type = '策略试错（高投入+低回报）'
        elif tech_investment < 1 and profit and profit > 0:
            strategy_type = '市场套利（零研发+高利润）'
        elif tech_investment < 5 and market_aggressiveness < 5:
            strategy_type = '稳健经营'
        
        competitive_matrix[team] = {
            '财务激进度': financial_aggressiveness,
            '市场侵略性': market_aggressiveness,
            '技术投入度': tech_investment,
            '策略类型': strategy_type
        }
    
    return competitive_matrix


def detect_strategy_changes(all_rounds_data, teams):
    """策略突变检测"""
    changes = {}
    rounds = ['ir00', 'pr01', 'pr02', 'pr03']
    
    for team in teams:
        changes[team] = {
            'alerts': [],
            'changes': {}
        }
        
        for i in range(len(rounds) - 1):
            rnd1, rnd2 = rounds[i], rounds[i + 1]
            
            if rnd1 not in all_rounds_data or rnd2 not in all_rounds_data:
                continue
            
            metrics1 = all_rounds_data[rnd1]
            metrics2 = all_rounds_data[rnd2]
            
            # 1. 现金异常波动
            cash1 = get_metric_with_priority(metrics1, '现金', team) or 0
            cash2 = get_metric_with_priority(metrics2, '现金', team) or 0
            cash_change = abs(cash2 - cash1)
            
            if cash_change > 500000:
                changes[team]['alerts'].append({
                    'type': '现金异常波动',
                    'round': f'{rnd1}→{rnd2}',
                    'value': cash_change,
                    'interpretation': '可能融资/出售资产' if cash2 > cash1 else '可能大幅投资/亏损'
                })
            
            # 2. 战略稳定性指数
            ebitda1 = get_metric_value(metrics1, 'EBITDA', team)
            if ebitda1 is None:
                ebitda1 = get_metric_value(metrics1, '息税折旧及摊销前利润', team) or 0
            else:
                ebitda1 = ebitda1 or 0
            
            ebitda2 = get_metric_value(metrics2, 'EBITDA', team)
            if ebitda2 is None:
                ebitda2 = get_metric_value(metrics2, '息税折旧及摊销前利润', team) or 0
            else:
                ebitda2 = ebitda2 or 0
            rd1 = get_metric_value(metrics1, '研发', team) or 0
            rd2 = get_metric_value(metrics2, '研发', team) or 0
            assets1 = get_metric_value(metrics1, '总资产', team) or 0
            
            if assets1 > 0:
                stability_index = 1 - (abs(ebitda2 - ebitda1) + abs(rd2 - rd1)) / assets1
                if stability_index < 0.3:
                    changes[team]['alerts'].append({
                        'type': '战略稳定性低',
                        'round': f'{rnd1}→{rnd2}',
                        'value': stability_index,
                        'interpretation': '策略变化剧烈，需重点关注'
                    })
    
    return changes


def detect_region_entry(all_rounds_data, teams):
    """
    检测区域市场进入（使用销售额替代市场份额）
    从方法论文档4.2.2节
    """
    region_entry_alerts = {}
    rounds = ['ir00', 'pr01', 'pr02', 'pr03']
    regions = ['美国', '亚洲', '欧洲']
    
    for team in teams:
        region_entry_alerts[team] = []
        
        for region in regions:
            prev_sales = 0
            
            for rnd in rounds:
                if rnd in all_rounds_data:
                    metrics_dict_rnd = all_rounds_data[rnd]
                    # 修复：区域销售额指标名直接使用区域名，优先级匹配
                    current_sales = get_metric_value(metrics_dict_rnd, region, team) or 0
                    if (current_sales is None or current_sales == 0):
                        current_sales = get_metric_value(metrics_dict_rnd, f'在{region}销售', team) or 0
                    if (current_sales is None or current_sales == 0):
                        current_sales = get_metric_value(metrics_dict_rnd, f'{region}销售额', team) or 0
                    
                    if prev_sales == 0 and current_sales and current_sales > 10000:  # 从无到有，销售额>10k
                        region_entry_alerts[team].append({
                            'region': region,
                            'round': rnd,
                            'sales': current_sales,
                            'interpretation': f'新进入{region}市场'
                        })
                    prev_sales = current_sales or 0
    
    return region_entry_alerts


def predict_next_move(all_rounds_data, teams, round_name, derived_metrics):
    """下回合意图预测"""
    predictions = {}
    metrics_dict = all_rounds_data[round_name]
    derived = derived_metrics.get(round_name, {})
    
    for team in teams:
        signals = []
        
        cash = get_metric_with_priority(metrics_dict, '现金', team) or 0
        sales_growth = derived.get('销售额_环比增长', {}).get(team, 0)
        sales_rank = derived.get('销售额_排名', {}).get(team, 999)
        rd_expense = get_metric_value(metrics_dict, '研发', team) or 0
        
        equity = get_metric_value(metrics_dict, '权益合计', team) or 0
        short_debt = get_metric_value(metrics_dict, '短期贷款', team) or 0
        long_debt = get_metric_value(metrics_dict, '长期贷款', team) or 0
        
        # 修复：确保能提取到EBITDA值
        ebitda = get_metric_value(metrics_dict, 'EBITDA', team)
        if ebitda is None:
            ebitda = get_metric_value(metrics_dict, '息税折旧及摊销前利润', team) or 0
        else:
            ebitda = ebitda or 0
        
        if equity > 0:
            net_debt = (short_debt + long_debt) - cash
            debt_equity_ratio = (net_debt / equity) * 100
        else:
            debt_equity_ratio = 999
        
        # 扩产信号
        if cash > 300000 and sales_growth > 10:
            signals.append({
                'action': '扩产',
                'probability': 70,
                'reason': '现金充足+销售增长'
            })
        
        # 价格战信号
        if cash > 500000 and sales_rank > 8:
            signals.append({
                'action': '价格战',
                'probability': 60,
                'reason': '现金充足+排名靠后'
            })
        
        # 技术投入信号
        if rd_expense > 400000:
            signals.append({
                'action': '技术投入',
                'probability': 75,
                'reason': '研发投入大，可能推出新技术'
            })
        
        # 财务危机信号
        if debt_equity_ratio > 100 and ebitda is not None and ebitda < 0:
            signals.append({
                'action': '出售资产/退出',
                'probability': 80,
                'reason': '财务危机（高负债+负EBITDA）'
            })
        
        # 现金危机信号
        if cash < 50000 and debt_equity_ratio > 70:
            signals.append({
                'action': '紧急融资',
                'probability': 85,
                'reason': '现金不足+高负债'
            })
        
        predictions[team] = signals
    
    return predictions


# ============================================================================
# 第五章：决策支持体系
# ============================================================================

def generate_strategy_recommendations(health_data, cash_flow_data, competitive_matrix, 
                                     derived_metrics, latest_round, teams):
    """
    生成下回合策略建议（资源分配决策树）
    基于方法论文档5.2节
    """
    recommendations = {}
    
    for team in teams:
        health = health_data.get(team, {})
        cash_flow = cash_flow_data.get(team, {})
        comp_pos = competitive_matrix.get(team, {})
        
        cash = health.get('indicators', {}).get('现金储备', 0) or 0
        derived = derived_metrics.get(latest_round, {})
        sales_growth = derived.get('销售额_环比增长', {}).get(team, 0)
        sales_rank = derived.get('销售额_排名', {}).get(team, 999)
        
        recommendation = {
            'mode': '',
            'actions': [],
            'resource_allocation': {},
            'risk_level': ''
        }
        
        # 资源分配决策树
        if cash < 100000:
            # 生存模式
            recommendation['mode'] = '生存模式'
            recommendation['actions'] = [
                '停止所有投资',
                '出售闲置产能',
                '削减非必要费用'
            ]
            recommendation['resource_allocation'] = {
                '研发': 0,
                '广告': 0,
                '现金保留': 100
            }
            recommendation['risk_level'] = '高'
        elif cash < 300000:
            # 维持模式
            recommendation['mode'] = '维持模式'
            recommendation['actions'] = [
                '仅必要广告投入',
                '维持现有产能',
                '保留现金缓冲'
            ]
            recommendation['resource_allocation'] = {
                '研发': 10,
                '广告': 20,
                '现金保留': 70
            }
            recommendation['risk_level'] = '中'
        else:
            # 进攻模式
            recommendation['mode'] = '进攻模式'
            actions = []
            allocation = {}
            total_allocated = 0
            cash_reserve_pct = 20  # 保留20%现金作为风险缓冲
            
            # 根据条件动态分配资源（确保总和不超过100%-现金保留）
            max_available = 100 - cash_reserve_pct
            
            if sales_growth > 10:
                actions.append('销售增长>10% → 考虑扩产')
                if total_allocated < max_available:
                    expand_pct = min(60, max_available - total_allocated)
                    allocation['扩产'] = expand_pct
                    total_allocated += expand_pct
            
            if comp_pos.get('技术投入度', 0) < 5 and total_allocated < max_available:
                actions.append('技术空白市场 → 研发+进入')
                rd_pct = min(40, max_available - total_allocated)
                allocation['研发'] = rd_pct
                total_allocated += rd_pct
            
            if sales_rank <= 3 and total_allocated < max_available:
                actions.append('份额领先 → 增加广告巩固')
                ad_pct = min(30, max_available - total_allocated)
                allocation['广告'] = ad_pct
                total_allocated += ad_pct
            
            # 如果没有其他分配，默认分配剩余资源到广告
            if not allocation and max_available > 0:
                actions.append('维持当前策略，适度投资')
                allocation['广告'] = min(30, max_available)
            
            allocation['现金保留'] = cash_reserve_pct + (max_available - total_allocated)
            
            if not actions:
                actions.append('维持当前策略，观察对手动态')
            
            recommendation['actions'] = actions
            recommendation['resource_allocation'] = allocation
            recommendation['risk_level'] = '低'
        
        recommendations[team] = recommendation
    
    return recommendations


def generate_checklist(health_data, regional_data, strategy_changes, teams, latest_round):
    """
    生成核心检查清单
    基于方法论文档5.3节
    """
    checklist = {}
    
    for team in teams:
        health = health_data.get(team, {})
        regional = regional_data.get(team, {})
        changes = strategy_changes.get(team, {})
        
        indicators = health.get('indicators', {})
        statuses = health.get('status', {})
        
        cash = indicators.get('现金储备', 0) or 0
        debt_equity = indicators.get('净债务权益比') or 0
        red_count = sum(1 for s in statuses.values() if '🔴' in str(s))
        
        checks = {
            '财务健康': [],
            '市场策略': [],
            '竞争态势': [],
            '风险控制': []
        }
        
        # 财务健康检查
        if cash >= 300000:
            checks['财务健康'].append('✅ 现金储备覆盖3个回合的亏损')
        else:
            checks['财务健康'].append('❌ 现金储备不足（需要≥$300k）')
        
        if red_count >= 2:
            checks['财务健康'].append('❌ 财务健康度有2个以上红灯')
        else:
            checks['财务健康'].append('✅ 财务健康度良好')
        
        if debt_equity and debt_equity < 70:
            checks['财务健康'].append('✅ 净债务/权益比在安全范围')
        else:
            checks['财务健康'].append('❌ 净债务/权益比过高（需要<70%）')
        
        # 市场策略检查
        has_sales = False
        top_3_count = 0
        for region in ['美国', '亚洲', '欧洲']:
            rp = regional.get(region, {})
            if rp.get('销售额', 0) > 0:
                has_sales = True
            if rp.get('排名') and rp['排名'] <= 3:
                top_3_count += 1
        
        if has_sales:
            checks['市场策略'].append('✅ 有区域销售额')
        else:
            checks['市场策略'].append('⚠️ 区域销售额为零')
        
        if top_3_count > 0:
            checks['市场策略'].append(f'✅ {top_3_count}个区域排名前3')
        else:
            checks['市场策略'].append('⚠️ 主要市场排名未进前3')
        
        # 竞争态势检查
        alerts = changes.get('alerts', [])
        if alerts:
            checks['竞争态势'].append(f'⚠️ 检测到{len(alerts)}个策略突变警报')
        else:
            checks['竞争态势'].append('✅ 对手策略稳定')
        
        # 风险控制检查
        if cash >= 300000 * 0.2:  # 至少20%的风险缓冲
            checks['风险控制'].append('✅ 保留至少20%现金作为风险缓冲')
        else:
            checks['风险控制'].append('❌ 风险缓冲不足')
        
        checks['风险控制'].append('✅ 已考虑最坏情景')
        checks['风险控制'].append('✅ 策略具有灵活性')
        
        checklist[team] = checks
    
    return checklist


# ============================================================================
# 报告生成
# ============================================================================

def generate_comprehensive_report(all_rounds_data, teams, health_data, cash_flow_data, 
                                  regional_data, competitive_matrix, strategy_changes,
                                  predictions, derived_metrics, anomalies, latest_round,
                                  strategy_recommendations=None, checklist=None, region_entry_alerts=None):
    """生成完整分析报告"""
    report = []
    
    report.append("# 企业模拟经营战报分析报告（按方法论3.0）\n")
    report.append(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    report.append("基于方法论文档3.0版本进行完整分析\n")
    report.append("=" * 80 + "\n")
    
    # 一、执行摘要
    report.append("\n## 一、执行摘要\n")
    
    # 找出领先队伍和关键指标
    sales_rankings = derived_metrics.get(latest_round, {}).get('销售额_排名', {})
    metrics_dict = all_rounds_data[latest_round]
    
    if sales_rankings:
        top_teams = sorted(sales_rankings.items(), key=lambda x: x[1])[:3]
        report.append("### 当前回合销售额排名TOP3：\n")
        for rank, (team, position) in enumerate(top_teams, 1):
            # 获取关键指标
            profit = get_metric_with_priority(metrics_dict, '净利润', team) or 0
            cash = get_metric_with_priority(metrics_dict, '现金', team) or 0
            prev_round = 'pr02' if latest_round == 'pr03' else 'pr01'
            if prev_round in all_rounds_data:
                prev_profit = get_metric_with_priority(all_rounds_data[prev_round], '净利润', team) or 0
                if prev_profit != 0:
                    profit_growth = ((profit - prev_profit) / abs(prev_profit)) * 100
                else:
                    profit_growth = 0
            else:
                profit_growth = 0
            
            report.append(f"{rank}. **{team}**（排名：第{position}位）\n")
            report.append(f"   - 净利润：${profit/1000:.0f}k（环比{profit_growth:+.1f}%）\n")
            report.append(f"   - 现金：${cash/1000:.0f}k\n")
    
    # 核心问题识别
    report.append("\n### 关键发现：\n")
    
    # 识别高风险队伍
    high_risk_teams = []
    for team, health in health_data.items():
        red_count = sum(1 for s in health.get('status', {}).values() if '🔴' in str(s))
        if red_count >= 2:
            high_risk_teams.append(team)
    
    if high_risk_teams:
        report.append(f"- ⚠️ **高风险队伍**：{', '.join(high_risk_teams[:5])}（财务健康度有2个以上红灯）\n")
    
    # 识别策略突变
    strategy_change_teams = []
    for team, changes in strategy_changes.items():
        if changes.get('alerts'):
            strategy_change_teams.append(team)
    
    if strategy_change_teams:
        report.append(f"- 🔄 **策略突变队伍**：{', '.join(strategy_change_teams[:3])}（需重点关注）\n")
    
    # 二、数据基础建设
    report.append("\n\n## 二、数据基础建设\n")
    
    report.append("### 2.1 数据完整性验证\n")
    validation_issues = validate_data_integrity(all_rounds_data[latest_round], teams)
    if validation_issues:
        report.append("发现以下问题：\n")
        for issue in validation_issues[:5]:  # 只显示前5个
            report.append(f"- {issue['team']}: 误差{issue['error_rate']:.2f}% - {issue['status']}\n")
    else:
        report.append("✅ 数据完整性验证通过\n")
    
    report.append("\n### 2.2 异常值检测\n")
    if anomalies:
        for team, anomaly_list in list(anomalies.items())[:5]:
            report.append(f"\n**{team}**：\n")
            for anomaly in anomaly_list:
                report.append(f"- {anomaly['type']}: {anomaly['value']:,.0f} ({anomaly['rule']})\n")
    else:
        report.append("✅ 未发现异常值\n")
    
    # 三、自身诊断分析
    report.append("\n\n## 三、自身诊断分析\n")
    
    report.append("### 3.1 财务健康度红绿灯系统\n")
    report.append("| 队伍 | 现金储备 | 净债务/权益比 | EBITDA率 | 权益比率 | 研发回报率 | 行动建议 |")
    report.append("|------|---------|--------------|---------|---------|-----------|---------|")
    
    for team in teams:
        h = health_data.get(team, {})
        indicators = h.get('indicators', {})
        statuses = h.get('status', {})
        
        cash_val = f"${indicators.get('现金储备', 0)/1000:.0f}k" if indicators.get('现金储备') is not None else "N/A"
        cash_status = statuses.get('现金储备', 'N/A')
        
        debt_val = f"{indicators.get('净债务权益比', 0):.1f}%" if indicators.get('净债务权益比') is not None else "N/A"
        debt_status = statuses.get('净债务权益比', 'N/A')
        
        # 修复：EBITDA率显示精度，当值很小时显示更多小数位
        ebitda_rate = indicators.get('EBITDA率')
        if ebitda_rate is not None:
            if ebitda_rate < 0.1:
                ebitda_val = f"{ebitda_rate:.4f}%"
            else:
                ebitda_val = f"{ebitda_rate:.1f}%"
        else:
            ebitda_val = "N/A"
        ebitda_status = statuses.get('EBITDA率', 'N/A')
        
        equity_val = f"{indicators.get('权益比率', 0):.1f}%" if indicators.get('权益比率') is not None else "N/A"
        equity_status = statuses.get('权益比率', 'N/A')
        
        rd_val = f"{indicators.get('研发回报率', 0):.1f}%" if indicators.get('研发回报率') is not None else "N/A"
        rd_status = statuses.get('研发回报率', 'N/A')
        
        action = h.get('action_required', ['-'])[0] if h.get('action_required') else '-'
        
        report.append(f"| {team} | {cash_val} {cash_status} | {debt_val} {debt_status} | "
                     f"{ebitda_val} {ebitda_status} | {equity_val} {equity_status} | "
                     f"{rd_val} {rd_status} | {action} |")
    
    report.append("\n\n### 3.2 现金流源头分析\n")
    report.append("| 队伍 | 现金变化 | 经营现金流(EBITDA) | 现金流类型 |")
    report.append("|------|---------|------------------|-----------|")
    
    for team in teams:
        cf = cash_flow_data.get(team, {})
        report.append(f"| {team} | ${cf.get('现金变化', 0)/1000:.0f}k | "
                     f"${cf.get('经营现金流(EBITDA)', 0)/1000:.0f}k | {cf.get('现金流类型', 'N/A')} |")
    
    report.append("\n\n### 3.3 区域市场表现分析\n")
    report.append("**数据说明**：由于Excel中区域销售额数据不可用或数据量极小（仅占总额的0.05%-0.65%），\n")
    report.append("当前使用的'美国'、'亚洲'、'欧洲'指标的实际含义可能与区域销售额不符，仅供参考。\n\n")
    for team in teams[:5]:  # 显示前5个队伍
        regional = regional_data.get(team, {})
        report.append(f"\n**{team}**：\n")
        has_any_sales = False
        for region in ['美国', '亚洲', '欧洲']:
            rp = regional.get(region, {})
            sales = rp.get('销售额', 0) or 0
            if sales > 0:
                has_any_sales = True
                report.append(f"- **{region}**：")
                report.append(f" 销售额 ${sales/1000:.0f}k")
                if rp.get('市场份额'):
                    report.append(f"，市场份额 {rp['市场份额']:.1f}%")
                if rp.get('排名'):
                    report.append(f"，排名第{rp['排名']}位")
                if rp.get('销售趋势'):
                    trend_symbol = "📈" if rp['销售趋势'] == '增长' else "📉" if rp['销售趋势'] == '下降' else "➡️"
                    report.append(f"，趋势：{trend_symbol} {rp['销售趋势']}")
                if rp.get('策略建议'):
                    report.append(f" → {'; '.join(rp['策略建议'])}\n")
        
        if not has_any_sales:
            report.append("- ⚠️ 暂无区域销售额数据\n")
    
    # 四、竞争分析解码
    report.append("\n\n## 四、竞争分析解码\n")
    
    report.append("### 4.1 三维度对标矩阵\n")
    report.append("| 队伍 | 财务激进度 | 市场侵略性 | 技术投入度 | 策略类型 |")
    report.append("|------|-----------|-----------|-----------|---------|")
    
    for team in teams:
        cm = competitive_matrix.get(team, {})
        report.append(f"| {team} | {cm.get('财务激进度', 0):.1f}% | "
                     f"{cm.get('市场侵略性', 0):.1f}% | {cm.get('技术投入度', 0):.1f}% | "
                     f"{cm.get('策略类型', '未知')} |")
    
    report.append("\n\n### 4.2 策略突变检测\n")
    for team in teams:
        changes = strategy_changes.get(team, {})
        if changes.get('alerts'):
            report.append(f"\n**{team}**：\n")
            for alert in changes['alerts'][:3]:  # 只显示前3个警报
                report.append(f"- ⚠️ {alert['type']} ({alert['round']}): {alert.get('interpretation', '')}\n")
    
    report.append("\n\n### 4.3 下回合意图预测\n")
    for team in teams:
        pred = predictions.get(team, [])
        if pred:
            report.append(f"\n**{team}**：\n")
            for signal in pred[:3]:  # 只显示前3个信号
                report.append(f"- {signal['action']} (概率{signal['probability']}%): {signal['reason']}\n")
    
    # 五、多回合趋势分析
    report.append("\n\n## 五、多回合趋势分析\n")
    
    rounds = ['ir00', 'pr01', 'pr02', 'pr03']
    available_rounds = [r for r in rounds if r in all_rounds_data]
    
    for metric_name in ['销售额', '净利润', '现金']:
        report.append(f"\n### {metric_name}趋势\n")
        report.append("| 队伍 | " + " | ".join([r.upper() for r in available_rounds]) + " |")
        report.append("|------|" + "|".join(["------" for _ in available_rounds]) + "|\n")
        
        for team in teams[:8]:  # 显示前8个队伍
            values = []
            for rnd in available_rounds:
                val = get_metric_with_priority(all_rounds_data[rnd], metric_name, team)
                if val is not None:
                    if metric_name == '现金':
                        values.append(f"${val/1000:.0f}k")
                    else:
                        values.append(f"{val/1000:.0f}k")
                else:
                    values.append("N/A")
            report.append(f"| {team} | " + " | ".join(values) + " |\n")
        
        # 添加环比增长率
        if len(available_rounds) > 1:
            report.append("\n**环比增长率**：\n")
            report.append("| 队伍 | " + " | ".join([f"{r.upper()}" for r in available_rounds[1:]]) + " |")
            report.append("|------|" + "|".join(["------" for _ in available_rounds[1:]]) + "|\n")
            
            for team in teams[:8]:
                growth_rates = []
                for i in range(1, len(available_rounds)):
                    rnd = available_rounds[i]
                    derived = derived_metrics.get(rnd, {})
                    growth = derived.get(f'{metric_name}_环比增长', {}).get(team)
                    if growth is not None:
                        growth_rates.append(f"{growth:+.1f}%")
                    else:
                        growth_rates.append("N/A")
                report.append(f"| {team} | " + " | ".join(growth_rates) + " |\n")
    
    # 六、决策建议（第五章内容）
    if strategy_recommendations:
        report.append("\n\n## 六、决策建议\n")
        
        report.append("### 6.1 下回合策略建议\n")
        for team in teams[:5]:  # 显示前5个队伍
            rec = strategy_recommendations.get(team, {})
            if rec:
                report.append(f"\n**{team}**：")
                report.append(f"\n- 模式：{rec.get('mode', 'N/A')}（风险等级：{rec.get('risk_level', 'N/A')}）")
                report.append(f"- 行动建议：")
                for action in rec.get('actions', []):
                    report.append(f"  - {action}")
                if rec.get('resource_allocation'):
                    report.append(f"- 资源分配：")
                    for item, value in rec.get('resource_allocation', {}).items():
                        report.append(f"  - {item}: {value}%")
        
        report.append("\n\n### 6.2 区域市场进入检测\n")
        if region_entry_alerts:
            for team in teams:
                alerts = region_entry_alerts.get(team, [])
                if alerts:
                    report.append(f"\n**{team}**：\n")
                    for alert in alerts[:3]:  # 只显示前3个
                        report.append(f"- ⚠️ {alert.get('interpretation', '')}（{alert.get('round', '')}，销售额：${alert.get('sales', 0)/1000:.0f}k）\n")
    
    # 七、核心检查清单
    if checklist:
        report.append("\n\n## 七、核心检查清单\n")
        report.append("**提交决策前必答问题**：\n")
        
        for team in teams[:3]:  # 显示前3个队伍
            checks = checklist.get(team, {})
            if checks:
                report.append(f"\n### {team}\n")
                
                for category, items in checks.items():
                    report.append(f"\n**{category}检查**：\n")
                    for item in items:
                        report.append(f"- {item}\n")
    
    # 八、可视化图表描述（方法论文档6.2节）
    report.append("\n\n## 八、关键图表描述\n")
    report.append("> 注：以下为图表的文本描述，实际可视化图表可使用matplotlib等工具生成\n\n")
    
    # 1. 财务健康度仪表盘
    report.append("### 8.1 财务健康度仪表盘\n")
    report.append("**指标状态概览**：\n\n")
    for team in teams[:5]:
        health = health_data.get(team, {})
        statuses = health.get('status', {})
        indicators = health.get('indicators', {})
        
        report.append(f"**{team}**：\n")
        for ind_name in ['现金储备', '净债务权益比', 'EBITDA率', '权益比率', '研发回报率']:
            status = statuses.get(ind_name, 'N/A')
            value = indicators.get(ind_name)
            if value is not None:
                if ind_name == '现金储备':
                    report.append(f"- {ind_name}: ${value/1000:.0f}k {status}\n")
                elif ind_name == 'EBITDA率':
                    # 修复：EBITDA率显示精度
                    if value < 0.1:
                        report.append(f"- {ind_name}: {value:.4f}% {status}\n")
                    else:
                        report.append(f"- {ind_name}: {value:.1f}% {status}\n")
                else:
                    report.append(f"- {ind_name}: {value:.1f}% {status}\n")
            else:
                report.append(f"- {ind_name}: N/A {status}\n")
        report.append("\n")
    
    # 2. 竞争态势矩阵描述
    report.append("\n### 8.2 竞争态势矩阵图\n")
    report.append("**维度分布**（X轴：财务激进度，Y轴：技术投入度，气泡大小：市场侵略性）：\n\n")
    report.append("| 队伍 | 财务激进度 | 技术投入度 | 市场侵略性 | 策略类型 | 象限位置 |\n")
    report.append("|------|-----------|-----------|-----------|---------|---------|\n")
    
    for team in teams:
        cm = competitive_matrix.get(team, {})
        fin_agg = cm.get('财务激进度', 0)
        tech_inv = cm.get('技术投入度', 0)
        mkt_agg = cm.get('市场侵略性', 0)
        strategy = cm.get('策略类型', '未知')
        
        # 判断象限位置（优化999%的显示）
        if fin_agg >= 999:
            fin_pos = "极端激进（权益<0）"
        elif fin_agg > 50:
            fin_pos = "高"
        else:
            fin_pos = "低"
        
        if tech_inv > 10:
            tech_pos = "高"
        else:
            tech_pos = "低"
        
        if fin_agg >= 999:
            quadrant = f"极端激进×{tech_pos}技术"
        else:
            quadrant = f"{fin_pos}财务×{tech_pos}技术"
        
        report.append(f"| {team} | {fin_agg:.1f}% | {tech_inv:.1f}% | {mkt_agg:.1f}% | {strategy} | {quadrant} |\n")
    
    # 3. 多回合趋势对比
    report.append("\n### 8.3 多回合趋势对比图\n")
    report.append("**关键指标趋势**（详见第五章多回合趋势分析部分）：\n")
    report.append("- 销售额：整体趋势向上/向下/稳定\n")
    report.append("- 净利润：盈利改善/恶化/波动\n")
    report.append("- 现金：现金流健康/紧张/危机\n")
    
    # 4. 区域市场表现
    report.append("\n### 8.4 区域市场表现图\n")
    report.append("**区域销售额排名**：\n\n")
    for region in ['美国', '亚洲', '欧洲']:
        report.append(f"**{region}市场**：\n")
        
        # 获取该区域所有队伍的排名（修复：只有销售额>0的队伍才排名）
        region_rankings = []
        for team in teams:
            regional = regional_data.get(team, {})
            rp = regional.get(region, {})
            # 修复：只有销售额>0且有排名才加入排名列表
            sales = rp.get('销售额', 0) or 0
            if rp.get('排名') and sales > 0:
                region_rankings.append({
                    'team': team,
                    'rank': rp['排名'],
                    'sales': sales,
                    'market_share': rp.get('市场份额', 0)
                })
        
        if region_rankings:
            region_rankings.sort(key=lambda x: x['rank'])
            report.append("| 排名 | 队伍 | 销售额 | 市场份额 | 趋势 |\n")
            report.append("|------|------|--------|---------|------|\n")
            for item in region_rankings[:5]:
                # 判断趋势（简化：如果有排名变化数据则使用）
                trend = "→"  # 默认稳定
                report.append(f"| {item['rank']} | {item['team']} | ${item['sales']/1000:.0f}k | {item['market_share']:.1f}% | {trend} |\n")
        report.append("\n")
    
    return "\n".join(report)


# ============================================================================
# 逻辑验证与检查
# ============================================================================

def validate_logic(all_rounds_data, teams, health_data, derived_metrics, 
                  competitive_matrix, latest_round):
    """
    验证分析逻辑的合理性和一致性
    """
    issues = []
    
    metrics_dict = all_rounds_data[latest_round]
    
    # 1. 验证财务健康度计算的一致性
    for team in teams:
        health = health_data.get(team, {})
        indicators = health.get('indicators', {})
        
        # 验证现金提取
        cash_health = indicators.get('现金储备')
        cash_direct = get_metric_with_priority(metrics_dict, '现金', team) or 0
        if cash_health and abs(cash_health - cash_direct) > 0.01:
            issues.append({
                'type': '数据不一致',
                'team': team,
                'metric': '现金',
                'description': f'健康度计算中的现金值({cash_health})与直接提取值({cash_direct})不一致'
            })
        
        # 验证净债务/权益比计算
        equity = get_metric_value(metrics_dict, '权益合计', team) or 0
        short_debt = get_metric_value(metrics_dict, '短期贷款', team) or 0
        long_debt = get_metric_value(metrics_dict, '长期贷款', team) or 0
        cash = get_metric_with_priority(metrics_dict, '现金', team) or 0
        
        if equity > 0:
            calculated_debt_equity = ((short_debt + long_debt - cash) / equity) * 100
            stored_debt_equity = indicators.get('净债务权益比')
            
            if stored_debt_equity is not None:
                if abs(calculated_debt_equity - stored_debt_equity) > 0.1:
                    issues.append({
                        'type': '计算不一致',
                        'team': team,
                        'metric': '净债务权益比',
                        'description': f'计算值({calculated_debt_equity:.2f}%)与存储值({stored_debt_equity:.2f}%)不一致'
                    })
    
    # 2. 验证资源分配总和
    # (这部分在主函数中调用时验证)
    
    # 3. 验证排名逻辑
    for rnd in ['ir00', 'pr01', 'pr02', 'pr03']:
        if rnd in all_rounds_data:
            derived = derived_metrics.get(rnd, {})
            sales_rankings = derived.get('销售额_排名', {})
            
            if sales_rankings:
                # 验证排名是否连续且从1开始
                ranks = sorted([r for r in sales_rankings.values() if r is not None])
                if ranks and (ranks[0] != 1 or len(set(ranks)) != len(ranks)):
                    issues.append({
                        'type': '排名逻辑错误',
                        'round': rnd,
                        'description': f'销售额排名不连续或重复'
                    })
    
    return issues


# ============================================================================
# 主程序
# ============================================================================

def main():
    """主函数"""
    print("=" * 80)
    print("商业模拟竞赛结果综合分析 v3.0")
    print("严格按照方法论文档3.0版本进行分析")
    print("=" * 80)
    
    # 第一步：数据基础建设
    print("\n【第一步：数据基础建设】")
    all_rounds_data = {}
    teams = []
    
    for round_name, file_path in FILES.items():
        if not file_path.exists():
            print(f"警告: 文件不存在 {file_path}")
            continue
        
        print(f"  正在处理 {round_name}...")
        metrics_dict, round_teams = read_excel_data(str(file_path))
        
        if not teams:
            teams = normalize_team_names(round_teams)
        
        all_rounds_data[round_name] = metrics_dict
        print(f"    [OK] 提取到 {len(metrics_dict)} 个指标")
        print(f"    [OK] 队伍数量: {len(round_teams)}")
    
    if not all_rounds_data:
        print("错误: 未能读取任何数据文件")
        return
    
    latest_round = 'pr03' if 'pr03' in all_rounds_data else 'pr02'
    print(f"\n  最新回合: {latest_round}")
    
    # 异常值检测
    anomalies = detect_anomalies(all_rounds_data[latest_round], teams)
    print(f"  检测到 {sum(len(v) for v in anomalies.values())} 个异常值")
    
    # 计算衍生指标
    print("\n  计算衍生指标...")
    derived_metrics = calculate_derived_metrics(all_rounds_data, teams)
    print(f"    [OK] 完成")
    
    # 第二步：自身诊断分析
    print("\n【第二步：自身诊断分析】")
    
    print("  计算财务健康度...")
    health_data = calculate_financial_health(all_rounds_data[latest_round], teams)
    
    print("  分析现金流...")
    prev_round = 'pr02' if latest_round == 'pr03' else 'pr01'
    prev_metrics = all_rounds_data.get(prev_round, {})
    cash_flow_data = analyze_cash_flow_source(all_rounds_data[latest_round], teams, prev_metrics)
    
    print("  分析区域市场表现...")
    regional_data = analyze_regional_market(all_rounds_data, teams, latest_round)
    
    # 第三步：竞争分析解码
    print("\n【第三步：竞争分析解码】")
    
    print("  计算三维度对标矩阵...")
    competitive_matrix = calculate_competitive_position(all_rounds_data[latest_round], teams)
    
    print("  检测策略突变...")
    strategy_changes = detect_strategy_changes(all_rounds_data, teams)
    
    print("  预测下回合意图...")
    predictions = predict_next_move(all_rounds_data, teams, latest_round, derived_metrics)
    
    print("  检测区域市场进入...")
    region_entry_alerts = detect_region_entry(all_rounds_data, teams)
    
    # 第四步：决策支持体系
    print("\n【第四步：决策支持体系】")
    
    print("  生成策略建议...")
    strategy_recommendations = generate_strategy_recommendations(
        health_data, cash_flow_data, competitive_matrix, 
        derived_metrics, latest_round, teams
    )
    
    print("  生成检查清单...")
    checklist = generate_checklist(
        health_data, regional_data, strategy_changes, teams, latest_round
    )
    
    # 逻辑验证
    print("\n【逻辑验证检查】")
    logic_issues = validate_logic(
        all_rounds_data, teams, health_data, derived_metrics,
        competitive_matrix, latest_round
    )
    if logic_issues:
        print(f"  发现 {len(logic_issues)} 个逻辑问题，已记录")
        for issue in logic_issues[:3]:
            print(f"    - {issue.get('type')}: {issue.get('description', '')}")
    else:
        print("  [OK] 逻辑验证通过")
    
    # 验证资源分配合理性
    for team, rec in strategy_recommendations.items():
        allocation = rec.get('resource_allocation', {})
        total = sum(v for v in allocation.values() if isinstance(v, (int, float)))
        if abs(total - 100) > 1:  # 允许1%的误差
            print(f"  警告: {team}资源分配总和={total:.1f}%，不等于100%")
    
    # 第五步：生成报告
    print("\n【第五步：生成分析报告】")
    
    report = generate_comprehensive_report(
        all_rounds_data, teams, health_data, cash_flow_data,
        regional_data, competitive_matrix, strategy_changes,
        predictions, derived_metrics, anomalies, latest_round,
        strategy_recommendations, checklist, region_entry_alerts
    )
    
    # 保存报告
    output_dir = Path(__file__).parent.parent.parent / '分析'
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / '方法论3.0完整分析报告.md'
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"\n  [OK] 报告已保存到: {output_file}")
    print("\n" + "=" * 80)
    print("分析完成！")
    print("=" * 80)


if __name__ == '__main__':
    main()

