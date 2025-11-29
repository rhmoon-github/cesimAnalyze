#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
数据分析和Excel处理工具模块
整合所有数据读取、检查、诊断功能
"""

import pandas as pd
import numpy as np
import os
from pathlib import Path

def read_excel_data(file_path, team_row_idx=4, data_start_row=5):
    """
    读取Excel文件并解析数据结构
    
    Args:
        file_path: Excel文件路径
        team_row_idx: 队伍名称所在行索引（默认4，即第5行）
        data_start_row: 数据开始行索引（默认5，即第6行）
    
    Returns:
        metrics_dict: 指标字典，格式为 {指标名: {队伍名: 数值}}
        teams: 队伍列表
    """
    df = pd.read_excel(file_path, sheet_name='Results', header=None)
    
    # 获取队伍名称
    team_row = df.iloc[team_row_idx]
    teams = [str(t).strip() for t in team_row[1:] if pd.notna(t) and str(t).strip() != '']
    
    # 从数据开始行读取数据
    data_df = df.iloc[data_start_row:].copy()
    data_df.columns = ['指标'] + teams + ['Unnamed'] * (len(data_df.columns) - len(teams) - 1)
    
    # 构建指标字典
    metrics_dict = {}
    
    for idx, row in data_df.iterrows():
        indicator = str(row['指标']).strip() if pd.notna(row['指标']) else ''
        
        if indicator == '' or indicator == 'nan':
            continue
        
        # 提取各队伍的数据
        team_data = {}
        for team in teams:
            val = row[team]
            if pd.notna(val):
                try:
                    # 转换为数值
                    if isinstance(val, (int, float)):
                        team_data[team] = float(val)
                    elif isinstance(val, str):
                        cleaned = val.replace(',', '').replace('$', '').replace('%', '').replace(' ', '').strip()
                        if cleaned:
                            team_data[team] = float(cleaned)
                        else:
                            team_data[team] = None
                    else:
                        team_data[team] = None
                except:
                    team_data[team] = None
            else:
                team_data[team] = None
        
        if any(v is not None for v in team_data.values()):
            metrics_dict[indicator] = team_data
    
    return metrics_dict, teams


def find_metric(metrics_dict, keywords, exact_match=False):
    """
    根据关键词查找指标
    
    Args:
        metrics_dict: 指标字典
        keywords: 关键词列表或单个关键词
        exact_match: 是否精确匹配
    
    Returns:
        找到的指标数据字典，如果未找到返回空字典
    """
    if isinstance(keywords, str):
        keywords = [keywords]
    
    for key in metrics_dict.keys():
        for keyword in keywords:
            if exact_match:
                if keyword == str(key).strip():
                    return metrics_dict[key]
            else:
                if keyword in str(key):
                    return metrics_dict[key]
    return {}


def list_all_metrics(file_path, max_count=200):
    """
    列出Excel文件中的所有指标名称
    
    Args:
        file_path: Excel文件路径
        max_count: 最多返回的指标数量
    
    Returns:
        指标列表
    """
    metrics_dict, teams = read_excel_data(file_path)
    return list(metrics_dict.keys())[:max_count]


def check_excel_structure(file_path):
    """
    检查Excel文件结构
    
    Args:
        file_path: Excel文件路径
    
    Returns:
        结构信息字典
    """
    excel_file = pd.ExcelFile(file_path)
    df = pd.read_excel(file_path, sheet_name='Results', header=None)
    
    # 获取队伍信息
    team_row = df.iloc[4]
    teams = [str(t).strip() for t in team_row[1:] if pd.notna(t) and str(t).strip() != '']
    
    # 获取所有指标
    metrics_dict, _ = read_excel_data(file_path)
    
    # 查找区域相关指标
    regions = ['美国', '亚洲', '欧洲', 'America', 'Asia', 'Europe']
    region_metrics = {}
    for region in regions:
        region_metrics[region] = [k for k in metrics_dict.keys() if region in str(k)]
    
    # 查找市场相关指标
    market_keywords = ['市场', '份额', '占有率']
    market_metrics = [k for k in metrics_dict.keys() if any(kw in str(k) for kw in market_keywords)]
    
    # 查找需求相关指标
    demand_keywords = ['需求', '未满足']
    demand_metrics = [k for k in metrics_dict.keys() if any(kw in str(k) for kw in demand_keywords)]
    
    # 查找产能相关指标
    capacity_keywords = ['产能', '利用率', '产量']
    capacity_metrics = [k for k in metrics_dict.keys() if any(kw in str(k) for kw in capacity_keywords)]
    
    return {
        'sheet_names': excel_file.sheet_names,
        'shape': df.shape,
        'teams': teams,
        'total_metrics': len(metrics_dict),
        'region_metrics': region_metrics,
        'market_metrics': market_metrics,
        'demand_metrics': demand_metrics,
        'capacity_metrics': capacity_metrics
    }


def diagnose_missing_data(file_path, target_metrics=None, target_team=None):
    """
    诊断缺失数据
    
    Args:
        file_path: Excel文件路径
        target_metrics: 目标指标列表，如果为None则使用默认列表
        target_team: 目标队伍名称
    
    Returns:
        诊断结果字典
    """
    metrics_dict, teams = read_excel_data(file_path)
    
    if target_metrics is None:
        target_metrics = [
            '销售额', '净利润', '现金', '权益', 'EBITDA',
            '美国市场份额', '亚洲市场份额', '欧洲市场份额',
            '美国未满足需求', '亚洲未满足需求', '欧洲未满足需求',
            '美国产能利用率', '亚洲产能利用率', '欧洲产能利用率'
        ]
    
    if target_team is None:
        target_team = teams[0] if teams else None
    
    diagnosis = {
        'found_metrics': {},
        'missing_metrics': [],
        'similar_metrics': {}
    }
    
    for metric in target_metrics:
        # 尝试精确匹配
        exact_match = find_metric(metrics_dict, [metric], exact_match=True)
        if exact_match:
            diagnosis['found_metrics'][metric] = exact_match.get(target_team)
            continue
        
        # 尝试部分匹配
        partial_match = find_metric(metrics_dict, [metric], exact_match=False)
        if partial_match:
            # 找到类似的指标
            similar_keys = [k for k in metrics_dict.keys() if metric in str(k)]
            if similar_keys:
                diagnosis['similar_metrics'][metric] = similar_keys
                diagnosis['found_metrics'][metric] = partial_match.get(target_team)
                continue
        
        diagnosis['missing_metrics'].append(metric)
    
    return diagnosis


def get_metric_value(metrics_dict, metric_name, team_name):
    """
    获取特定队伍和指标的数值（支持优先级列表）
    
    Args:
        metrics_dict: 指标字典
        metric_name: 指标名称（可以是字符串或列表，列表表示按优先级匹配）
        team_name: 队伍名称
    
    Returns:
        指标值，如果未找到返回None
    """
    # 如果metric_name是列表，按优先级顺序尝试匹配
    if isinstance(metric_name, list):
        for name in metric_name:
            metric_data = find_metric(metrics_dict, [name])
            if metric_data and team_name in metric_data:
                val = metric_data.get(team_name)
                if val is not None:  # 只返回非None值
                    return val
        return None
    else:
        # 单个字符串，直接查找
        metric_data = find_metric(metrics_dict, [metric_name])
        if metric_data:
            return metric_data.get(team_name)
        return None


def print_structure_info(structure_info):
    """打印Excel结构信息"""
    print("=" * 80)
    print("Excel文件结构分析")
    print("=" * 80)
    print(f"\n工作表: {structure_info['sheet_names']}")
    print(f"数据形状: {structure_info['shape']}")
    print(f"队伍数量: {len(structure_info['teams'])}")
    print(f"队伍列表: {', '.join(structure_info['teams'])}")
    print(f"指标总数: {structure_info['total_metrics']}")
    
    print(f"\n区域相关指标:")
    for region, metrics in structure_info['region_metrics'].items():
        if metrics:
            print(f"  {region}: {len(metrics)} 个")
            for m in metrics[:5]:
                print(f"    - {m}")
    
    print(f"\n市场相关指标: {len(structure_info['market_metrics'])} 个")
    for m in structure_info['market_metrics'][:10]:
        print(f"  - {m}")
    
    print(f"\n需求相关指标: {len(structure_info['demand_metrics'])} 个")
    for m in structure_info['demand_metrics'][:10]:
        print(f"  - {m}")
    
    print(f"\n产能相关指标: {len(structure_info['capacity_metrics'])} 个")
    for m in structure_info['capacity_metrics'][:10]:
        print(f"  - {m}")


def print_diagnosis(diagnosis):
    """打印诊断结果"""
    print("=" * 80)
    print("数据诊断结果")
    print("=" * 80)
    
    if diagnosis['found_metrics']:
        print(f"\n✓ 找到的指标 ({len(diagnosis['found_metrics'])} 个):")
        for metric, value in diagnosis['found_metrics'].items():
            print(f"  {metric}: {value}")
    
    if diagnosis['similar_metrics']:
        print(f"\n~ 找到类似指标 ({len(diagnosis['similar_metrics'])} 个):")
        for metric, similar in diagnosis['similar_metrics'].items():
            print(f"  {metric}:")
            for sim in similar[:3]:
                print(f"    - {sim}")
    
    if diagnosis['missing_metrics']:
        print(f"\n✗ 缺失的指标 ({len(diagnosis['missing_metrics'])} 个):")
        for metric in diagnosis['missing_metrics']:
            print(f"  - {metric}")

