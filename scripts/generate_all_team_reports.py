#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
批量生成所有队伍的详细分析报告
"""

import sys
from pathlib import Path

# 添加utils目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent / 'utils'))
sys.path.insert(0, str(Path(__file__).parent))

from utils_data_analysis import read_excel_data
from analyze_team_detail import analyze_team_detailed

def main(input_dir, output_dir):
    """为所有队伍生成详细分析报告"""
    
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 读取数据文件获取所有队伍名称
    ir00_path = input_dir / 'results-ir00.xls'
    if not ir00_path.exists():
        # 尝试读取 r01 或 pr01
        r01_path = input_dir / 'results-r01.xls'
        if not r01_path.exists():
            r01_path = input_dir / 'results-pr01.xls'
        if r01_path.exists():
            _, teams = read_excel_data(str(r01_path))
        else:
            print(f"错误: 未找到数据文件")
            return
    else:
        _, teams = read_excel_data(str(ir00_path))
    
    print(f"找到 {len(teams)} 支队伍")
    print(f"队伍列表: {', '.join(teams)}")
    print("\n开始生成各队伍详细分析报告...\n")
    
    # 为每支队伍生成报告
    for i, team in enumerate(teams, 1):
        print(f"[{i}/{len(teams)}] 正在生成 {team} 的分析报告...")
        try:
            analyze_team_detailed(team, str(input_dir), str(output_dir))
            print(f"  ✓ {team} 报告生成成功")
        except Exception as e:
            print(f"  ✗ {team} 报告生成失败: {e}")
    
    print(f"\n所有报告已生成到: {output_dir}")

if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='批量生成所有队伍的详细分析报告')
    parser.add_argument('--input-dir', '-i', type=str, required=True, help='数据输入目录')
    parser.add_argument('--output-dir', '-o', type=str, required=True, help='报告输出目录')
    
    args = parser.parse_args()
    main(args.input_dir, args.output_dir)

