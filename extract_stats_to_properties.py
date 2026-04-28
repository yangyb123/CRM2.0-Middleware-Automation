#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从 Summary.html 报告中提取测试统计数据
生成 Ant 属性文件格式
"""

import re
import sys
from pathlib import Path

def extract_stats_from_html(html_file):
    """从 Summary.html 文件中提取统计数据"""
    try:
        with open(html_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 初始化数据
        stats = {
            'total': '-',
            'success': '-',
            'failure': '-',
            'success_rate': '-'
        }
        
        # 提取总请求数 - 查找 "总请求数" 后面的数字
        total_match = re.search(r'<div class="metric-label">总请求数</div>\s*<div class="metric-value">(\d+(?:\.\d+)?)</div>', content)
        if total_match:
            stats['total'] = total_match.group(1)
        
        # 提取成功请求数 - 查找 "成功请求" 后面的数字
        success_match = re.search(r'<div class="metric-label">成功请求</div>\s*<div class="metric-value">(\d+(?:\.\d+)?)</div>', content)
        if success_match:
            stats['success'] = success_match.group(1)
        
        # 提取失败请求数 - 查找 "失败请求" 后面的数字
        failure_match = re.search(r'<div class="metric-label">失败请求</div>\s*<div class="metric-value">(\d+(?:\.\d+)?)</div>', content)
        if failure_match:
            stats['failure'] = failure_match.group(1)
        
        # 提取成功率 - 从成功请求的 metric-unit 中查找 "成功率: XX.XX%"
        rate_match = re.search(r'<div class="metric-label">成功请求</div>\s*<div class="metric-value">[\d.]+</div>\s*<div class="metric-unit">成功率:\s*([\d.]+)%</div>', content)
        if rate_match:
            stats['success_rate'] = rate_match.group(1)
        
        return stats
    
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return None

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: python extract_stats_to_properties.py <summary_html_file> <output_properties_file>")
        sys.exit(1)
    
    html_file = sys.argv[1]
    properties_file = sys.argv[2]
    
    if not Path(html_file).exists():
        print(f"Error: File not found: {html_file}", file=sys.stderr)
        sys.exit(1)
    
    stats = extract_stats_from_html(html_file)
    
    if stats:
        # 生成 Ant 属性文件格式
        with open(properties_file, 'w', encoding='utf-8') as f:
            f.write(f"stats.total={stats['total']}\n")
            f.write(f"stats.success={stats['success']}\n")
            f.write(f"stats.failure={stats['failure']}\n")
            f.write(f"stats.success_rate={stats['success_rate']}\n")
        
        print(f"✓ 属性文件生成成功: {properties_file}")
        print(f"  总请求数: {stats['total']}")
        print(f"  成功数: {stats['success']}")
        print(f"  失败数: {stats['failure']}")
        print(f"  成功率: {stats['success_rate']}%")
    else:
        print("Error: Failed to extract stats", file=sys.stderr)
        sys.exit(1)
