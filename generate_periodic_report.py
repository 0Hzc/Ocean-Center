#!/usr/bin/env python3
"""
生成月度/季度/年度汇总报告
用于汇总和分析指定时间范围内的日报告数据
"""
import os
import glob
import re
from datetime import datetime, timedelta
from collections import defaultdict
import argparse
import json


def collect_daily_reports(start_date, end_date, report_dir='./output/05_reports'):
    """
    收集指定时间范围内的日报告

    :param start_date: 开始日期
    :param end_date: 结束日期
    :param report_dir: 报告目录
    :return: 报告文件列表和统计信息
    """
    reports = []
    report_data = defaultdict(list)

    # 扫描所有report文件
    report_pattern = re.compile(r'report_(\w+)_(\w+)_(\w+)_(\d{8})_(\d{6})\.txt')

    if not os.path.exists(report_dir):
        print(f"警告: 报告目录不存在: {report_dir}")
        return reports, report_data

    for filename in os.listdir(report_dir):
        match = report_pattern.match(filename)
        if match and filename.endswith('.txt'):
            satellite, reference, var_name, date_str, time_str = match.groups()

            try:
                file_date = datetime.strptime(date_str, '%Y%m%d')

                # 检查是否在时间范围内
                if start_date <= file_date <= end_date:
                    filepath = os.path.join(report_dir, filename)
                    reports.append(filepath)

                    # 读取报告数据
                    data = parse_report_file(filepath)
                    if data:
                        data['date'] = date_str
                        data['satellite'] = satellite
                        data['reference'] = reference
                        data['var_name'] = var_name
                        report_data[var_name].append(data)
            except ValueError:
                continue

    print(f"收集到 {len(reports)} 个报告文件")
    return reports, report_data


def parse_report_file(filepath):
    """
    解析报告文件,提取关键指标

    :param filepath: 报告文件路径
    :return: 包含指标的字典
    """
    data = {}
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line.startswith('/bias='):
                    data['bias'] = float(line.split('=')[1])
                elif line.startswith('/RMS='):
                    data['rms'] = float(line.split('=')[1])
                elif line.startswith('/Effective pixel count='):
                    data['pixel_count'] = int(line.split('=')[1])
                elif line.startswith('/valresult='):
                    data['valresult'] = int(line.split('=')[1])
                elif line.startswith('/Valid Ratio='):
                    data['valid_ratio'] = float(line.split('=')[1])
                elif line.startswith('/CV Value='):
                    data['cv_value'] = float(line.split('=')[1])
                elif line.startswith('/Relative Bias='):
                    data['relative_bias'] = float(line.split('=')[1])
    except Exception as e:
        print(f"解析文件失败 {filepath}: {e}")
        return None

    return data if data else None


def analyze_reports(report_data):
    """
    分析报告数据,生成统计信息

    :param report_data: 报告数据字典
    :return: 统计结果字典
    """
    stats = {}

    for var_name, data_list in report_data.items():
        if not data_list:
            continue

        var_stats = {
            'count': len(data_list),
            'avg_bias': 0,
            'avg_rms': 0,
            'avg_pixel_count': 0,
            'min_bias': float('inf'),
            'max_bias': float('-inf'),
            'min_rms': float('inf'),
            'max_rms': float('-inf')
        }

        bias_values = []
        rms_values = []
        pixel_counts = []

        for data in data_list:
            if 'bias' in data:
                bias_values.append(data['bias'])
                var_stats['min_bias'] = min(var_stats['min_bias'], data['bias'])
                var_stats['max_bias'] = max(var_stats['max_bias'], data['bias'])

            if 'rms' in data:
                rms_values.append(data['rms'])
                var_stats['min_rms'] = min(var_stats['min_rms'], data['rms'])
                var_stats['max_rms'] = max(var_stats['max_rms'], data['rms'])

            if 'pixel_count' in data:
                pixel_counts.append(data['pixel_count'])

        # 计算平均值
        if bias_values:
            var_stats['avg_bias'] = sum(bias_values) / len(bias_values)
        if rms_values:
            var_stats['avg_rms'] = sum(rms_values) / len(rms_values)
        if pixel_counts:
            var_stats['avg_pixel_count'] = sum(pixel_counts) / len(pixel_counts)
            var_stats['total_pixel_count'] = sum(pixel_counts)

        stats[var_name] = var_stats

    return stats


def generate_monthly_report(year, month, output_dir='./output/05_reports'):
    """
    生成月度报告

    :param year: 年份
    :param month: 月份
    :param output_dir: 输出目录
    """
    print(f"\n生成 {year}年{month}月 的月度报告...")

    start_date = datetime(year, month, 1)
    if month == 12:
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = datetime(year, month + 1, 1) - timedelta(days=1)

    reports, report_data = collect_daily_reports(start_date, end_date, output_dir)

    if not reports:
        print("警告: 未找到任何报告文件")
        return

    # 统计分析
    stats = analyze_reports(report_data)

    # 生成汇总报告文件
    summary_file = os.path.join(output_dir, f'monthly_summary_{year}{month:02d}.json')
    summary_data = {
        'report_type': 'monthly',
        'year': year,
        'month': month,
        'start_date': start_date.strftime('%Y-%m-%d'),
        'end_date': end_date.strftime('%Y-%m-%d'),
        'total_reports': len(reports),
        'generated_at': datetime.now().isoformat(),
        'statistics': stats
    }

    with open(summary_file, 'w', encoding='utf-8') as f:
        json.dump(summary_data, f, indent=2, ensure_ascii=False)

    print(f"月度汇总报告已生成: {summary_file}")
    print_summary_stats(stats)


def generate_quarterly_report(year, quarter, output_dir='./output/05_reports'):
    """
    生成季度报告

    :param year: 年份
    :param quarter: 季度 (1-4)
    :param output_dir: 输出目录
    """
    if quarter < 1 or quarter > 4:
        print("错误: 季度必须在1-4之间")
        return

    print(f"\n生成 {year}年第{quarter}季度 的季度报告...")

    # 计算季度的起始和结束日期
    start_month = (quarter - 1) * 3 + 1
    start_date = datetime(year, start_month, 1)

    end_month = quarter * 3
    if end_month == 12:
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = datetime(year, end_month + 1, 1) - timedelta(days=1)

    reports, report_data = collect_daily_reports(start_date, end_date, output_dir)

    if not reports:
        print("警告: 未找到任何报告文件")
        return

    # 统计分析
    stats = analyze_reports(report_data)

    # 生成汇总报告文件
    summary_file = os.path.join(output_dir, f'quarterly_summary_{year}Q{quarter}.json')
    summary_data = {
        'report_type': 'quarterly',
        'year': year,
        'quarter': quarter,
        'start_date': start_date.strftime('%Y-%m-%d'),
        'end_date': end_date.strftime('%Y-%m-%d'),
        'total_reports': len(reports),
        'generated_at': datetime.now().isoformat(),
        'statistics': stats
    }

    with open(summary_file, 'w', encoding='utf-8') as f:
        json.dump(summary_data, f, indent=2, ensure_ascii=False)

    print(f"季度汇总报告已生成: {summary_file}")
    print_summary_stats(stats)


def generate_annual_report(year, output_dir='./output/05_reports'):
    """
    生成年度报告

    :param year: 年份
    :param output_dir: 输出目录
    """
    print(f"\n生成 {year}年 的年度报告...")

    start_date = datetime(year, 1, 1)
    end_date = datetime(year, 12, 31)

    reports, report_data = collect_daily_reports(start_date, end_date, output_dir)

    if not reports:
        print("警告: 未找到任何报告文件")
        return

    # 统计分析
    stats = analyze_reports(report_data)

    # 生成汇总报告文件
    summary_file = os.path.join(output_dir, f'annual_summary_{year}.json')
    summary_data = {
        'report_type': 'annual',
        'year': year,
        'start_date': start_date.strftime('%Y-%m-%d'),
        'end_date': end_date.strftime('%Y-%m-%d'),
        'total_reports': len(reports),
        'generated_at': datetime.now().isoformat(),
        'statistics': stats
    }

    with open(summary_file, 'w', encoding='utf-8') as f:
        json.dump(summary_data, f, indent=2, ensure_ascii=False)

    print(f"年度汇总报告已生成: {summary_file}")
    print_summary_stats(stats)


def print_summary_stats(stats):
    """打印统计摘要"""
    print("\n统计摘要:")
    print("-" * 80)
    print(f"{'产品':<15} {'样本数':<10} {'平均Bias':<15} {'平均RMS':<15} {'平均像素数':<15}")
    print("-" * 80)

    for var_name, var_stats in stats.items():
        print(f"{var_name:<15} {var_stats['count']:<10} "
              f"{var_stats['avg_bias']:<15.4f} {var_stats['avg_rms']:<15.4f} "
              f"{var_stats.get('avg_pixel_count', 0):<15.0f}")

    print("-" * 80)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='生成周期性汇总报告',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 生成2025年3月的月度报告
  python generate_periodic_report.py --type monthly --year 2025 --month 3

  # 生成2025年第1季度的季度报告
  python generate_periodic_report.py --type quarterly --year 2025 --quarter 1

  # 生成2025年的年度报告
  python generate_periodic_report.py --type annual --year 2025

  # 指定输出目录
  python generate_periodic_report.py --type monthly --year 2025 --month 3 --output ./reports
        """
    )

    parser.add_argument('--type', choices=['monthly', 'quarterly', 'annual'],
                       required=True, help='报告类型')
    parser.add_argument('--year', type=int, required=True, help='年份')
    parser.add_argument('--month', type=int, help='月份 (1-12, 用于月度报告)')
    parser.add_argument('--quarter', type=int, choices=[1, 2, 3, 4],
                       help='季度 (1-4, 用于季度报告)')
    parser.add_argument('--output', type=str, default='./output/05_reports',
                       help='输出目录 (默认: ./output/05_reports)')

    args = parser.parse_args()

    # 确保输出目录存在
    os.makedirs(args.output, exist_ok=True)

    try:
        if args.type == 'monthly':
            if not args.month:
                parser.error("月度报告需要指定 --month 参数")
            if args.month < 1 or args.month > 12:
                parser.error("月份必须在 1-12 之间")
            generate_monthly_report(args.year, args.month, args.output)

        elif args.type == 'quarterly':
            if not args.quarter:
                parser.error("季度报告需要指定 --quarter 参数")
            generate_quarterly_report(args.year, args.quarter, args.output)

        elif args.type == 'annual':
            generate_annual_report(args.year, args.output)

        print("\n报告生成完成！")

    except Exception as e:
        print(f"\n错误: 生成报告时发生异常: {e}")
        import traceback
        traceback.print_exc()
        exit(1)
