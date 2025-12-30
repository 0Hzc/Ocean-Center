#!/usr/bin/env python3
"""
IPAR数据诊断脚本
检查IPAR数据在各个处理步骤中的nan/inf情况
"""

import os
import glob
import numpy as np

def check_data_file(filepath, max_samples=5):
    """检查单个数据文件中的nan/inf情况"""
    result = {
        'exists': False,
        'total_lines': 0,
        'nan_count': 0,
        'inf_count': 0,
        'neg_inf_count': 0,
        'zero_count': 0,
        'minus999_count': 0,
        'valid_count': 0,
        'min_val': None,
        'max_val': None,
        'samples_nan': [],
        'samples_inf': [],
        'samples_valid': []
    }

    if not os.path.exists(filepath):
        return result

    result['exists'] = True

    try:
        with open(filepath, 'r') as f:
            lines = f.readlines()

        result['total_lines'] = len(lines)
        values = []

        for i, line in enumerate(lines):
            line = line.strip()
            if not line or line.startswith('/'):
                continue

            try:
                # 处理可能有多列的情况
                parts = line.split()
                for part in parts:
                    try:
                        val = float(part)
                        if np.isnan(val):
                            result['nan_count'] += 1
                            if len(result['samples_nan']) < max_samples:
                                result['samples_nan'].append((i+1, part))
                        elif np.isinf(val):
                            if val > 0:
                                result['inf_count'] += 1
                            else:
                                result['neg_inf_count'] += 1
                            if len(result['samples_inf']) < max_samples:
                                result['samples_inf'].append((i+1, part))
                        elif val == -999:
                            result['minus999_count'] += 1
                        elif val == 0:
                            result['zero_count'] += 1
                        else:
                            result['valid_count'] += 1
                            values.append(val)
                            if len(result['samples_valid']) < max_samples:
                                result['samples_valid'].append((i+1, val))
                    except ValueError:
                        pass
            except Exception:
                pass

        if values:
            result['min_val'] = min(values)
            result['max_val'] = max(values)

    except Exception as e:
        result['error'] = str(e)

    return result

def check_report_file(filepath):
    """检查report文件中的bias和RMS值"""
    result = {
        'exists': False,
        'bias': None,
        'rms': None,
        'effective_count': None
    }

    if not os.path.exists(filepath):
        return result

    result['exists'] = True

    try:
        with open(filepath, 'r') as f:
            for line in f:
                line = line.strip()
                if line.startswith('/bias='):
                    result['bias'] = line.split('=')[1]
                elif line.startswith('/RMS='):
                    result['rms'] = line.split('=')[1]
                elif line.startswith('/Effective pixel count='):
                    result['effective_count'] = line.split('=')[1]
    except Exception as e:
        result['error'] = str(e)

    return result

def check_statistic_file(filepath):
    """检查statistic文件中的统计值"""
    result = {
        'exists': False,
        'header_info': {},
        'stats_line': None
    }

    if not os.path.exists(filepath):
        return result

    result['exists'] = True

    try:
        with open(filepath, 'r') as f:
            lines = f.readlines()

        in_header = False
        for line in lines:
            line = line.strip()
            if line == '/begin header':
                in_header = True
                continue
            if line == '/end header':
                in_header = False
                continue
            if in_header and line.startswith('/'):
                parts = line[1:].split('=', 1)
                if len(parts) == 2:
                    result['header_info'][parts[0]] = parts[1]
            elif not in_header and line and not line.startswith('/'):
                result['stats_line'] = line

    except Exception as e:
        result['error'] = str(e)

    return result

def print_separator(title):
    print("\n" + "="*70)
    print(f" {title}")
    print("="*70)

def main():
    # 配置路径 - 请根据实际情况修改
    base_path = r"/mnt/d/work/海洋中心/output"

    preprocess_path = os.path.join(base_path, "01_sat_preprocess")
    collocation_path = os.path.join(base_path, "03_collocation")

    # 如果目录不存在，尝试其他路径
    if not os.path.exists(base_path):
        base_path = r"D:\work\海洋中心\output"
        preprocess_path = os.path.join(base_path, "01_sat_preprocess")
        collocation_path = os.path.join(base_path, "03_collocation")

    print("="*70)
    print(" IPAR数据诊断报告")
    print("="*70)
    print(f"基础路径: {base_path}")
    print(f"预处理路径: {preprocess_path}")
    print(f"匹配路径: {collocation_path}")

    # ==================== Step 1-2: 预处理数据 ====================
    print_separator("Step 1-2: 预处理数据检查")

    preprocess_files = [
        ("HY1E_ipar_*.txt", "HY卫星IPAR数据"),
        ("TERRA1_ipar_*.txt", "TERRA参考源IPAR数据"),
        ("HY1E_flag1_ipar_*.txt", "HY标志位数据"),
    ]

    for pattern, desc in preprocess_files:
        files = glob.glob(os.path.join(preprocess_path, pattern))
        print(f"\n【{desc}】 ({pattern})")

        if not files:
            # 尝试不区分大小写
            files = glob.glob(os.path.join(preprocess_path, pattern.replace('ipar', '[iI][pP][aA][rR]')))

        if not files:
            print(f"  未找到匹配文件")
            continue

        for f in files[:2]:  # 最多检查2个文件
            print(f"  文件: {os.path.basename(f)}")
            result = check_data_file(f)

            if not result['exists']:
                print(f"    文件不存在")
                continue

            print(f"    总行数: {result['total_lines']}")
            print(f"    NaN数量: {result['nan_count']}")
            print(f"    +Inf数量: {result['inf_count']}")
            print(f"    -Inf数量: {result['neg_inf_count']}")
            print(f"    零值数量: {result['zero_count']}")
            print(f"    -999数量: {result['minus999_count']}")
            print(f"    有效值数量: {result['valid_count']}")

            if result['min_val'] is not None:
                print(f"    有效值范围: [{result['min_val']:.6f}, {result['max_val']:.6f}]")

            if result['samples_nan']:
                print(f"    NaN样例(行号,值): {result['samples_nan']}")
            if result['samples_inf']:
                print(f"    Inf样例(行号,值): {result['samples_inf']}")

    # ==================== Step 3-5: 匹配结果 ====================
    print_separator("Step 3-5: 空间匹配结果检查")

    space_files = glob.glob(os.path.join(collocation_path, "spaceresult_*[iI][pP][aA][rR]*.txt"))
    if not space_files:
        space_files = glob.glob(os.path.join(collocation_path, "spaceresult_*ipar*.txt"))

    if not space_files:
        print("  未找到IPAR空间匹配文件")
    else:
        for f in space_files[:2]:
            print(f"\n  文件: {os.path.basename(f)}")
            result = check_data_file(f)
            print(f"    总行数: {result['total_lines']}")
            print(f"    NaN数量: {result['nan_count']}")
            print(f"    Inf数量: {result['inf_count'] + result['neg_inf_count']}")

    # ==================== Step 6: 验证结果 ====================
    print_separator("Step 6: 验证结果检查")

    # valresult文件
    print("\n【valresult文件】")
    val_files = glob.glob(os.path.join(collocation_path, "valresult_*[iI][pP][aA][rR]*.txt"))
    if not val_files:
        val_files = glob.glob(os.path.join(collocation_path, "valresult_*ipar*.txt"))

    if not val_files:
        print("  未找到IPAR验证结果文件")
    else:
        for f in val_files[:2]:
            print(f"\n  文件: {os.path.basename(f)}")
            result = check_data_file(f)
            print(f"    总行数: {result['total_lines']}")
            print(f"    NaN数量: {result['nan_count']}")
            print(f"    +Inf数量: {result['inf_count']}")
            print(f"    -Inf数量: {result['neg_inf_count']}")

            if result['samples_nan']:
                print(f"    NaN样例: {result['samples_nan']}")
            if result['samples_inf']:
                print(f"    Inf样例: {result['samples_inf']}")
            if result['min_val'] is not None:
                print(f"    有效值范围: [{result['min_val']:.6f}, {result['max_val']:.6f}]")

    # statistic文件
    print("\n【statistic文件】")
    stat_files = glob.glob(os.path.join(collocation_path, "statistic_*[iI][pP][aA][rR]*.txt"))
    if not stat_files:
        stat_files = glob.glob(os.path.join(collocation_path, "statistic_*ipar*.txt"))

    if not stat_files:
        print("  未找到IPAR统计文件")
    else:
        for f in stat_files[:2]:
            print(f"\n  文件: {os.path.basename(f)}")
            result = check_statistic_file(f)

            if result['header_info']:
                print(f"    有效像元数: {result['header_info'].get('Effective pixel count', 'N/A')}")
                print(f"    验证结果: {result['header_info'].get('validation result', 'N/A')}")

            if result['stats_line']:
                print(f"    统计行(bias/STD/RMS/R): {result['stats_line']}")
                # 解析统计值
                parts = result['stats_line'].split('\t')
                if len(parts) >= 4:
                    print(f"      bias={parts[0]}, STD={parts[1]}, RMS={parts[2]}, R={parts[3]}")

    # report文件
    print("\n【report文件 (表1数据来源)】")
    report_files = glob.glob(os.path.join(collocation_path, "*[iI][pP][aA][rR]*report*.txt"))
    if not report_files:
        report_files = glob.glob(os.path.join(collocation_path, "*ipar*report*.txt"))
    if not report_files:
        report_files = glob.glob(os.path.join(collocation_path, "*IPAR*report*.txt"))

    if not report_files:
        print("  未找到IPAR报告文件")
    else:
        for f in report_files[:2]:
            print(f"\n  文件: {os.path.basename(f)}")
            result = check_report_file(f)
            print(f"    bias: {result['bias']}")
            print(f"    RMS: {result['rms']}")
            print(f"    有效像元数: {result['effective_count']}")

    # ==================== 诊断结论 ====================
    print_separator("诊断结论")
    print("""
请根据上述检查结果判断问题出在哪个步骤:

1. 如果【预处理数据】中就有nan/inf → 问题在Step1-2数据读取
2. 如果【空间匹配】中出现nan/inf → 问题在Step3-5匹配过程
3. 如果【valresult】中出现nan/inf → 问题在Step6计算差异时
4. 如果【statistic】的统计行有nan/inf → 问题在Step6计算bias/RMS时

常见原因:
- 原始数据包含无效值(nan/inf/-999)
- 除法时分母接近0导致溢出
- 数据类型转换问题
""")

if __name__ == "__main__":
    main()
