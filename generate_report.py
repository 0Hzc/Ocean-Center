#!/usr/bin/env python3
"""
月季年报告生成器 - V5最终版
直接从map文件生成GEO图:
1. map文件已包含: latitude\tlongitude\terror
2. map文件位置: output/04_visualization/map_*.txt
3. 聚合多天数据生成地理分布图
"""
import os
import sys
import re
import json
import argparse
import configparser
from datetime import datetime
from collections import defaultdict
from calendar import monthrange
import glob
import subprocess

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

try:
    from mpl_toolkits.basemap import Basemap
    from scipy.interpolate import griddata
    BASEMAP_AVAILABLE = True
except ImportError:
    BASEMAP_AVAILABLE = False
    print("[WARNING] Basemap未安装，将生成简化版地理图")


# ============================================================================
# 产品配置
# ============================================================================

VAR_CONFIGS = {
    'AQUA': {
        'sst': {'sources': ['AQUA'], 'unit': 'K'},
        'chl': {'sources': ['AQUA'], 'unit': 'mg/m³'},
        'Rrs412': {'sources': ['AQUA'], 'unit': 'sr⁻¹'},
        'Rrs443': {'sources': ['AQUA'], 'unit': 'sr⁻¹'},
        'Rrs490': {'sources': ['AQUA'], 'unit': 'sr⁻¹'},
        'Rrs520': {'sources': ['AQUA'], 'unit': 'sr⁻¹'},
        'Rrs565': {'sources': ['AQUA'], 'unit': 'sr⁻¹'},
        'Rrs670': {'sources': ['AQUA'], 'unit': 'sr⁻¹'},
        'AOT': {'sources': ['AQUA'], 'unit': ''},
    },
    'TERRA': {
        'sst': {'sources': ['TERRA'], 'unit': 'K'},
        'chl': {'sources': ['TERRA'], 'unit': 'mg/m³'},
        'Rrs412': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'Rrs443': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'Rrs490': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'Rrs520': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'Rrs565': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'Rrs670': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'AOT': {'sources': ['TERRA'], 'unit': ''},
    },
    'SNPP': {
        'sst': {'sources': ['SNPP'], 'unit': 'K'},
        'chl': {'sources': ['SNPP'], 'unit': 'mg/m³'},
        'Rrs412': {'sources': ['SNPP'], 'unit': 'sr⁻¹'},
        'Rrs443': {'sources': ['SNPP'], 'unit': 'sr⁻¹'},
        'Rrs490': {'sources': ['SNPP'], 'unit': 'sr⁻¹'},
        'Rrs565': {'sources': ['SNPP'], 'unit': 'sr⁻¹'},
        'Rrs670': {'sources': ['SNPP'], 'unit': 'sr⁻¹'},
        'AOT': {'sources': ['SNPP'], 'unit': ''},
    },
    'JPSS': {
        'sst': {'sources': ['JPSS'], 'unit': 'K'},
        'chl': {'sources': ['JPSS'], 'unit': 'mg/m³'},
        'Rrs412': {'sources': ['JPSS'], 'unit': 'sr⁻¹'},
        'Rrs443': {'sources': ['JPSS'], 'unit': 'sr⁻¹'},
        'Rrs490': {'sources': ['JPSS'], 'unit': 'sr⁻¹'},
        'Rrs565': {'sources': ['JPSS'], 'unit': 'sr⁻¹'},
        'Rrs670': {'sources': ['JPSS'], 'unit': 'sr⁻¹'},
        'AOT': {'sources': ['JPSS'], 'unit': ''},
    },
    'HY1C': {
        'sst': {'sources': ['HY1C'], 'unit': 'K'},
        'chl': {'sources': ['HY1C'], 'unit': 'mg/m³'},
        'Rrs412': {'sources': ['HY1C'], 'unit': 'sr⁻¹'},
        'Rrs443': {'sources': ['HY1C'], 'unit': 'sr⁻¹'},
        'Rrs490': {'sources': ['HY1C'], 'unit': 'sr⁻¹'},
        'Rrs520': {'sources': ['HY1C'], 'unit': 'sr⁻¹'},
        'Rrs565': {'sources': ['HY1C'], 'unit': 'sr⁻¹'},
        'Rrs670': {'sources': ['HY1C'], 'unit': 'sr⁻¹'},
        'AOT': {'sources': ['HY1C'], 'unit': ''},
        'TSM': {'sources': ['HY1C'], 'unit': 'mg/L'},
        'CDOM': {'sources': ['HY1C'], 'unit': '1/m'},
    },
    'HY1D': {
        'sst': {'sources': ['HY1D'], 'unit': 'K'},
        'chl': {'sources': ['HY1D'], 'unit': 'mg/m³'},
        'Rrs412': {'sources': ['HY1D'], 'unit': 'sr⁻¹'},
        'Rrs443': {'sources': ['HY1D'], 'unit': 'sr⁻¹'},
        'Rrs490': {'sources': ['HY1D'], 'unit': 'sr⁻¹'},
        'Rrs520': {'sources': ['HY1D'], 'unit': 'sr⁻¹'},
        'Rrs565': {'sources': ['HY1D'], 'unit': 'sr⁻¹'},
        'Rrs670': {'sources': ['HY1D'], 'unit': 'sr⁻¹'},
        'AOT': {'sources': ['HY1D'], 'unit': ''},
        'TSM': {'sources': ['HY1D'], 'unit': 'mg/L'},
        'CDOM': {'sources': ['HY1D'], 'unit': '1/m'},
    },
    'HY1E': {
        'sst': {'sources': ['HY1E'], 'unit': 'K'},
        'chl': {'sources': ['HY1E'], 'unit': 'mg/m³'},
        'Rrs412': {'sources': ['HY1E'], 'unit': 'sr⁻¹'},
        'Rrs443': {'sources': ['HY1E'], 'unit': 'sr⁻¹'},
        'Rrs490': {'sources': ['HY1E'], 'unit': 'sr⁻¹'},
        'Rrs520': {'sources': ['HY1E'], 'unit': 'sr⁻¹'},
        'Rrs565': {'sources': ['HY1E'], 'unit': 'sr⁻¹'},
        'Rrs620': {'sources': ['HY1E'], 'unit': 'sr⁻¹'},
        'Rrs665': {'sources': ['HY1E'], 'unit': 'sr⁻¹'},
        'Rrs670': {'sources': ['HY1E'], 'unit': 'sr⁻¹'},
        'Rrs681': {'sources': ['HY1E'], 'unit': 'sr⁻¹'},
        'Rrs750': {'sources': ['HY1E'], 'unit': 'sr⁻¹'},
        'AOT': {'sources': ['HY1E'], 'unit': ''},
        'TSM': {'sources': ['HY1E'], 'unit': 'mg/L'},
        'CDOM': {'sources': ['HY1E'], 'unit': '1/m'},
        'Kd': {'sources': ['HY1E'], 'unit': 'm⁻¹'},
        'IPAR': {'sources': ['HY1E'], 'unit': 'Einstein m⁻² s⁻¹'},
    },
    'XC': {
        'sst': {'sources': ['XC'], 'unit': 'K'},
        'chl': {'sources': ['XC'], 'unit': 'mg/m³'},
        'AOT': {'sources': ['XC'], 'unit': ''},
    }
}

PRODUCT_NAMES = {
    'AOT': '气溶胶光学厚度',
    'chl': '叶绿素浓度',
    'sst': '海表温度',
    'Rrs412': '412nm遥感反射率',
    'Rrs443': '443nm遥感反射率',
    'Rrs490': '490nm遥感反射率',
    'Rrs520': '520nm遥感反射率',
    'Rrs565': '565nm遥感反射率',
    'Rrs620': '620nm遥感反射率',
    'Rrs665': '665nm遥感反射率',
    'Rrs670': '670nm遥感反射率',
    'Rrs681': '681nm遥感反射率',
    'Rrs750': '750nm遥感反射率',
    'TSM': '总悬浮物浓度',
    'CDOM': '有色溶解有机物吸收系数',
    'Kd': '漫射衰减系数',
    'IPAR': '光合有效辐射'
}


# ============================================================================
# 字体管理
# ============================================================================

class FontManager:
    def __init__(self, config):
        self.config = config
        self.font_loaded = False
        self._load_font()
    
    def _load_font(self):
        try:
            font_path = self.config.get('font', 'font_path')
            font_path = font_path.replace('\\', '/')
            
            print(f"[INFO] 加载字体: {font_path}")
            
            if os.path.exists(font_path):
                from matplotlib import font_manager
                font_manager.fontManager.addfont(font_path)
                plt.rcParams['font.sans-serif'] = ['SimHei']
                plt.rcParams['axes.unicode_minus'] = False
                self.font_loaded = True
                print(f"[INFO] ✓ 字体加载成功")
            else:
                print(f"[WARNING] 字体文件不存在: {font_path}")
        except Exception as e:
            print(f"[WARNING] 字体加载失败: {e}")


# ============================================================================
# 报告分类器
# ============================================================================

class ReportClassifier:
    PATTERN_WITH_TIME = re.compile(
        r'([A-Z0-9]+)_COCTS_([A-Z0-9]+)_val_report_(\d{8})_(\d{6})\.pdf$'
    )
    
    PATTERN_WITHOUT_TIME = re.compile(
        r'([A-Z0-9]+)_COCTS_([A-Z0-9]+)_val_report_(\d{8})\.pdf$'
    )
    
    @classmethod
    def parse_report_filename(cls, filename):
        match = cls.PATTERN_WITH_TIME.search(filename)
        if match:
            return {
                'satellite': match.group(1),
                'source': match.group(2),
                'date': match.group(3),
                'time': match.group(4),
                'has_time': True
            }
        
        match = cls.PATTERN_WITHOUT_TIME.search(filename)
        if match:
            return {
                'satellite': match.group(1),
                'source': match.group(2),
                'date': match.group(3),
                'time': '000000',
                'has_time': False
            }
        
        return None
    
    @classmethod
    def get_validation_label(cls, satellite, source):
        if source.upper() == 'XC':
            return f"{satellite}-现场"
        else:
            return f"{satellite}-{source}"


# ============================================================================
# 数据聚合器
# ============================================================================

class DataAggregator:
    """数据聚合器"""
    
    def __init__(self, reports_dir, space_size, time_size, debug=False):
        self.reports_dir = reports_dir
        self.space_size = space_size
        self.time_size = time_size
        self.debug = debug
    
    def aggregate_validation_data(self, classified_reports):
        """聚合验证数据"""
        aggregated = {}
        
        for val_label, date_reports in classified_reports['by_validation'].items():
            print(f"[INFO] 聚合: {val_label}")
            
            report_info = list(date_reports.values())[0][0]
            satellite = report_info['satellite']
            source = report_info['source']
            
            products_data = self._aggregate_products(date_reports, satellite, source)
            
            if products_data:
                aggregated[val_label] = {
                    'satellite': satellite,
                    'source': source,
                    'products': products_data,
                    'dates': sorted(date_reports.keys())
                }
        
        return aggregated
    
    def _aggregate_products(self, date_reports, satellite, source):
        """按产品聚合数据"""
        products_data = {}
        
        # 获取所有report文件
        all_report_files = []
        for date_str in date_reports.keys():
            patterns = [
                f'report_{satellite}_{source}_*_{date_str}_*.txt',
                f'{satellite}_COCTS_{source}_*_report_{date_str}_*.txt'
            ]
            
            for pattern in patterns:
                full_pattern = os.path.join(self.reports_dir, pattern)
                all_report_files.extend(glob.glob(full_pattern))
        
        all_report_files = list(set(all_report_files))
        
        # 解析report文件
        for report_file in all_report_files:
            basename = os.path.basename(report_file)
            
            # 从文件名提取产品名
            if basename.startswith('report_'):
                parts = basename.split('_')
                if len(parts) >= 4:
                    product = parts[3]
            else:
                parts = basename.split('_')
                if len(parts) >= 4:
                    product = parts[3]
            
            if product:
                data = self._parse_report_file(report_file)
                if data:
                    if product not in products_data:
                        products_data[product] = {
                            'bias_list': [],
                            'rms_list': [],
                            'n_list': []
                        }
                    
                    products_data[product]['bias_list'].append(data['bias'])
                    products_data[product]['rms_list'].append(data['rms'])
                    products_data[product]['n_list'].append(data['n'])
        
        # 计算总体统计
        for product, data in products_data.items():
            if data['bias_list']:
                total_n = sum(data['n_list'])
                weighted_bias = sum(b * n for b, n in zip(data['bias_list'], data['n_list'])) / total_n if total_n > 0 else 0
                weighted_rms = sum(r * n for r, n in zip(data['rms_list'], data['n_list'])) / total_n if total_n > 0 else 0
                
                data['bias'] = weighted_bias
                data['rms'] = weighted_rms
                data['n'] = total_n
                data['space_window'] = f"{self.space_size}*{self.space_size}"
                data['time_window'] = f"{self.time_size}"
                
                print(f"[INFO]   {product}: bias={weighted_bias:.2f}, RMS={weighted_rms:.2f}, n={total_n}")
        
        return products_data
    
    def _parse_report_file(self, filepath):
        """解析report文件"""
        try:
            bias = None
            rms = None
            n = None
            
            with open(filepath, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line.startswith('/bias='):
                        bias = float(line.split('=')[1])
                    elif line.startswith('/RMS='):
                        rms = float(line.split('=')[1])
                    elif line.startswith('/Effective pixel count='):
                        n = int(line.split('=')[1])
                    elif line.startswith('/valresult=') and n is None:
                        n = int(line.split('=')[1])
            
            if bias is not None and rms is not None and n is not None:
                return {'bias': bias, 'rms': rms, 'n': n}
            
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] 解析失败 {filepath}: {e}")
        
        return None


# ============================================================================
# 图表生成器 - 从matchup文件生成GEO图
# ============================================================================

class ChartGenerator:
    def __init__(self, output_dir, collocation_dir, sat_input_dir, font_manager, debug=False):
        self.output_dir = output_dir
        self.collocation_dir = collocation_dir
        self.sat_input_dir = sat_input_dir
        self.font_manager = font_manager
        self.debug = debug
    
    def generate_pie_chart(self, satellite, source, product, matchup_files, timestamp):
        """生成饼图"""
        differences = []
        for matchup_file in matchup_files:
            diffs = self._parse_matchup_file(matchup_file)
            if diffs:
                differences.extend(diffs)

        if not differences:
            return None

        product_name = PRODUCT_NAMES.get(product, product)

        # 问题5: SST产品使用K作为单位，其他产品使用%
        product_lower = product.lower()
        if 'sst' in product_lower:
            # SST使用开尔文(K)
            difference_counts = {
                '0~5K': 0,
                '5~10K': 0,
                '10~15K': 0,
                '15~20K': 0,
                '≥20K': 0
            }
        else:
            # 其他产品使用百分比
            difference_counts = {
                '0~5%': 0,
                '5~10%': 0,
                '10~15%': 0,
                '15~20%': 0,
                '≥20%': 0
            }

        for diff in differences:
            abs_diff = abs(diff)
            if abs_diff < 5:
                key = list(difference_counts.keys())[0]
                difference_counts[key] += 1
            elif abs_diff < 10:
                key = list(difference_counts.keys())[1]
                difference_counts[key] += 1
            elif abs_diff < 15:
                key = list(difference_counts.keys())[2]
                difference_counts[key] += 1
            elif abs_diff < 20:
                key = list(difference_counts.keys())[3]
                difference_counts[key] += 1
            else:
                key = list(difference_counts.keys())[4]
                difference_counts[key] += 1
        
        plt.figure(figsize=(10, 8))
        sizes = list(difference_counts.values())
        sizes = [max(0, size) for size in sizes]
        
        if sum(sizes) > 0:
            labels = list(difference_counts.keys())
            plt.pie(sizes, labels=labels, autopct='%1.1f%%')
            plt.title(f"{product_name}检验结果情况")
            
            output_path = os.path.join(
                self.output_dir,
                f'{satellite}_COCTS_{source}_{product}_PIE_{timestamp}.jpg'
            )
            plt.savefig(output_path)
            plt.close()
            
            print(f"[INFO] ✓ 生成饼图: {os.path.basename(output_path)}")
            return output_path
        
        plt.close()
        return None
    
    def generate_geo_map(self, satellite, source, product, date_strs, timestamp):
        """
        生成地理分布图 - 从map文件读取
        map文件格式：latitude\tlongitude\terror
        map文件位置：output/04_visualization/map_*.txt
        """
        if self.debug:
            print(f"[DEBUG] ========== 开始生成GEO图 ==========")
            print(f"[DEBUG] 产品: {satellite}_{source}_{product}")
            print(f"[DEBUG] 日期数: {len(date_strs)}")
        
        try:
            # 查找所有对应的map文件
            map_files = []
            
            for date_str in date_strs:
                # map文件命名格式: map_{SAT}_{SRC}_{PROD}_{DATE}{TIME}.txt
                patterns = [
                    f'map_{satellite}_{source}_{product}_{date_str}*.txt',
                    f'map_{satellite}_{source}_{product.upper()}_{date_str}*.txt'
                ]
                
                for pattern in patterns:
                    full_pattern = os.path.join(self.output_dir, pattern)
                    found = glob.glob(full_pattern)
                    map_files.extend(found)
                    if self.debug and found:
                        print(f"[DEBUG] 找到map文件: {[os.path.basename(f) for f in found]}")
            
            map_files = list(set(map_files))
            
            if not map_files:
                if self.debug:
                    print(f"[DEBUG] 未找到map文件")
                return None
            
            if self.debug:
                print(f"[DEBUG] 总计map文件: {len(map_files)}")
            
            # 聚合所有map文件的数据
            all_latitudes = []
            all_longitudes = []
            all_errors = []
            
            for map_file in map_files:
                if self.debug:
                    print(f"[DEBUG] 读取map文件: {os.path.basename(map_file)}")
                
                lats, lons, errs = self._read_map_file(map_file)
                
                if lats:
                    all_latitudes.extend(lats)
                    all_longitudes.extend(lons)
                    all_errors.extend(errs)
                    if self.debug:
                        print(f"[DEBUG] 读取到 {len(lats)} 个数据点")
            
            if not all_latitudes:
                if self.debug:
                    print(f"[DEBUG] 没有有效数据点")
                return None
            
            if self.debug:
                print(f"[DEBUG] 总共聚合 {len(all_latitudes)} 个数据点")
            
            # 绘制地理图
            output_path = os.path.join(
                self.output_dir,
                f'{satellite}_COCTS_{source}_{product}_GEO_{timestamp}.jpg'
            )
            
            self._plot_error_map(all_latitudes, all_longitudes, all_errors, product, output_path)
            
            print(f"[INFO] ✓ 生成地理图: {os.path.basename(output_path)}")
            return output_path
            
        except Exception as e:
            if self.debug:
                import traceback
                print(f"[DEBUG] 地理图生成失败:")
                traceback.print_exc()
            return None
    
    def _read_map_file(self, filepath):
        """
        读取map文件
        格式: latitude\tlongitude\terror
        """
        try:
            latitudes = []
            longitudes = []
            errors = []
            
            with open(filepath, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    
                    try:
                        parts = line.split('\t')
                        if len(parts) >= 3:
                            lat = float(parts[0])
                            lon = float(parts[1])
                            err = float(parts[2])
                            
                            latitudes.append(lat)
                            longitudes.append(lon)
                            errors.append(err)
                    except (ValueError, IndexError):
                        continue
            
            return latitudes, longitudes, errors
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] 读取map文件失败: {e}")
            return [], [], []
    
    def _plot_error_map(self, latitudes, longitudes, errors, product, output_path):
        """绘制误差地图 - 按照H1CD.py的plot_error_map"""
        plt.figure(figsize=(10, 8))
        
        product_name = PRODUCT_NAMES.get(product, product)
        product_lower = product.lower()
        
        # 设置误差范围
        if 'sst' in product_lower:
            valid_errors = [err for err in errors if not np.isnan(err)]
            if valid_errors:
                error_max = max(valid_errors)
                max_error = error_max * 1.2
            else:
                max_error = 10
            vmin, vmax = 0, max_error
        else:
            max_error = 100
            vmin, vmax = 0, 100
        
        # 过滤数据
        if max_error is not None:
            valid_indices = [i for i, err in enumerate(errors) if abs(err) <= max_error]
            if not valid_indices:
                plt.text(0.5, 0.5, '无有效数据', ha='center', va='center')
                plt.savefig(output_path)
                plt.close()
                return
            
            latitudes = [latitudes[i] for i in valid_indices]
            longitudes = [longitudes[i] for i in valid_indices]
            errors = [abs(errors[i]) for i in valid_indices]
        
        min_lat, max_lat = min(latitudes), max(latitudes)
        min_lon, max_lon = min(longitudes), max(longitudes)
        
        if BASEMAP_AVAILABLE:
            # 使用Basemap绘图
            m = Basemap(projection='cyl', llcrnrlat=min_lat, urcrnrlat=max_lat,
                        llcrnrlon=min_lon, urcrnrlon=max_lon, resolution='l')
            
            m.drawcoastlines(color='gray')
            m.fillcontinents(color='burlywood', lake_color='lightblue')
            m.drawparallels(np.arange(round(min_lat), round(max_lat)+1, 2), 
                            labels=[1,0,0,0], fmt='%.1f°N', fontsize=8)
            m.drawmeridians(np.arange(round(min_lon), round(max_lon)+1, 2), 
                            labels=[0,0,0,1], fmt='%.1f°E', fontsize=8)
            
            x, y = m(longitudes, latitudes)
            scatter = m.scatter(x, y, c=errors, cmap='jet', s=20, alpha=0.6, 
                               vmin=vmin, vmax=vmax)
        else:
            # 简化版：直接散点图
            scatter = plt.scatter(longitudes, latitudes, c=errors, cmap='jet', 
                                 s=20, alpha=0.6, vmin=vmin, vmax=vmax)
            plt.xlabel('Longitude')
            plt.ylabel('Latitude')
            plt.grid(True, alpha=0.3)
        
        cbar = plt.colorbar(scatter, orientation='vertical', pad=0.05)
        # 问题5: SST产品使用K作为单位，其他产品使用%
        if 'sst' in product_lower:
            cbar.set_label('Error (K)')
        else:
            cbar.set_label('Error (%)')

        plt.title(f"{product_name}检验误差分布")
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close()
    
    def _parse_matchup_file(self, filepath):
        """解析matchup文件用于饼图（读取误差列）"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            differences = []
            in_header = True
            
            for line in lines:
                line = line.strip()
                
                # 跳过header
                if line.startswith('/'):
                    if line.startswith('/end header'):
                        in_header = False
                    continue
                
                if in_header or not line:
                    continue
                
                try:
                    parts = line.split('\t')
                    if len(parts) >= 4:
                        # 第4列是误差
                        diff = float(parts[3])
                        differences.append(diff)
                except (ValueError, IndexError):
                    continue

            return differences
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] 解析matchup失败: {e}")
            return []

    def generate_time_series_chart(self, satellite, source, product, report_files, timestamp):
        """
        生成时间序列图

        :param satellite: 卫星名称
        :param source: 数据源
        :param product: 产品名称
        :param report_files: report文件列表
        :param timestamp: 时间戳
        """
        if not report_files:
            if self.debug:
                print(f"[DEBUG]   {product}: 无数据文件，跳过时序图")
            return None

        times = []
        biases = []

        # 从report文件中读取时间和bias数据
        for report_file in sorted(report_files):
            try:
                basename = os.path.basename(report_file)
                # 从文件名中提取时间：report_HY1C_TERRA_sst_20251009_104500.txt
                parts = basename.split('_')
                if len(parts) >= 5:
                    date_str = parts[-2]  # 20251009
                    time_str = parts[-1].replace('.txt', '')  # 104500

                    # 读取bias值
                    bias = None
                    with open(report_file, 'r', encoding='utf-8') as f:
                        for line in f:
                            line = line.strip()
                            if line.startswith('/bias='):
                                bias = float(line.split('=')[1])
                                break

                    if bias is not None:
                        # 组合时间字符串
                        time_point = f"{date_str}_{time_str}"
                        times.append(pd.to_datetime(time_point, format='%Y%m%d_%H%M%S'))
                        biases.append(bias)
            except Exception as e:
                if self.debug:
                    print(f"[DEBUG] 解析report文件失败 {report_file}: {e}")
                continue

        if not times:
            if self.debug:
                print(f"[DEBUG]   {product}: 无有效数据，跳过时序图")
            return None

        # 绘制时序图
        try:
            plt.figure(figsize=(12, 6))

            # 根据产品类型设置单位
            if product.lower() == 'sst':
                ylabel = 'Bias (K)'
            else:
                # 其他产品需要转换为百分比
                biases = [b * 100 for b in biases]
                ylabel = 'Bias (%)'

            plt.plot(times, biases, 'b-o', markersize=4, linewidth=1.5, label=f'{satellite} vs {source}')

            plt.title(f'{satellite} vs {source} {product} Time Series')
            plt.xlabel('Time')
            plt.ylabel(ylabel)
            plt.legend()
            plt.grid(True, alpha=0.3)
            plt.xticks(rotation=45)
            plt.tight_layout()

            # 保存图片
            output_filename = f'{satellite}_COCTS_{source}_{product.upper()}_TIMESERIES_{timestamp}.jpg'
            output_path = os.path.join(self.output_dir, output_filename)
            plt.savefig(output_path, bbox_inches='tight', dpi=100)
            plt.close()

            print(f"[INFO]   ✓ 生成时序图: {output_filename}")
            return output_path

        except Exception as e:
            print(f"[ERROR] 生成时序图失败: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            return None


# ============================================================================
# 模板填充器
# ============================================================================

class TemplateFiller:
    def __init__(self, template_dir):
        self.template_dir = template_dir
    
    def get_template_path(self, source):
        """获取模板路径"""
        if source.upper() in ['AQUA', 'TERRA']:
            return os.path.join(self.template_dir, 'new_auqa_terra_template.docx')
        elif source.upper() in ['SNPP', 'JPSS']:
            return os.path.join(self.template_dir, 'new_snpp_jpss_template.docx')
        elif source.upper() in ['HY1C', 'HY1D', 'HY1E']:
            return os.path.join(self.template_dir, 'satelite_template.docx')
        elif source.upper() == 'XC':
            return os.path.join(self.template_dir, 'new_xc_template.docx')
        else:
            return None
    
    def generate_replacements(self, satellite, source, products_data, period_str, timestamp, image_dir):
        """生成替换字典"""
        replacements = {
            'text': {
                '{{date}}': period_str,
                '{{date_cn}}': period_str,
                '{{satellite_type}}': satellite,
                '{{source_type}}': source,
            },
            'tables': {},
            'images': {}
        }
        
        var_config = VAR_CONFIGS.get(source.upper(), {})

        for var_name, config in var_config.items():
            if var_name in products_data:
                data = products_data[var_name]

                # 获取产品中文名称和单位
                product_cn_name = PRODUCT_NAMES.get(var_name, var_name)
                unit = config.get('unit', '')
                n = data.get('n', 0)

                # 验证结果表格（表一）- n=0时不添加，让cleanup清除占位符
                if n > 0:
                    # 根据产品类型设置单位
                    if var_name == 'sst':
                        # 海温：bias和rms都用K（绝对误差）
                        bias_str = f"{data.get('bias', 0):.4f}{unit}"
                        rms_str = f"{data.get('rms', 0):.4f}{unit}"
                    else:
                        # 其他产品：bias用%（相对误差，需要乘100），rms用原单位
                        bias_str = f"{data.get('bias', 0)*100:.2f}%"
                        rms_str = f"{data.get('rms', 0):.4f}{unit}"

                    val_results = [[
                        f'{satellite} vs {source}',
                        bias_str,
                        rms_str
                    ]]
                    replacements['tables'][f'{{{{val_results_{var_name}}}}}'] = val_results

                # 匹配结果表格（表二）- n=0时全部填0
                if n > 0:
                    space_window = data.get('space_window', '汇总')
                    time_window = data.get('time_window', '汇总')
                    # 为时间窗口添加单位（如果不是"汇总"）
                    if time_window != '汇总' and not time_window.endswith('h'):
                        time_window = f"{time_window}h"
                else:
                    # n=0时：产品名、空间窗口0、时间窗口0、匹配数0
                    space_window = "0"
                    time_window = "0"

                col_results = [[
                    product_cn_name,
                    space_window,
                    time_window,
                    f"{n}"
                ]]
                replacements['tables'][f'{{{{col_results_{var_name}}}}}'] = col_results

                # 只有n>0时才添加图片路径
                if n > 0:
                    # 图片路径 - 占位符固定为terra
                    pie_key = f'hy1c_vs_{var_name}_terra_sct'
                    geo_key = f'hy1c_vs_{var_name}_terra_geo'
                    timeseries_key = f'hy1c_vs_{var_name}_terra_timeseries'

                    pie_path = os.path.join(image_dir, f'{satellite}_COCTS_{source}_{var_name}_PIE_{timestamp}.jpg')
                    geo_path = os.path.join(image_dir, f'{satellite}_COCTS_{source}_{var_name}_GEO_{timestamp}.jpg')
                    timeseries_path = os.path.join(image_dir, f'{satellite}_COCTS_{source}_{var_name.upper()}_TIMESERIES_{timestamp}.jpg')

                    replacements['images'][f'{{{{{pie_key}}}}}'] = pie_path
                    replacements['images'][f'{{{{{geo_key}}}}}'] = geo_path
                    replacements['images'][f'{{{{{timeseries_key}}}}}'] = timeseries_path

                replacements['text']['{{unit}}'] = config.get('unit', '')
        
        return replacements
    
    def fill_template(self, template_path, output_docx, replacements):
        """填充模板"""
        doc = Document(template_path)
        
        if 'text' in replacements:
            for placeholder, text in replacements['text'].items():
                self._replace_text(doc, placeholder, text)
        
        if 'tables' in replacements:
            for placeholder, table_data in replacements['tables'].items():
                self._fill_table(doc, placeholder, table_data)
        
        if 'images' in replacements:
            for placeholder, image_path in replacements['images'].items():
                self._insert_image(doc, placeholder, image_path)

        # 处理模板中未使用的占位符
        self._handle_unused_placeholders(doc, replacements)

        # XC替换为现场
        self._replace_text(doc, 'XC卫星', '现场')
        self._replace_text(doc, 'XC', '现场')

        # 先删除所有章节中没有图片的小节并重新编号（必须在清理占位符之前执行）
        self._remove_empty_subsections_and_renumber(doc)

        # 问题7: 清理所有未替换的占位符（移除大括号形式的参数名称）
        self._cleanup_placeholders(doc)

        doc.save(output_docx)
        print(f"[INFO] ✓ 保存文档: {os.path.basename(output_docx)}")
    
    def _replace_text(self, doc, placeholder, replacement):
        for p in doc.paragraphs:
            if placeholder in p.text:
                self._replace_text_in_paragraph(p, placeholder, replacement)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if placeholder in p.text:
                            self._replace_text_in_paragraph(p, placeholder, replacement)
    
    def _replace_text_in_paragraph(self, paragraph, placeholder, replacement):
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, replacement)
    
    def _fill_table(self, doc, placeholder, table_data):
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if placeholder in cell.text:
                        for c in row.cells:
                            c.text = ""
                        # 确保不超出表格列数
                        num_cells = len(row.cells)
                        for i, value in enumerate(table_data[0]):
                            if i >= num_cells:
                                break  # 防止索引越界
                            cell = row.cells[i]
                            cell.text = str(value)
                            for paragraph in cell.paragraphs:
                                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        return
    
    def _insert_image(self, doc, placeholder, image_path):
        if not os.path.exists(image_path):
            # 图片不存在时清除占位符
            for p in doc.paragraphs:
                if placeholder in p.text:
                    p.text = p.text.replace(placeholder, '')
                    return
            
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if placeholder in cell.text:
                            cell.text = cell.text.replace(placeholder, '')
                            return
            return
        
        for p in doc.paragraphs:
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, '')
                run = p.add_run()
                run.add_picture(image_path, width=Inches(6))
                return
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, '')
                        run = cell.add_run()
                        run.add_picture(image_path, width=Inches(4))
                        return

    def _handle_unused_placeholders(self, doc, replacements):
        """
        处理模板中存在但未使用的占位符
        - 表一（val_results）：填写检验类型和"/"
        - 表二（col_results）：删除整行
        - 第三章图片占位符：由_remove_empty_subsections_and_renumber处理
        """
        import re

        print("\n" + "="*80)
        print("【调试】开始处理未使用的占位符")
        print("="*80)

        # 扫描所有占位符
        all_placeholders = set()
        placeholder_pattern = re.compile(r'\{\{([^}]+)\}\}')

        # 从所有段落中提取占位符（记录章节信息）
        chapter_placeholders = {}  # {placeholder: [章节列表]}
        chapter_pattern = re.compile(r'^(\d+)[\.\s]')  # 匹配章节号

        current_chapter = None
        for i, p in enumerate(doc.paragraphs):
            text = p.text.strip()

            # 检测章节标题
            chapter_match = chapter_pattern.match(text)
            if chapter_match:
                current_chapter = int(chapter_match.group(1))

            # 提取占位符
            matches = placeholder_pattern.findall(p.text)
            for match in matches:
                all_placeholders.add(match)
                if current_chapter and current_chapter >= 3:  # 只记录第3章及以后
                    if match not in chapter_placeholders:
                        chapter_placeholders[match] = []
                    if current_chapter not in chapter_placeholders[match]:
                        chapter_placeholders[match].append(current_chapter)

        # 从所有表格中提取占位符
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        matches = placeholder_pattern.findall(p.text)
                        all_placeholders.update(matches)

        print(f"\n【调试】模板中发现的所有占位符总数: {len(all_placeholders)}")

        # 按类型分类占位符
        image_placeholders = [p for p in all_placeholders if any(x in p for x in ['_sct', '_geo', '_map', '_chart'])]
        val_placeholders = [p for p in all_placeholders if p.startswith('val_results_')]
        col_placeholders = [p for p in all_placeholders if p.startswith('col_results_')]
        other_placeholders = [p for p in all_placeholders if p not in image_placeholders and p not in val_placeholders and p not in col_placeholders]

        print(f"  - 图片占位符: {len(image_placeholders)}")
        print(f"  - 表一占位符 (val_results): {len(val_placeholders)}")
        print(f"  - 表二占位符 (col_results): {len(col_placeholders)}")
        print(f"  - 其他占位符: {len(other_placeholders)}")

        # 输出第3章及以后的占位符
        print(f"\n【调试】第3章及以后的占位符分布:")
        for placeholder, chapters in sorted(chapter_placeholders.items()):
            chapters_str = ', '.join([f"第{ch}章" for ch in sorted(chapters)])
            print(f"  {{{{{{placeholder}}}}}}: {chapters_str}")

        # 识别val_results和col_results占位符
        val_pattern = re.compile(r'val_results_(\w+)')
        col_pattern = re.compile(r'col_results_(\w+)')

        used_tables = set(replacements.get('tables', {}).keys())
        used_images = set(replacements.get('images', {}).keys())

        print(f"\n【调试】已使用的占位符:")
        print(f"  - 表格占位符: {len(used_tables)}")
        for placeholder in sorted(used_tables):
            print(f"      {placeholder}")
        print(f"  - 图片占位符: {len(used_images)}")
        for placeholder in sorted(used_images):
            print(f"      {placeholder}")

        # 找出未使用的占位符
        all_placeholders_with_braces = {f'{{{{{p}}}}}' for p in all_placeholders}
        unused_placeholders = all_placeholders_with_braces - used_tables - used_images

        print(f"\n【调试】未使用的占位符总数: {len(unused_placeholders)}")
        unused_images = [p for p in unused_placeholders if any(x in p for x in ['_sct', '_geo', '_map', '_chart'])]
        unused_vals = [p for p in unused_placeholders if 'val_results_' in p]
        unused_cols = [p for p in unused_placeholders if 'col_results_' in p]

        print(f"  - 未使用的图片占位符: {len(unused_images)}")
        for p in sorted(unused_images):
            print(f"      {p}")
        print(f"  - 未使用的表一占位符: {len(unused_vals)}")
        for p in sorted(unused_vals):
            print(f"      {p}")
        print(f"  - 未使用的表二占位符: {len(unused_cols)}")
        for p in sorted(unused_cols):
            print(f"      {p}")

        # 从replacements中提取卫星和数据源信息
        satellite = replacements.get('text', {}).get('{{satellite_type}}', 'HY1C')
        source = replacements.get('text', {}).get('{{source_type}}', 'AQUA')

        # 处理未使用的val_results占位符（表一）
        for placeholder in all_placeholders:
            match = val_pattern.match(placeholder)
            if match:
                full_placeholder = f'{{{{{placeholder}}}}}'
                if full_placeholder not in used_tables:
                    # 填写检验类型和"/"
                    var_name = match.group(1)
                    print(f"\n【调试】处理未使用的val_results: {full_placeholder}")
                    self._fill_unused_val_results(doc, var_name, satellite, source)

        # 处理未使用的col_results占位符（表二）
        for placeholder in all_placeholders:
            match = col_pattern.match(placeholder)
            if match:
                full_placeholder = f'{{{{{placeholder}}}}}'
                if full_placeholder not in used_tables:
                    # 删除整行
                    var_name = match.group(1)
                    print(f"【调试】删除未使用的col_results行: {full_placeholder}")
                    self._delete_col_results_row(doc, var_name)

        print("\n" + "="*80)
        print("【调试】未使用占位符处理完成")
        print("="*80 + "\n")

    def _fill_unused_val_results(self, doc, var_name, satellite, source):
        """为未使用的val_results占位符填写检验类型和"/" """
        placeholder = f'{{{{val_results_{var_name}}}}}'

        # 使用与其他产品相同的检验类型格式
        if source.upper() == 'XC':
            validation_type = f"{satellite} vs 现场"
        else:
            validation_type = f"{satellite} vs {source}"

        # 填充表格：检验类型 | / | /
        val_results = [[validation_type, '/', '/']]
        self._fill_table(doc, placeholder, val_results)

    def _delete_col_results_row(self, doc, var_name):
        """删除未使用的col_results占位符所在的整行"""
        placeholder = f'{{{{col_results_{var_name}}}}}'

        # 遍历所有表格，找到包含此占位符的行并删除
        for table in doc.tables:
            rows_to_delete = []
            for i, row in enumerate(table.rows):
                for cell in row.cells:
                    if placeholder in cell.text:
                        rows_to_delete.append(i)
                        break

            # 从后往前删除行，避免索引变化
            for row_idx in sorted(rows_to_delete, reverse=True):
                table._element.remove(table.rows[row_idx]._element)

    def _cleanup_placeholders(self, doc):
        """清理所有未替换的占位符（问题7: 移除大括号形式的参数名称）"""
        import re
        placeholder_pattern = re.compile(r'\{\{[^}]+\}\}')

        # 清理段落中的占位符
        for p in doc.paragraphs:
            # 检查段落的整体文本（处理占位符跨多个run的情况）
            if placeholder_pattern.search(p.text):
                cleaned_text = placeholder_pattern.sub('', p.text)
                # 清空所有run并设置新文本
                for run in p.runs:
                    run.text = ''
                if cleaned_text.strip():  # 如果清理后还有文本
                    p.runs[0].text = cleaned_text
                elif p.runs:  # 如果清理后没有文本，保持空段落
                    p.runs[0].text = cleaned_text

        # 清理表格中的占位符
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        # 检查段落的整体文本
                        if placeholder_pattern.search(p.text):
                            cleaned_text = placeholder_pattern.sub('', p.text)
                            # 清空所有run并设置新文本
                            for run in p.runs:
                                run.text = ''
                            if p.runs:
                                p.runs[0].text = cleaned_text

    def _remove_empty_subsections_and_renumber(self, doc):
        """
        删除所有没有图片和实质内容的小节，并重新编号
        识别规则：小节标题格式为"X.Y"（如"3.1"、"8.2"、"10.1"等）
        """
        import re
        from collections import defaultdict

        print("\n" + "="*80)
        print("【调试】开始删除空小节并重新编号")
        print("="*80)

        # 第一步：找出需要删除的段落范围
        paragraphs_to_delete_objs = []  # 存储段落对象而不是索引
        subsection_info = []  # 记录所有小节信息：(para_idx, para_obj, chapter, subsection, text)

        # 识别所有小节标题（X.Y格式，如3.1、8.2、10.1等）
        subsection_pattern = re.compile(r'^(\d+)\.(\d+)')

        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            # 检查是否是小节标题
            match = subsection_pattern.match(text)
            if match:
                chapter = int(match.group(1))
                subsection = int(match.group(2))
                subsection_info.append((i, para, chapter, subsection, text))

        print(f"\n【调试】找到 {len(subsection_info)} 个小节标题")
        for para_idx, para_obj, chapter, subsection, text in subsection_info:
            print(f"  {chapter}.{subsection}: {text[:50]}...")

        # 第二步：检查每个小节是否包含内容（非空段落或有实际内容）
        empty_subsections = []
        print(f"\n【调试】开始检查每个小节的内容...")
        for idx in range(len(subsection_info)):
            start_idx = subsection_info[idx][0]
            start_para_obj = subsection_info[idx][1]
            chapter = subsection_info[idx][2]
            subsection = subsection_info[idx][3]
            title = subsection_info[idx][4]

            # 确定小节结束位置（下一个小节开始前，或文档结束）
            end_idx = subsection_info[idx + 1][0] if idx + 1 < len(subsection_info) else len(doc.paragraphs)

            # 检查这个小节是否为空（只有标题，没有其他有意义内容）
            has_content = False
            has_image = False
            has_placeholder = False
            content_details = []

            for para_idx in range(start_idx + 1, end_idx):
                if para_idx < len(doc.paragraphs):
                    para = doc.paragraphs[para_idx]
                    text = para.text.strip()

                    # 先检查是否包含图片（即使文本为空也要检查，因为插入图片后文本会被清空）
                    for run in para.runs:
                        if hasattr(run, '_element'):
                            drawings = run._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
                            if drawings:
                                has_image = True
                                content_details.append(f"图片")
                                break

                    # 如果有图片，则认为有内容
                    if has_image:
                        has_content = True
                        break

                    # 跳过空白、小节标题和章节标题
                    if not text or subsection_pattern.match(text):
                        continue

                    # 跳过章节标题（格式：数字 + 空格 + 文字，如"8 HY1C水色水温扫描仪..."）
                    chapter_title_pattern = re.compile(r'^\d+\s+[^\d]')
                    if chapter_title_pattern.match(text):
                        content_details.append(f"章节标题: {text[:40]}...")
                        continue

                    # 检查是否只是占位符（包含{{}}的文本）
                    if '{{' in text:
                        has_placeholder = True
                        content_details.append(f"占位符: {text[:40]}...")
                        continue  # 跳过占位符，继续检查下一段

                    # 检查是否只是图片标题（以"图"开头，包含关键词）
                    if text.startswith('图') and any(word in text for word in ['HY1C', 'vs', '检验', '分布', 'TERRA', 'AQUA']):
                        # 这是图片标题，但没有实际图片，不算实质内容
                        content_details.append(f"图片标题: {text[:40]}...")
                        continue  # 跳过图片标题，继续检查下一段

                    # 其他有实质内容的文本
                    if len(text) > 5:
                        has_content = True
                        content_details.append(f"文本: {text[:40]}...")
                        break

            # 输出每个小节的检查结果（用于调试）
            print(f"\n【调试】小节 {chapter}.{subsection}:")
            print(f"  标题: {title[:60]}...")
            print(f"  有内容: {has_content}, 有图片: {has_image}, 有占位符: {has_placeholder}")
            print(f"  内容详情: {content_details if content_details else '无'}")

            # 如果小节为空，收集该小节的段落（不包括下一小节/章节的标题）
            if not has_content:
                empty_subsections.append((chapter, subsection, title))
                print(f"  >>> 标记为删除")

                # 收集要删除的段落对象（只收集当前小节的内容）
                chapter_title_pattern_local = re.compile(r'^\d+\s+[^\d]')
                for para_idx in range(start_idx, end_idx):
                    if para_idx < len(doc.paragraphs):
                        para_obj = doc.paragraphs[para_idx]
                        # 跳过章节标题（会在第四步单独处理）
                        para_text = para_obj.text.strip()
                        if para_idx > start_idx and chapter_title_pattern_local.match(para_text):
                            # 这是下一个章节的标题，不应包含在当前小节的删除范围内
                            break
                        if para_obj not in paragraphs_to_delete_objs:
                            paragraphs_to_delete_objs.append(para_obj)

        print(f"\n【调试】共标记 {len(empty_subsections)} 个小节待删除:")
        for chapter, subsection, title in empty_subsections:
            print(f"  {chapter}.{subsection}: {title[:50]}...")

        # 第三步：识别空章节（所有小节都被删除的章节）
        deleted_chapters = set()
        for chapter, subsection, title in empty_subsections:
            deleted_chapters.add(chapter)

        # 检查每个章节是否所有小节都被删除
        chapters_to_delete = set()
        for chapter in deleted_chapters:
            # 获取该章节的所有小节
            chapter_subsections = [s for s in subsection_info if s[2] == chapter]
            # 检查是否所有小节都被标记为删除
            all_deleted = all((chapter, s[3], s[4]) in empty_subsections for s in chapter_subsections)
            if all_deleted:
                chapters_to_delete.add(chapter)

        print(f"\n【调试】需要删除的空章节: {sorted(chapters_to_delete)}")

        # 第四步：收集章节标题对象
        chapter_title_pattern = re.compile(r'^(\d+)\s+[^\d]')
        for para in doc.paragraphs:
            text = para.text.strip()
            match = chapter_title_pattern.match(text)
            if match:
                chapter_num = int(match.group(1))
                if chapter_num in chapters_to_delete:
                    print(f"【调试】标记删除章节标题: {text[:60]}...")
                    if para not in paragraphs_to_delete_objs:
                        paragraphs_to_delete_objs.append(para)

        # 第五步：统一删除所有标记的段落（使用对象而不是索引）
        print(f"\n【调试】开始删除 {len(paragraphs_to_delete_objs)} 个段落...")
        for para_obj in paragraphs_to_delete_objs:
            try:
                p_element = para_obj._element
                p_element.getparent().remove(p_element)
            except Exception as e:
                print(f"【警告】删除段落时出错: {e}")
                continue

        print(f"【调试】删除完成")

        # 第六步：重新编号章节和小节
        print(f"\n【调试】开始重新编号章节和小节...")

        # 建立章节映射：旧章节号 -> 新章节号
        remaining_chapters = []
        for para in doc.paragraphs:
            text = para.text.strip()
            match = chapter_title_pattern.match(text)
            if match:
                chapter_num = int(match.group(1))
                if chapter_num >= 3 and chapter_num not in remaining_chapters:
                    remaining_chapters.append(chapter_num)

        chapter_mapping = {}
        for new_num, old_num in enumerate(sorted(remaining_chapters), start=3):
            chapter_mapping[old_num] = new_num

        print(f"【调试】章节映射: {chapter_mapping}")

        # 重新编号章节标题
        for para in doc.paragraphs:
            text = para.text.strip()
            match = chapter_title_pattern.match(text)
            if match:
                old_chapter = int(match.group(1))
                if old_chapter in chapter_mapping:
                    new_chapter = chapter_mapping[old_chapter]
                    if old_chapter != new_chapter:
                        old_text = text
                        new_text = re.sub(r'^\d+', str(new_chapter), text)
                        para.text = new_text
                        print(f"  重新编号章节: {old_text[:40]}... -> {new_text[:40]}...")

        # 重新编号小节标题
        chapter_counters = defaultdict(int)
        for para in doc.paragraphs:
            text = para.text.strip()
            match = subsection_pattern.match(text)
            if match:
                old_chapter = int(match.group(1))
                if old_chapter in chapter_mapping:
                    new_chapter = chapter_mapping[old_chapter]
                    chapter_counters[new_chapter] += 1
                    old_text = text
                    new_text = re.sub(r'^\d+\.\d+', f'{new_chapter}.{chapter_counters[new_chapter]}', text)
                    para.text = new_text
                    if old_text != new_text:
                        print(f"  重新编号小节: {old_text[:40]}... -> {new_text[:40]}...")

        # 第七步：重新编号图片标题
        print(f"\n【调试】开始重新编号图片...")
        figure_counter = 1
        figure_pattern = re.compile(r'^图\s*\d+')
        for para in doc.paragraphs:
            text = para.text.strip()
            if figure_pattern.match(text):
                old_text = text
                new_text = re.sub(r'^图\s*\d+', f'图{figure_counter}', text)
                para.text = new_text
                if old_text != new_text:
                    print(f"  重新编号图片: {old_text[:50]}... -> {new_text[:50]}...")
                figure_counter += 1

        print("\n" + "="*80)
        print("【调试】空小节删除和重新编号完成")
        print("="*80 + "\n")


# ============================================================================
# PDF转换函数
# ============================================================================

def word_to_pdf(input_dir, output_dir, specific_files=None):
    """
    将指定目录中的 .docx 文件转换为 PDF 文件。

    参数:
        input_dir (str): 包含 .docx 文件的输入目录。
        output_dir (str): 保存转换后的 PDF 文件的输出目录。
        specific_files (list): 可选,指定要转换的docx文件名列表。如果为None,则转换所有docx文件。
    """
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 确定要处理的文件列表
    if specific_files is not None:
        # 只处理指定的文件
        files_to_convert = [f for f in specific_files if f.endswith(".docx")]
        print(f"\n[INFO] 仅转换本次生成的 {len(files_to_convert)} 个docx文件")
    else:
        # 处理目录中所有的docx文件(向后兼容)
        files_to_convert = [f for f in os.listdir(input_dir) if f.endswith(".docx")]
        print(f"\n[INFO] 转换目录中所有 {len(files_to_convert)} 个docx文件")

    # 遍历要转换的文件
    for filename in files_to_convert:
        # 构建输入文件的完整路径
        input_file = os.path.join(input_dir, filename)

        # 检查文件是否存在
        if not os.path.exists(input_file):
            print(f"[WARNING] 文件不存在,跳过: {input_file}")
            continue

        # 构建输出文件的完整路径（将 .docx 替换为 .pdf）
        output_file = os.path.join(output_dir, filename.replace(".docx", ".pdf"))

        # 构建 LibreOffice 命令
        cmd = [
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            input_file,
            '--outdir', output_dir
        ]

        # 运行命令
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        # 检查转换是否成功
        if result.returncode == 0:
            print(f"[INFO] ✓ 成功转换：{filename} -> {filename.replace('.docx', '.pdf')}")
        else:
            print(f"[ERROR] 转换失败：{filename}")
            print(f"[ERROR] 错误信息：{result.stderr.decode()}")


# ============================================================================
# 主生成器
# ============================================================================

class MonthlyReportGenerator:
    def __init__(self, report_type, year, month=None, quarter=None, output_dir=None, debug=False):
        self.report_type = report_type
        self.year = year
        self.month = month
        self.quarter = quarter
        self.debug = debug
        
        self.start_date, self.end_date = self._calculate_date_range()
        
        self.config = self._load_config()
        self.font_manager = FontManager(self.config)
        
        # 读取空间窗口和时间窗口
        self.space_size = int(self.config.get('PARAMS', 'window_size', fallback='25'))
        self.time_size = int(self.config.get('PARAMS', 'time_threshold', fallback='1800'))
        
        if output_dir:
            self.output_base = output_dir
        else:
            self.output_base = self.config.get('PATH', 'output_dir', fallback='output')
        
        self.output_base = self.output_base.replace('\\', '/')

        self.summary_dir = os.path.join(self.output_base, '05_reports')
        self.daily_reports_dir = os.path.join(self.output_base, '05_reports')
        self.collocation_dir = os.path.join(self.output_base, '03_collocation')
        self.image_dir = os.path.join(self.output_base, '04_visualization')
        
        # 卫星数据目录 (lat/lon文件)
        input_dir = self.config.get('PATH', 'input_dir', fallback='input')
        self.sat_input_dir = os.path.join(input_dir.replace('\\', '/'), '01_sat')
        
        # 模板目录
        self.template_dir = os.path.join(input_dir.replace('\\', '/'), '05_reports')
        
        os.makedirs(self.summary_dir, exist_ok=True)
        os.makedirs(self.image_dir, exist_ok=True)
        
        # 初始化组件
        self.data_aggregator = DataAggregator(self.daily_reports_dir, self.space_size, self.time_size, debug)
        self.chart_generator = ChartGenerator(self.image_dir, self.collocation_dir, 
                                              self.sat_input_dir, self.font_manager, debug)
        self.template_filler = TemplateFiller(self.template_dir)
        
        print(f"[INFO] {'='*80}")
        print(f"[INFO] 月季年报告生成器 V5")
        print(f"[INFO] {'='*80}")
        print(f"[INFO] 报告: {self._get_report_title()}")
        print(f"[INFO] 时间: {self.start_date.strftime('%Y-%m-%d')} ~ {self.end_date.strftime('%Y-%m-%d')}")
        print(f"[INFO] 空间窗口: {self.space_size}x{self.space_size}")
        print(f"[INFO] 时间窗口: {self.time_size}秒")
        print(f"[INFO] 模板: {self.template_dir}")
        print(f"[INFO] 卫星数据: {self.sat_input_dir}")
    
    def _calculate_date_range(self):
        if self.report_type == 'month':
            _, last_day = monthrange(self.year, self.month)
            start = datetime(self.year, self.month, 1)
            end = datetime(self.year, self.month, last_day)
        elif self.report_type == 'quarter':
            start_month = (self.quarter - 1) * 3 + 1
            end_month = self.quarter * 3
            start = datetime(self.year, start_month, 1)
            _, last_day = monthrange(self.year, end_month)
            end = datetime(self.year, end_month, last_day)
        else:
            start = datetime(self.year, 1, 1)
            end = datetime(self.year, 12, 31)
        return start, end
    
    def _get_report_title(self):
        if self.report_type == 'month':
            return f"{self.year}年{self.month:02d}月汇总报告"
        elif self.report_type == 'quarter':
            return f"{self.year}年第{self.quarter}季度汇总报告"
        else:
            return f"{self.year}年汇总报告"
    
    def _get_report_suffix(self):
        if self.report_type == 'month':
            return f"{self.year}{self.month:02d}"
        elif self.report_type == 'quarter':
            return f"{self.year}S{self.quarter}"
        else:
            return f"{self.year}"
    
    def _load_config(self):
        config = configparser.ConfigParser()
        if os.path.exists('config.ini'):
            config.read('config.ini', encoding='utf-8')
        return config
    
    def scan_reports(self):
        print(f"\n[INFO] {'─'*80}")
        print(f"[INFO] 步骤1: 扫描报告")
        print(f"[INFO] {'─'*80}")
        
        pattern = os.path.join(self.daily_reports_dir, '*_val_report_*.pdf')
        all_reports = glob.glob(pattern)
        
        print(f"[INFO] 找到 {len(all_reports)} 个报告")
        
        classified = {
            'by_date': defaultdict(list),
            'by_validation': defaultdict(lambda: defaultdict(list)),
            'summary': defaultdict(int)
        }
        
        for report_path in all_reports:
            filename = os.path.basename(report_path)
            info = ReportClassifier.parse_report_filename(filename)
            
            if not info:
                continue
            
            try:
                report_date = datetime.strptime(info['date'], '%Y%m%d')
                
                if not (self.start_date <= report_date <= self.end_date):
                    continue
                
                val_label = ReportClassifier.get_validation_label(info['satellite'], info['source'])
                
                report_info = {
                    'path': report_path,
                    'filename': filename,
                    'satellite': info['satellite'],
                    'source': info['source'],
                    'date': info['date'],
                    'time': info['time'],
                    'validation_label': val_label
                }
                
                classified['by_date'][info['date']].append(report_info)
                classified['by_validation'][val_label][info['date']].append(report_info)
                classified['summary'][val_label] += 1
            except:
                continue
        
        print(f"\n[INFO] 报告分类:")
        for val_label, count in sorted(classified['summary'].items()):
            print(f"[INFO]   {val_label}: {count} 个")
        
        return classified
    
    def run(self):
        try:
            # 1. 扫描报告
            classified = self.scan_reports()
            if not classified['summary']:
                print("[ERROR] 未找到报告")
                return False
            
            # 2. 聚合数据
            print(f"\n[INFO] {'─'*80}")
            print(f"[INFO] 步骤2: 聚合数据")
            print(f"[INFO] {'─'*80}")
            
            aggregated = self.data_aggregator.aggregate_validation_data(classified)
            
            # 3. 生成报告
            print(f"\n[INFO] {'─'*80}")
            print(f"[INFO] 步骤3: 生成图表和报告")
            print(f"[INFO] {'─'*80}")
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            suffix = self._get_report_suffix()
            period_str = self._get_report_title().replace('汇总报告', '')

            # 跟踪生成的docx文件列表，用于PDF转换
            generated_docx_files = []

            for val_label, val_data in aggregated.items():
                print(f"\n[INFO] 生成报告: {val_label}")
                
                satellite = val_data['satellite']
                source = val_data['source']
                products_data = val_data['products']
                
                # 生成图表
                for product, data in products_data.items():
                    if self.debug:
                        print(f"\n[DEBUG] ========== 处理产品: {product} ==========")
                    
                    # 获取matchup文件（用于饼图）
                    matchup_files = []
                    
                    for date_str in val_data['dates']:
                        patterns = [
                            f'{satellite}_COCTS_{source}_{product}_matchup_{date_str}_*.txt',
                            f'{satellite}_COCTS_{source}_{product}_matchup_{date_str}.txt'
                        ]
                        for pattern in patterns:
                            full_pattern = os.path.join(self.collocation_dir, pattern)
                            found = glob.glob(full_pattern)
                            matchup_files.extend(found)
                            if self.debug and found:
                                print(f"[DEBUG] 找到matchup: {[os.path.basename(f) for f in found]}")
                    
                    matchup_files = list(set(matchup_files))
                    
                    if self.debug:
                        print(f"[DEBUG] 总计matchup: {len(matchup_files)}")
                    
                    if matchup_files:
                        # 生成饼图
                        self.chart_generator.generate_pie_chart(
                            satellite, source, product, matchup_files, timestamp
                        )
                    
                    # 生成地理图 - 传递日期列表
                    self.chart_generator.generate_geo_map(
                        satellite, source, product, val_data['dates'], timestamp
                    )

                    # 生成时序图 - 收集report文件
                    report_files = []
                    for date_str in val_data['dates']:
                        patterns = [
                            f'report_{satellite}_{source}_{product}_{date_str}_*.txt',
                            f'{satellite}_COCTS_{source}_{product}_report_{date_str}_*.txt'
                        ]
                        for pattern in patterns:
                            full_pattern = os.path.join(self.daily_reports_dir, pattern)
                            found = glob.glob(full_pattern)
                            report_files.extend(found)

                    report_files = list(set(report_files))
                    if report_files:
                        self.chart_generator.generate_time_series_chart(
                            satellite, source, product, report_files, timestamp
                        )

                # 填充模板
                template_path = self.template_filler.get_template_path(source)
                if not template_path or not os.path.exists(template_path):
                    print(f"[WARNING] 模板不存在: {template_path}")
                    continue
                
                replacements = self.template_filler.generate_replacements(
                    satellite, source, products_data, period_str, timestamp, self.image_dir
                )
                
                # 问题8: 调整文件名格式，使用大写以匹配日报格式
                output_docx = os.path.join(
                    self.summary_dir,
                    f'{satellite.upper()}_COCTS_{source.upper()}_val_report_{suffix}.docx'
                )

                self.template_filler.fill_template(template_path, output_docx, replacements)

                # 记录生成的docx文件
                generated_docx_files.append(os.path.basename(output_docx))

            # 4. 转换为PDF
            if generated_docx_files:
                print(f"\n[INFO] {'─'*80}")
                print(f"[INFO] 步骤4: 转换为PDF")
                print(f"[INFO] {'─'*80}")
                word_to_pdf(self.summary_dir, self.summary_dir, generated_docx_files)

            print(f"\n[INFO] {'='*80}")
            print(f"[INFO] ✓ 完成")
            print(f"[INFO] {'='*80}")
            print(f"[INFO] 输出: {self.summary_dir}")
            
            return True
            
        except Exception as e:
            print(f"[ERROR] 失败: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            return False


def main():
    parser = argparse.ArgumentParser(description='月季年报告生成器 V5 - 从map文件生成GEO图')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--month', metavar='YYYYMM')
    group.add_argument('--quarter', nargs=2, metavar=('YYYY', 'Q'))
    group.add_argument('--year', metavar='YYYY', type=int)
    parser.add_argument('--output')
    parser.add_argument('--debug', action='store_true')
    
    args = parser.parse_args()
    
    try:
        if args.month:
            year = int(args.month[:4])
            month = int(args.month[4:6])
            generator = MonthlyReportGenerator('month', year, month=month, 
                                              output_dir=args.output, debug=args.debug)
        elif args.quarter:
            year = int(args.quarter[0])
            quarter = int(args.quarter[1])
            generator = MonthlyReportGenerator('quarter', year, quarter=quarter,
                                              output_dir=args.output, debug=args.debug)
        else:
            generator = MonthlyReportGenerator('year', args.year,
                                              output_dir=args.output, debug=args.debug)
        
        return 0 if generator.run() else 1
    except Exception as e:
        print(f"错误: {e}")
        return 1


if __name__ == '__main__':
    exit(main())
