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
        'sst': {'sources': ['AQUA'], 'unit': '℃'},
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
        'sst': {'sources': ['TERRA'], 'unit': '℃'},
        'chl': {'sources': ['TERRA'], 'unit': 'mg/m³'},
        'Rrs412': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'Rrs443': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'Rrs490': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'Rrs520': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'Rrs565': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'Rrs670': {'sources': ['TERRA'], 'unit': 'sr⁻¹'},
        'AOT': {'sources': ['TERRA'], 'unit': ''},
    },
    'XC': {
        'sst': {'sources': ['XC'], 'unit': '℃'},
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
    'Rrs670': '670nm遥感反射率'
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

                # 验证结果表格 - 问题1&2: 添加单位，增加小数位数；问题6: n=0时填"/"
                if n > 0:
                    val_results = [[
                        f'{satellite} vs {source}',
                        f"{data.get('bias', 0):.4f}{unit}",  # 增加到4位小数，添加单位
                        f"{data.get('rms', 0):.4f}{unit}"    # 增加到4位小数，添加单位
                    ]]
                else:
                    # 无匹配数据时填"/"
                    val_results = [[
                        f'{satellite} vs {source}',
                        '/',
                        '/'
                    ]]

                # 匹配结果表格 - 问题3: 使用产品中文名；问题4: 添加时间窗口单位
                space_window = data.get('space_window', '汇总')
                time_window = data.get('time_window', '汇总')
                # 为时间窗口添加单位（如果不是"汇总"）
                if time_window != '汇总' and not time_window.endswith('h'):
                    time_window = f"{time_window}h"

                col_results = [[
                    product_cn_name,  # 问题3: 改为产品中文名称
                    space_window,
                    time_window,
                    f"{n}"  # 问题6: 即使n=0也显示0
                ]]

                replacements['tables'][f'{{{{val_results_{var_name}}}}}'] = val_results
                replacements['tables'][f'{{{{col_results_{var_name}}}}}'] = col_results

                # 问题6: 只有n>0时才添加图片路径（第3章才显示该产品）
                if n > 0:
                    # 图片路径 - 占位符固定为terra
                    pie_key = f'hy1c_vs_{var_name}_terra_sct'
                    geo_key = f'hy1c_vs_{var_name}_terra_geo'

                    pie_path = os.path.join(image_dir, f'{satellite}_COCTS_{source}_{var_name}_PIE_{timestamp}.jpg')
                    geo_path = os.path.join(image_dir, f'{satellite}_COCTS_{source}_{var_name}_GEO_{timestamp}.jpg')

                    replacements['images'][f'{{{{{pie_key}}}}}'] = pie_path
                    replacements['images'][f'{{{{{geo_key}}}}}'] = geo_path

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
        
        # XC替换为现场
        self._replace_text(doc, 'XC卫星', '现场')
        self._replace_text(doc, 'XC', '现场')

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
                        for i, value in enumerate(table_data[0]):
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

    def _cleanup_placeholders(self, doc):
        """清理所有未替换的占位符（问题7: 移除大括号形式的参数名称）"""
        import re
        placeholder_pattern = re.compile(r'\{\{[^}]+\}\}')

        # 清理段落中的占位符
        for p in doc.paragraphs:
            for run in p.runs:
                if placeholder_pattern.search(run.text):
                    run.text = placeholder_pattern.sub('', run.text)

        # 清理表格中的占位符
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            if placeholder_pattern.search(run.text):
                                run.text = placeholder_pattern.sub('', run.text)


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
        
        self.summary_dir = os.path.join(self.output_base, '06_summary_reports', f'{report_type}ly')
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
            return f"{self.year}Q{self.quarter}"
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
                    f'{satellite.upper()}_COCTS_{source.upper()}_summary_{suffix}.docx'
                )
                
                self.template_filler.fill_template(template_path, output_docx, replacements)
            
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
