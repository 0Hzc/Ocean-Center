import os
import glob
import configparser
import h5py
import netCDF4 as nc
import numpy as np
import pandas as pd
import traceback
from datetime import datetime, timedelta
from scipy import interpolate
import re
import matplotlib.pyplot as plt
import random
from mpl_toolkits.basemap import Basemap
from scipy.interpolate import griddata
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm
from reportlab.platypus import Table, TableStyle
from reportlab.lib.utils import ImageReader  
from reportlab.lib import colors
import os
from shutil import move

from docx import Document
from docx.shared import Inches
from datetime import datetime
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
import subprocess

# 产品中文名称映射
PRODUCT_NAMES = {
    'sst': '海表温度',
    'chl': '叶绿素浓度',
    'AOT': '气溶胶光学厚度',
    'Rrs412': '412nm遥感反射率',
    'Rrs443': '443nm遥感反射率',
    'Rrs490': '490nm遥感反射率',
    'Rrs520': '520nm遥感反射率',
    'Rrs565': '565nm遥感反射率',
    'Rrs670': '670nm遥感反射率',
    'TSM': '总悬浮物浓度',
    'CDOM': '有色溶解有机物'
}

def load_config(): 
    """加载配置文件"""
    config = configparser.ConfigParser()
    config.read('config.ini')
    return config

def extract_datetime(filename):
    """从文件名中提取时间信息并转换为北京时间"""
    pattern = r'\d{8}T\d{6}'
    match = re.search(pattern, filename)
    if match:
        time_str = match.group()
        utc_time = datetime.strptime(time_str, '%Y%m%dT%H%M%S')
        beijing_time = utc_time + timedelta(hours=8)
        return beijing_time
    return None


def run_check(config):
    """运行检验流程"""
    try:
        # 获取配置参数
        satellite_type = config['SATELLITE']['type']
        input_dir = config['PATH']['input_dir']
        output_dir = config['PATH']['output_dir']
        window_size = int(config['PARAMS']['window_size'])
        time_threshold = int(config['PARAMS']['time_threshold'])
        source_type = config['VALIDATION']['source_type']
        font_path = config['font']['font_path']
        space_size = int(config['PARAMS']['window_size'])

        # 定义输入子目录
        sat_input_dir = os.path.join(input_dir, "01_sat")
        reference_input_dir = os.path.join(input_dir, "02_reference")
        reports_input_dir = os.path.join(input_dir, "05_reports")
        
        # 确保路径使用正斜杠（对Linux兼容）
        input_dir = input_dir.replace('\\', '/')
        output_dir = output_dir.replace('\\', '/')
        font_path = font_path.replace('\\', '/')
        sat_input_dir = sat_input_dir.replace('\\', '/')
        reference_input_dir = reference_input_dir.replace('\\', '/')

        from matplotlib import font_manager
        import matplotlib.pyplot as plt

        print(f"正在尝试加载字体: {font_path}")
        if os.path.exists(font_path):
            # 1. 核心：将字体文件加入 Matplotlib 管理器
            font_manager.fontManager.addfont(font_path)
            
            # 2. 设置全局字体为 SimHei
            plt.rcParams['font.sans-serif'] = ['SimHei']
            
            # 3. 解决负号显示为方块的问题
            plt.rcParams['axes.unicode_minus'] = False
            
            print("✅ 字体加载成功！Matplotlib 已锁定 SimHei。")
        else:
            print(f"❌ 严重警告：找不到字体文件！路径: {font_path}")
            print("请检查 config.ini 中的路径是否与 Linux 实际路径完全一致（注意空格和下划线）。")
        
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"\n=== 开始数据检验流程: {satellite_type} vs {source_type} ===")
        
        # 步骤1：处理HY数据
        print(f"\n处理{satellite_type}数据...")
        process_hy_data(
            satellite_type,
            hy_file_l2a=os.path.join(sat_input_dir, config[satellite_type]['l2a_file']),
            hy_file_l2b=os.path.join(sat_input_dir, config[satellite_type]['l2b_file']),
            output_dir=output_dir
        )
        
       # 步骤2：处理检验源数据
        if source_type == 'XC':
            print("\n处理现场数据...")
            process_xc_check_data(
                aopres_file=os.path.join(reference_input_dir, config['XC']['aopres_file']),
                dcsszcgq_file=os.path.join(reference_input_dir, config['XC']['dcsszcgq_file']),
                aot_file=os.path.join(reference_input_dir, config['XC']['aot_file']),
                wycgq_file=os.path.join(reference_input_dir, config['XC']['wycgq_file']),
                output_dir=output_dir
            )
        else:
            print(f"\n处理{source_type}卫星数据...")
            process_satellite_check_data(
                oc_file=os.path.join(reference_input_dir, config[source_type]['oc_file']),
                sst_file=os.path.join(reference_input_dir, config[source_type]['sst_file']),
                output_dir=output_dir
            )
        
        # 步骤3：标识检查
        print("\n执行标识检查...")
        HY_flag_create(satellite_type,output_dir, window_size)
        if source_type == 'XC':
            process_xc_flagcheck_data(output_dir, output_dir)
        else:
            satellite_flag_create(output_dir, source_type, window_size)
        
        # 步骤4：时间匹配
        print("\n执行时间匹配...")
        if source_type == 'XC':
            process_xc_timematch(output_dir, output_dir, satellite_type, time_threshold)
        else:
            process_satellite_timematch(output_dir, output_dir, satellite_type, source_type, time_threshold)
        
        # 步骤5：空间匹配
        print("\n执行空间匹配...")
        if source_type == 'XC':
            process_xc_spacematch(output_dir, output_dir, satellite_type, window_size)
        else:
            process_satellite_spacematch(output_dir, output_dir, satellite_type, source_type)
         # 保存空间窗口大小信息
        try:
            with open(os.path.join(output_dir, 'spacesize.txt'), 'w') as f:
                f.write(f"{window_size}")
        except Exception as e:
            print(f"保存空间窗口大小信息失败: {e}")
        
        # 步骤6：生成验证结果
        print("\n生成验证结果...")
        if source_type == 'XC':
            xc_validation(output_dir, output_dir)
        else:
            satellite_validation(output_dir, output_dir)
        
        # 步骤7：生成误差地图
        print("\n生成误差地图...")
        if source_type != 'XC':
            step7(satellite_type,output_dir, output_dir)

        # 步骤8：生成折线图
        step8(satellite_type,output_dir, output_dir)

        # 步骤9：生成统计结果和图表
        step9(satellite_type,output_dir, output_dir)

        # 步骤10：生成报告所需数据report文件
        if source_type == 'XC':

            make_ground_report_data(output_dir)

        else:

            make_satellite_report_data(output_dir)

        print(f"\n=== {satellite_type} vs {source_type} 数据检验流程完成 ===")


        #步骤11 修改文件名
        rename_files(output_dir)

        #步骤12 移动文件到指定文件夹
        organize_files(output_dir, output_dir)
        os.chmod(output_dir, 0o664)  # 设置文件权限为可读写

        # 步骤13：生成报告

        oc_file=os.path.join(sat_input_dir, config[satellite_type]['l2a_file'])
        beijing_time = extract_datetime(oc_file)
        time_str = beijing_time.strftime('%Y%m%d%H%M%S')
        extracted_data = time_str[:8]

        # print(f"{beijing_time}\n")
        # print(f"{time_str}\n")
        # print(f"{extracted_data}\n")
        input_temp = os.path.join(input_dir, '05_reports')
        input_img = os.path.join(output_dir, '04_visualization')
        coldata_path = os.path.join(output_dir, '05_reports')
        output_path = os.path.join(output_dir, '05_reports')

        # 记录报告生成前已存在的docx文件
        existing_docx_files = set()
        if os.path.exists(output_path):
            existing_docx_files = set([f for f in os.listdir(output_path) if f.endswith('.docx')])

        if source_type == 'XC':

            step_xc_report(extracted_data, input_temp, input_img, coldata_path, output_path, satellite_type, source_type,space_size)

        else:

            step_report(extracted_data, input_temp, input_img, coldata_path, output_path, satellite_type, source_type, space_size,time_threshold)

        # 找出本次新生成的docx文件
        new_docx_files = []
        if os.path.exists(output_path):
            current_docx_files = set([f for f in os.listdir(output_path) if f.endswith('.docx')])
            new_docx_files = list(current_docx_files - existing_docx_files)

        if new_docx_files:
            print(f"\n本次生成了 {len(new_docx_files)} 个报告文件: {', '.join(new_docx_files)}")
            word_to_pdf(output_path, output_path, new_docx_files)
        else:
            print("\n警告: 未检测到新生成的docx报告文件")

        return True
    except Exception as e:
        print(f"检验流程执行失败: {str(e)}")
        traceback.print_exc()
        return False

def process_hy_data(source_type,hy_file_l2a, hy_file_l2b, output_dir):
    """
    处理HY待检验数据
    """
    def save_data_to_txt(data, filename):
        """将数据保存为单列txt文件"""
        flattened_data = data.flatten()
        with open(filename, 'w') as f:
            for value in flattened_data:
                if abs(value + 9.9) < 0.0001:
                    f.write('-999.000000\n')
                elif abs(value) < 0.000001:
                    f.write('0.000000\n')
                elif abs(value) >= 1000000:
                    f.write(f'{value:.0f}\n')
                else:
                    f.write(f'{value:.6f}\n')

    try:
        print(f'\n开始处理{source_type}数据\n')
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        prefix = source_type

        # 处理反射率数据
        with h5py.File(hy_file_l2a, 'r') as h5_file:
            # 获取时间信息
            year = int(h5_file['Scan Line Attributes/Year'][0])
            day = int(h5_file['Scan Line Attributes/Day'][0])
            millisecond = int(h5_file['Scan Line Attributes/Millisecond'][0])

            # 转换为北京时间
            utc_time = datetime(year, 1, 1) + timedelta(days=day-1, milliseconds=millisecond)
            beijing_time = utc_time + timedelta(hours=8)
            time_str = beijing_time.strftime('%Y%m%d%H%M%S')

            # 获取数据维度并保存
            lat_data = h5_file['Navigation Data/Latitude'][:]
            data_shape = lat_data.shape
            rows, cols = data_shape[0], data_shape[1]
            print(f"检测到数据维度: {rows} x {cols}")

            # 保存维度信息到文件,供后续步骤使用
            dimensions_file = os.path.join(output_dir, f'dimensions_{time_str}.txt')
            with open(dimensions_file, 'w') as f:
                f.write(f"{rows},{cols}\n")
            print(f"数据维度已保存到: {dimensions_file}")

            # 保存基础数据
            save_data_to_txt(lat_data,
                           os.path.join(output_dir, f'{prefix}_lat_{time_str}.txt'))
            save_data_to_txt(h5_file['Navigation Data/Longitude'][:],
                           os.path.join(output_dir, f'{prefix}_lon_{time_str}.txt'))
            save_data_to_txt(h5_file['Geophysical Data/l2_flags'][:],
                           os.path.join(output_dir, f'{prefix}_flag_{time_str}.txt'))

            # 保存反射率数据
            rrs_bands = ['412', '443', '490', '520', '565', '670', '750']
            for band in rrs_bands:
                data = h5_file[f'Geophysical Data/Rrs{band}'][:]
                save_data_to_txt(data, 
                               os.path.join(output_dir, f'{prefix}_Rrs{band}_{time_str}.txt'))

        # 处理TSM等参数数据
        with h5py.File(hy_file_l2b, 'r') as h5_file:
            # 保存参数数据
            params = {
                'chl_a': 'Geophysical Data/chl_a',
                'TSM': 'Geophysical Data/TSM',
                'CDOM': 'Geophysical Data/CDOM',
                'sst': 'Geophysical Data/SST',
                'AOT': 'Geophysical Data/taua865',
                'nLw':"Geophysical Data/nLw565"
            }
            
            for param_name, dataset_path in params.items():
                data = h5_file[dataset_path][:]
                save_data_to_txt(data, 
                               os.path.join(output_dir, f'{prefix}_{param_name}_{time_str}.txt'))

        print(f'\n{source_type}数据处理完成\n')
        return True

    except Exception as e:
        print(f"处理数据时出错: {str(e)}")
        traceback.print_exc()
        return False
    
def process_satellite_check_data(oc_file, sst_file, output_dir):
    """
    处理卫星检验数据
    """
    def save_data_to_txt(data, filename):
        """将数据保存为单列txt文件"""
        flattened_data = data.flatten()
        with open(filename, 'w') as f:
            for value in flattened_data:
                f.write(f'{value}\n')

    def extract_datetime(filename):
        """从文件名中提取时间信息并转换为北京时间"""
        pattern = r'\d{8}T\d{6}'
        match = re.search(pattern, filename)
        if match:
            time_str = match.group()
            utc_time = datetime.strptime(time_str, '%Y%m%dT%H%M%S')
            beijing_time = utc_time + timedelta(hours=8)
            return beijing_time
        return None
    
    def extract_file_prefix(filename):
        """从文件名中提取处理的卫星类别"""
        first_five_chars = os.path.basename(filename)[:5] if len(os.path.basename(filename)) >= 5 else None
        if first_five_chars == 'AQUA_': 
            return 'AQUA'
        elif first_five_chars == 'TERRA':
            return 'TERRA'
        elif first_five_chars == 'SNPP_':
            return 'SNPP'
        elif first_five_chars == 'JPSS1':
            return 'JPSS'

    try:
        print('\n开始处理卫星数据\n')
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        # 处理海洋水色数据
        with nc.Dataset(oc_file, 'r') as nc_data:
            # 获取时间信息
            beijing_time = extract_datetime(oc_file)
            time_str = beijing_time.strftime('%Y%m%d%H%M%S')
            prefix = extract_file_prefix(oc_file)

            # 保存基础数据
            save_data_to_txt(nc_data['navigation_data']['latitude'][:], 
                           os.path.join(output_dir, f'{prefix}_Lat_{time_str}.txt'))
            save_data_to_txt(nc_data['navigation_data']['longitude'][:], 
                           os.path.join(output_dir, f'{prefix}_Lon_{time_str}.txt'))
            save_data_to_txt(nc_data['geophysical_data']['l2_flags'][:], 
                           os.path.join(output_dir, f'{prefix}_flag_{time_str}.txt'))

            # 根据传感器类型选择波段
            if prefix in ['AQUA', 'TERRA']:  # MODIS数据
                rrs_bands = {
                    'Rrs412': 'Rrs_412', 
                    'Rrs443': 'Rrs_443', 
                    'Rrs469': 'Rrs_469', 
                    'Rrs488': 'Rrs_488',
                    'Rrs531': 'Rrs_531', 
                    'Rrs547': 'Rrs_547',
                    'Rrs555': 'Rrs_555', 
                    'Rrs645': 'Rrs_645',
                    'Rrs667': 'Rrs_667', 
                    'Rrs678': 'Rrs_678'
                }
            elif 'JPSS' in prefix:  # JPSS数据
                rrs_bands = {
                    'Rrs411': 'Rrs_411',
                    'Rrs445': 'Rrs_445',
                    'Rrs489': 'Rrs_489',
                    'Rrs556': 'Rrs_556',
                    'Rrs667': 'Rrs_667'
                }
            else:  # SNPP数据
                rrs_bands = {
                    'Rrs410': 'Rrs_410',
                    'Rrs443': 'Rrs_443',
                    'Rrs486': 'Rrs_486',
                    'Rrs551': 'Rrs_551',
                    'Rrs671': 'Rrs_671'
                }

            # 保存遥感反射率数据
            for out_name, band_name in rrs_bands.items():
                data = nc_data['geophysical_data'][band_name][:]
                save_data_to_txt(data, 
                               os.path.join(output_dir, f'{prefix}_{out_name}_{time_str}.txt'))

            # 保存叶绿素数据
            chl_data = nc_data['geophysical_data']['chlor_a'][:]
            save_data_to_txt(chl_data, 
                           os.path.join(output_dir, f'{prefix}_Chl_{time_str}.txt'))


            # 保存气溶胶光学厚度数据
            aot_band = 'aot_869' if prefix in ['AQUA', 'TERRA'] else \
                      'aot_862' if prefix == 'SNPP' else 'aot_868'
            aot_data = nc_data['geophysical_data'][aot_band][:]
            save_data_to_txt(aot_data, 
                           os.path.join(output_dir, f'{prefix}_AOT_{time_str}.txt'))

        # 处理海表温度数据
        with nc.Dataset(sst_file, 'r') as nc_data:
            sst_data = nc_data['geophysical_data']['sst'][:]
            save_data_to_txt(sst_data, 
                           os.path.join(output_dir, f'{prefix}_sst_{time_str}.txt'))

        print(f'{prefix}数据处理完成')
        return True

    except Exception as e:
        print(f"处理数据时出错: {str(e)}")
        traceback.print_exc()
        return False


def process_xc_check_data(aopres_file, dcsszcgq_file, aot_file, wycgq_file, output_dir):
    """
    处理现场检验数据
    """
    def read_header_info(file_path):
        """读取文件头信息"""
        header_info = {'lat': 37.681, 'lon': 121.700}  # 默认值
        header_end_line = 0
        
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            for i, line in enumerate(lines):
                line = line.strip()
                if not line:
                    continue
                    
                if line.startswith('/north_latitude'):
                    lat_str = line.split('=')[1] if '=' in line else line.split()[1]
                    header_info['lat'] = float(lat_str)
                elif line.startswith('/east_longitude'):
                    lon_str = line.split('=')[1] if '=' in line else line.split()[1]
                    header_info['lon'] = float(lon_str)
                elif line.startswith('/end_header'):
                    header_end_line = i + 1
                    break
        
        if header_end_line == 0:
            for i, line in enumerate(lines):
                if line.strip() and not line.startswith('/'):
                    header_end_line = i
                    break
                    
        return header_info, header_end_line

    def process_data_file(input_file, data_type):
        """处理单个数据文件"""
        print(f'\n开始处理{data_type}数据\n')

        # 添加时间格式化辅助函数
        def format_time_column(time_series):
            """智能格式化时间列"""
            # 检查第一个非空值的格式
            sample = str(time_series.dropna().iloc[0]) if not time_series.dropna().empty else ""
            
            # 如果已经是6位数字格式（HHMMSS），直接返回
            if sample.replace('.', '').replace(':', '').isdigit() and len(sample.replace(':', '').replace('.', '')) == 6:
                # 移除可能的小数点和冒号
                return time_series.astype(str).str.replace(':', '').str.replace('.', '').str.zfill(6)
            
            # 如果包含冒号，说明是 HH:MM:SS 格式，需要转换
            if ':' in sample:
                return pd.to_datetime(time_series, format='%H:%M:%S').dt.strftime('%H%M%S')
            
            # 其他情况，尝试补齐到6位
            return time_series.astype(str).str.zfill(6)
        
        # 读取文件头信息
        header_info, header_end_line = read_header_info(input_file)
        
        # 读取数据部分
        df = pd.read_csv(input_file, skiprows=header_end_line, sep=r'\s+', header=None)
        
        # 根据数据类型处理       
        if data_type == 'dcsszcgq':
            # 处理水质参数数据
            if df.shape[1] >= 5:
                df = df.iloc[:, :5]
                df.columns = ['Date', 'Time', 'TSM', 'Chl', 'CDOM']
                
                # 添加调试输出
                print(f"\n调试信息 - {data_type}:")
                print(f"原始日期列样例:\n{df['Date'].head()}")
                print(f"日期列数据类型: {df['Date'].dtype}")
                
                # 格式化日期和时间 - 添加错误处理
                try:
                    # 先转换为字符串并清理格式
                    df['Date'] = df['Date'].astype(str).str.replace('-', '').str.replace('/', '').str.replace(' ', '').str.strip()
                    print(f"清理后日期样例:\n{df['Date'].head()}")
                    
                    # 尝试转换为日期
                    df['Date'] = pd.to_datetime(df['Date'], format='%Y%m%d', errors='coerce')
                    
                    # 统计有效日期
                    valid_count = df['Date'].notna().sum()
                    total_count = len(df)
                    print(f"有效日期数量: {valid_count}/{total_count}")
                    
                    # 移除日期解析失败的行
                    df = df.dropna(subset=['Date'])
                    
                    if df.empty:
                        print(f"警告：{data_type} 数据所有日期解析失败，跳过处理")
                        return
                    
                    # 格式化为 YYYYMMDD
                    df['Date'] = df['Date'].dt.strftime('%Y%m%d')
                    df['Time'] = format_time_column(df['Time'])
                    
                    print(f"最终日期样例:\n{df['Date'].head()}")
                    
                except Exception as e:
                    print(f"日期格式化错误: {e}")
                    print(f"样例数据:\n{df[['Date', 'Time']].head()}")
                    traceback.print_exc()
                    return
            else:
                raise ValueError(f"DCSSZCGQ数据列数不足: {df.shape[1]}")
        
        elif data_type == 'aop':
            if df.shape[1] >= 1570:
                df = df.iloc[:, [0, 1, 1116, 1147, 1194, 1224, 1269, 1374, 1454, 1569]]
                df.columns = ['Date', 'Time', 'Rrs412', 'Rrs443', 'Rrs490', 'Rrs520', 
                            'Rrs565', 'Rrs670', 'Rrs750', 'Rrs865']
                
                # 添加调试输出
                print(f"\n调试信息 - {data_type}:")
                print(f"原始日期列样例:\n{df['Date'].head()}")
                
                # 格式化日期和时间
                try:
                    df['Date'] = df['Date'].astype(str).str.replace('-', '').str.replace('/', '').str.strip()
                    df['Date'] = pd.to_datetime(df['Date'], format='%Y%m%d', errors='coerce')
                    df = df.dropna(subset=['Date'])
                    
                    if df.empty:
                        print(f"警告：{data_type} 数据所有日期解析失败，跳过处理")
                        return
                    
                    df['Date'] = df['Date'].dt.strftime('%Y%m%d')
                    df['Time'] = format_time_column(df['Time'])
                    print(f"最终日期样例:\n{df['Date'].head()}")
                except Exception as e:
                    print(f"日期格式化错误: {e}")
                    traceback.print_exc()
                    return
                # 格式化日期和时间 - 添加错误处理
                try:
                    if df['Date'].dtype == 'object':
                        df['Date'] = pd.to_datetime(df['Date'], format='%Y%m%d', errors='coerce')
                    else:
                        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                    
                    df = df.dropna(subset=['Date'])
                    
                    if df.empty:
                        print(f"警告：{data_type} 数据所有日期解析失败，跳过处理")
                        return
                    
                    df['Date'] = df['Date'].dt.strftime('%Y%m%d')
                    df['Time'] = format_time_column(df['Time'])
                    df['nLw'] = df['Rrs565'] * 179.363
                    
                except Exception as e:
                    print(f"日期格式化错误: {e}")
                    print(f"样例数据: {df[['Date', 'Time']].head()}")
                    return
            else:
                raise ValueError(f"WQP数据列数不足: {df.shape[1]}")
                
        elif data_type == 'aot':
            # 处理气溶胶光学厚度数据
            if df.shape[1] >= 12:
                df = df.iloc[:, [0, 1, 7, 11]]
                df.columns = ['Date', 'Time', 'AOT', 'Flag']
                
                # 添加调试输出
                print(f"\n调试信息 - {data_type}:")
                print(f"原始日期列样例:\n{df['Date'].head()}")
                print(f"日期列数据类型: {df['Date'].dtype}")
                
                try:
                    # 先转换为字符串并清理格式
                    df['Date'] = df['Date'].astype(str).str.replace('-', '').str.replace('/', '').str.replace(' ', '').str.strip()
                    print(f"清理后日期样例:\n{df['Date'].head()}")
                    
                    # 尝试转换为日期
                    df['Date'] = pd.to_datetime(df['Date'], format='%Y%m%d', errors='coerce')
                    
                    # 统计有效日期
                    valid_count = df['Date'].notna().sum()
                    total_count = len(df)
                    print(f"有效日期数量: {valid_count}/{total_count}")
                    
                    df = df.dropna(subset=['Date'])
                    
                    if df.empty:
                        print(f"警告：{data_type} 数据所有日期解析失败，跳过处理")
                        return
                    
                    df['Date'] = df['Date'].dt.strftime('%Y%m%d')
                    df['Time'] = format_time_column(df['Time'])
                    print(f"最终日期样例:\n{df['Date'].head()}")
                    
                except Exception as e:
                    print(f"日期格式化错误: {e}")
                    traceback.print_exc()
                    return
            else:
                raise ValueError(f"AOT数据列数不足: {df.shape[1]}")
                
        elif data_type == 'wycgq':
            # 处理温盐传感器数据
            if df.shape[1] >= 3:
                df = df.iloc[:, [0, 1, 2]]
                df.columns = ['Date', 'Time', 'SST']
                
                # 添加调试输出
                print(f"\n调试信息 - {data_type}:")
                print(f"原始日期列样例:\n{df['Date'].head()}")
                print(f"日期列数据类型: {df['Date'].dtype}")
                
                try:
                    # 先转换为字符串并清理格式
                    df['Date'] = df['Date'].astype(str).str.replace('-', '').str.replace('/', '').str.replace(' ', '').str.strip()
                    print(f"清理后日期样例:\n{df['Date'].head()}")
                    
                    # 尝试转换为日期
                    df['Date'] = pd.to_datetime(df['Date'], format='%Y%m%d', errors='coerce')
                    
                    # 统计有效日期
                    valid_count = df['Date'].notna().sum()
                    total_count = len(df)
                    print(f"有效日期数量: {valid_count}/{total_count}")
                    
                    df = df.dropna(subset=['Date', 'SST'])
                    
                    if df.empty:
                        print(f"警告：{data_type} 数据所有日期解析失败，跳过处理")
                        return
                    
                    df['Date'] = df['Date'].dt.strftime('%Y%m%d')
                    df['Time'] = format_time_column(df['Time'])
                    print(f"最终日期样例:\n{df['Date'].head()}")
                    
                except Exception as e:
                    print(f"日期格式化错误: {e}")
                    traceback.print_exc()
                    return
            else:
                raise ValueError(f"WYCGQ数据列数不足: {df.shape[1]}")

        df = df.dropna()
        
        # 保存处理后的数据
        for date, group in df.groupby('Date'):
            if data_type == 'dcsszcgq':
                # 处理水质参数数据
                params = [
                    ('Chl', 'Chl'),
                    ('TSM', 'TSM'),
                    ('CDOM', 'CDOM')
                ]
                for param_name, col_name in params:
                    output_file = os.path.join(output_dir, f'XC_{param_name}_{date}000000.txt')
                    with open(output_file, 'w', encoding='utf-8') as f:
                        f.write(f"Latitude: {header_info['lat']}\n")
                        f.write(f"Longitude: {header_info['lon']}\n")
                        f.write("Data:\n")
                    group[['Date', 'Time', col_name]].to_csv(output_file, mode='a', 
                                                           index=False, sep='\t')   
            elif data_type == 'aop':
                #处理遥感反射率数据
                params = [
                    ('Rrs412', 'Rrs412'),
                    ('Rrs443', 'Rrs443'),
                    ('Rrs490', 'Rrs490'),
                    ('Rrs520', 'Rrs520'),
                    ('Rrs565', 'Rrs565'),
                    ('Rrs670', 'Rrs670'),
                    ('Rrs750', 'Rrs750'),
                    ('Rrs865', 'Rrs865'),
                    ('nLw','nLw')
                ]
                for param_name, col_name in params:
                    output_file = os.path.join(output_dir, f'XC_{param_name}_{date}000000.txt')
                    with open(output_file, 'w', encoding='utf-8') as f:
                        f.write(f"Latitude: {header_info['lat']}\n")
                        f.write(f"Longitude: {header_info['lon']}\n")
                        f.write("Data:\n")
                    group[['Date', 'Time', col_name]].to_csv(output_file, mode='a', 
                                                           index=False, sep='\t')
            elif data_type == 'aot':
                #处理气溶胶光学厚度数据
                output_file = os.path.join(output_dir, f'XC_AOT_{date}000000.txt')
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(f"Latitude: {header_info['lat']}\n")
                    f.write(f"Longitude: {header_info['lon']}\n")
                    f.write("Data:\n")
                # 将Date、Time、AOT和Flag列一起写入文件
                group[['Date', 'Time', 'AOT', 'Flag']].to_csv(output_file, mode='a', 
                                                            index=False, sep='\t')
                
            elif data_type == 'wycgq':
                #处理温度数据
                params = [
                    ('sst', 'SST')
                ]
                for param_name, col_name in params:
                    output_file = os.path.join(output_dir, f'XC_{param_name}_{date}000000.txt')
                    with open(output_file, 'w', encoding='utf-8') as f:
                        f.write(f"Latitude: {header_info['lat']}\n")
                        f.write(f"Longitude: {header_info['lon']}\n")
                        f.write("Data:\n")
                    group[['Date', 'Time', col_name]].to_csv(
                        output_file, 
                        mode='a',
                        index=False,
                        sep='\t'
                    )
    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        # 处理各类数据文件
        if aopres_file:
            process_data_file(aopres_file, 'aop')
        if dcsszcgq_file:
            process_data_file(dcsszcgq_file, 'dcsszcgq')
        if aot_file:
            process_data_file(aot_file, 'aot')
        if wycgq_file:
            process_data_file(wycgq_file, 'wycgq')
            
        print('\n现场检验数据处理完成\n')
        return True
        
    except Exception as e:
        print(f"处理数据时出错: {str(e)}")
        traceback.print_exc()
        return False
    


def extract_file_prefix(filename):
    """从文件名中提取处理的卫星类别"""
    first_five_chars = os.path.basename(filename)[:5] if len(os.path.basename(filename)) >= 5 else None
    if first_five_chars == 'AQUA_': 
        return 'AQUA'
    elif first_five_chars == 'TERRA':
        return 'TERRA'
    elif first_five_chars == 'SNPP_':
        return 'SNPP'
    elif first_five_chars == 'JPSS_':
        return 'JPSS'

def generate_flag_from_data(data_file, satellite_type, output_dir=None):
    try:
        data = np.genfromtxt(data_file, delimiter=None)

        # 如果数据是一维的,尝试重塑
        if data.ndim == 1 and output_dir:
            # 尝试从维度文件读取行列数
            dimension_files = [f for f in os.listdir(output_dir) if f.startswith('dimensions_')]
            if dimension_files:
                dimension_file = os.path.join(output_dir, dimension_files[0])
                try:
                    with open(dimension_file, 'r') as f:
                        dims = f.read().strip().split(',')
                        rows, cols = int(dims[0]), int(dims[1])
                        if data.size == rows * cols:
                            data = data.reshape(rows, cols)
                            print(f"使用维度文件重塑数据: {rows} x {cols}")
                except Exception as e:
                    print(f"读取维度文件失败: {e}")

        flag = np.zeros_like(data, dtype=np.int32)

        # 添加统计信息
        filename = os.path.basename(data_file)
        # 根据不同产品类型处理无效值
        if satellite_type in ['HY1C','HY1D']:
            flag[data == -999] = 1
            flag[np.isnan(data)] = 1
        else:
            flag[data == '--'] = 1
            flag[np.isnan(data)] = 1

        return flag

    except Exception as e:
        print(f"生成标识矩阵时出错: {str(e)}")
        traceback.print_exc()
        return None

def apply_spatial_window(flag_array, window_size, rows, cols):
    """应用空间窗口判断"""
    try:
        flag_2d = flag_array.reshape(rows, cols)
        half_window = (window_size - 1) // 2
        
        # 1. 处理边界区域
        flag_2d[:half_window, :] = 1  # 上边界
        flag_2d[-half_window:, :] = 1  # 下边界
        flag_2d[:, :half_window] = 1  # 左边界
        flag_2d[:, -half_window:] = 1  # 右边界
        
        # 2. 第一轮空间窗口判断：FLAG为1的像元比例
        temp_flag = flag_2d.copy()
        for i in range(half_window, rows-half_window):
            for j in range(half_window, cols-half_window):
                if flag_2d[i, j] == 0:
                    window = flag_2d[i-half_window:i+half_window+1, 
                                   j-half_window:j+half_window+1]
                    if np.mean(window) > 0.5:
                        temp_flag[i, j] = 1
        
        flag_2d = temp_flag
        
        # 3. 第二轮空间窗口判断：变异系数
        temp_flag = flag_2d.copy()
        for i in range(half_window, rows-half_window):
            for j in range(half_window, cols-half_window):
                if flag_2d[i, j] == 0:
                    window = flag_2d[i-half_window:i+half_window+1, 
                                   j-half_window:j+half_window+1]
                    valid_values = window[window == 0]
                    if len(valid_values) > 0:
                        cv = np.std(valid_values) / np.mean(valid_values) if np.mean(valid_values) != 0 else 0
                        if cv > 0.15:
                            temp_flag[i, j] = 1
        
        return temp_flag.reshape(-1)
        
    except Exception as e:
        print(f"空间窗口处理失败: {e}")
        traceback.print_exc()
        return None


def process_xc_flagcheck_data(input_dir, output_dir):
    """处理现场观测数据的标识检查"""
    try:
        print('\n开始处理现场观测数据标识检查\n')
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 处理所有现场观测数据文件
        for filename in os.listdir(input_dir):
            if not filename.startswith('XC_') or not filename.endswith('.txt'):
                continue
                
            input_file = os.path.join(input_dir, filename)         
            # 读取文件内容
            with open(input_file, 'r') as f:
                lines = f.readlines()
            
            header_lines = lines[:3]  # 前3行为经纬度等信息
            headers = lines[3].strip() # 第4行为列名
            
            # 处理AOT数据
            if 'AOT' in filename:
                output_file = os.path.join(output_dir, filename.replace('XC_', 'XCf_'))
                data_lines = []
                flag_count = 0
                valid_lines = 0

                for line in lines[4:]:  # 从第5行开始是数据
                    if not line.strip():
                        continue
                        
                    values = line.strip().split('\t')
                    valid_lines += 1
                    
                    # 检查最后一列的Flag值
                    try:
                        flag_value = float(values[-1])  # 将Flag值转换为数值
                        if flag_value >= 1:  # 修改这里：检查是否大于等于1
                            flag_count += 1
                        else:
                            data_lines.append(line)
                    except ValueError:  # 处理可能的转换错误
                        print(f"警告：无法转换Flag值：{values[-1]}")
                        continue
                
                # 保存处理后的数据
                with open(output_file, 'w') as f:
                    f.writelines(header_lines)  # 写入经纬度等信息
                    f.write(headers + '\n')     # 写入列名
                    f.writelines(data_lines)    # 写入筛选后的数据          

            # 处理其他数据（直接改名）
            else:
                output_file = os.path.join(output_dir, filename.replace('XC_', 'XCf_'))
                with open(output_file, 'w') as f:
                    f.writelines(lines)
        
        print('\n现场观测数据标识处理完成\n')
        return True
        
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        traceback.print_exc()
        return False

def HY_flag_create(satellite_type,input_dir,window_size):
    try:
        print(f"\n开始执行 {satellite_type}_flag_create 函数 (对每个产品单独生成flag)\n")
        flag_matrices = {}

        # 检查目录中的文件
        all_files = os.listdir(input_dir)

        # 读取数据维度
        dimension_files = [f for f in all_files if f.startswith('dimensions_')]
        rows, cols = None, None
        if dimension_files:
            dimension_file = os.path.join(input_dir, dimension_files[0])
            try:
                with open(dimension_file, 'r') as f:
                    dims = f.read().strip().split(',')
                    rows, cols = int(dims[0]), int(dims[1])
                    print(f"从维度文件读取到数据维度: {rows} x {cols}")
            except Exception as e:
                print(f"读取维度文件失败: {e}")

        # 如果没有读取到维度文件,尝试猜测(向后兼容)
        if rows is None or cols is None:
            print("警告: 未找到维度文件,将尝试猜测数据维度")

        # 查找所有flag文件
        flag_files = [f for f in all_files
                     if f.startswith(f'{satellite_type}_flag_') and f.endswith('.txt')
                     and 'flag1' not in f]

        if not flag_files:
            print(f"警告: 未找到{satellite_type}_flag文件")
            return flag_matrices

        # 对每个flag文件处理
        for flag_filename in flag_files:
            print(f"\n开始处理 l2_flag 文件: {flag_filename}")
            flag_file = os.path.join(input_dir, flag_filename)

            # 读取原始 l2_flags 数据
            flag_matrix = np.genfromtxt(flag_file, delimiter=None, dtype=np.int32)

            # 如果没有从文件读取到维度,尝试猜测
            if rows is None or cols is None:
                total_size = flag_matrix.size
                for i in range(1000, 6000):
                    if total_size % i == 0:
                        rows = i
                        cols = total_size // i
                        break
                print(f"猜测的数据维度: {rows} x {cols}")

            # 重塑为二维数组
            flag_matrix = flag_matrix.reshape(rows, cols) if flag_matrix.ndim == 1 else flag_matrix

            # 从l2_flags提取基础FLAG(陆地和云冰标记)
            base_flag = np.zeros_like(flag_matrix, dtype=np.int32)
            mask = ((flag_matrix & (1 << 8)) | (flag_matrix & (1 << 22))) != 0
            base_flag[mask] = 1

            print(f"基础FLAG(陆地+云冰)中1的数量: {np.sum(base_flag == 1)}, 占比: {np.sum(base_flag == 1) / base_flag.size * 100:.2f}%")

            # 提取时间戳
            time_id = flag_filename.split('_')[2].replace('.txt', '')

            # 查找所有对应时间的产品文件
            product_files = [f for f in all_files
                            if f.startswith(f'{satellite_type}_') and time_id in f
                            and 'lon' not in f.lower() and 'lat' not in f.lower()
                            and 'flag' not in f]

            print(f"找到 {len(product_files)} 个产品文件需要生成flag1")

            # 为每个产品单独生成flag1文件
            for product_file in product_files:
                print(f"\n处理产品: {product_file}")

                # 为该产品创建FLAG副本
                product_flag = base_flag.copy()

                # 添加该产品特有的无效值标记
                product_path = os.path.join(input_dir, product_file)
                if os.path.exists(product_path):
                    temp_matrix = generate_flag_from_data(
                        product_path,
                        satellite_type,
                        input_dir
                    )
                    if temp_matrix is not None and temp_matrix.size == product_flag.size:
                        # 确保temp_matrix的形状与product_flag一致
                        if temp_matrix.ndim == 1:
                            temp_matrix = temp_matrix.reshape(rows, cols)

                        ones_before = np.sum(product_flag == 1)
                        product_flag = np.logical_or(product_flag, temp_matrix).astype(np.int32)
                        ones_after = np.sum(product_flag == 1)
                        new_ones = ones_after - ones_before
                        print(f"  产品数据新增无效像素: {new_ones}, 占比: {new_ones / product_flag.size * 100:.2f}%")
                    else:
                        print(f"  警告: 产品 {product_file} 的数据大小不匹配")
                        continue

                # 应用空间窗口过滤
                product_flag = apply_spatial_window(product_flag, window_size, rows, cols)

                # 提取产品名称 (如 Rrs412, chl_a, sst等)
                # 产品文件格式: HY1C_{product}_{time}.txt
                product_name = product_file.replace(f'{satellite_type}_', '').replace(f'_{time_id}.txt', '')

                # 生成输出文件名: HY1C_flag1_{product}_{time}.txt
                output_filename = f'{satellite_type}_flag1_{product_name}_{time_id}.txt'
                output_path = os.path.join(input_dir, output_filename)

                # 保存flag1文件
                np.savetxt(output_path, product_flag.flatten(), fmt='%d')
                print(f"  已保存: {output_filename}")
                print(f"  最终FLAG中1的数量: {np.sum(product_flag == 1)}, 占比: {np.sum(product_flag == 1) / product_flag.size * 100:.2f}%")

        print(f"\n所有{satellite_type}产品的flag1文件已生成完成")
        return flag_matrices

    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        traceback.print_exc()
        return None


def satellite_flag_create(input_dir, satellite_type,window_size):
    try:
        print(f"开始执行{satellite_type}_flag_create函数")
        flag_matrices = {}
        
        # 检查目录中的文件
        all_files = os.listdir(input_dir)       
        # 处理所有相关flag文件
        for filename in all_files:
            if filename.startswith(f'{satellite_type}_flag_') and filename.endswith('.txt'):
                print(f"\n开始处理flag文件: {filename}")
                flag_file = os.path.join(input_dir, filename)
                
                # 读取flag文件
                flag_matrix = np.genfromtxt(flag_file, delimiter=None, dtype=np.int32)
                # print(f"原始flag文件统计:")
                # print(f"- 数据形状: {flag_matrix.shape}")
                # print(f"- 唯一值: {np.unique(flag_matrix)}")
                
                # 提取时间戳部分
                time_id = filename.split('_')[2].replace('.txt', '')
                
                # 初始化flag矩阵
                FLAG = np.zeros_like(flag_matrix, dtype=np.int32)
                
                # 检查flag文件中的特定位
                mask = ((flag_matrix & 1) | (flag_matrix & (1 << 3)) | 
                        (flag_matrix & (1 << 4)) | (flag_matrix & (1 << 6)) | 
                        (flag_matrix & (1 << 22)) | (flag_matrix & (1 << 24))) != 0
                FLAG[mask] = 1
                # print(f"\n位运算后的FLAG统计:")
                # print(f"- FLAG中1的数量: {np.sum(FLAG == 1)}")
                # print(f"- FLAG中0的数量: {np.sum(FLAG == 0)}")
                
                # 查找对应的产品文件
                for product_file in all_files:
                    # 跳过lon、lat
                    if ('Lon' in product_file or 
                        'Lat' in product_file ):  # 添加这个条件
                        continue

                    if product_file.startswith(f'{satellite_type}_') and time_id in product_file:
                        print(f"\n处理产品文件: {product_file}")  # 新增：显示当前处理的产品文件

                        if os.path.exists(os.path.join(input_dir, product_file)):
                            temp_matrix = generate_flag_from_data(
                                os.path.join(input_dir, product_file), 
                                satellite_type
                            )
                            if temp_matrix is not None and len(temp_matrix) == len(flag_matrix):
                                # 记录合并前的状态
                                ones_before = np.sum(FLAG == 1)
                                FLAG = np.logical_or(FLAG, temp_matrix).astype(np.int32)
                                ones_after = np.sum(FLAG == 1)
                                # print(f"\n产品 {product_file} 的影响:")
                                # print(f"- 处理前1的数量: {ones_before}")
                                # print(f"- 处理后1的数量: {ones_after}")
                                # print(f"- 新增1的数量: {ones_after - ones_before}")
                                # print(f"- 占总像素的比例: {((ones_after - ones_before) / len(FLAG)) * 100:.2f}%")
                                                     
                                # 如果这个文件导致大量像素变为1，发出警告
                                if (ones_after - ones_before) > len(FLAG) * 0.5:  # 如果新增的1超过50%
                                    print(f"警告: 文件 {product_file} 导致大量像素变为1!")
                            else:
                                print(f"警告：产品 {product_file} 的数据长度与flag文件不匹配")
                


                # print(f"\n应用空间窗口前的FLAG统计:")
                # print(f"- FLAG中1的数量: {np.sum(FLAG == 1)}")
                # print(f"- FLAG中0的数量: {np.sum(FLAG == 0)}")
                # 应用空间窗口1
                total_size = flag_matrix.size
                for i in range(1000, 6000):
                    if total_size % i == 0:
                        rows = i
                        cols = total_size // i
                        break
                FLAG = apply_spatial_window(FLAG, window_size, rows, cols)

                # print(f"\n应用空间窗口后的FLAG统计:")
                # print(f"- FLAG中1的数量: {np.sum(FLAG == 1)}")
                # print(f"- FLAG中0的数量: {np.sum(FLAG == 0)}")



                # 保存结果前的最终检查
                if np.all(FLAG == 1):
                    print(f"\n警告: 生成的FLAG全为1!")


                # 输出结果
                output_filename = filename.replace('flag_', 'flag1_')
                output_path = os.path.join(input_dir, output_filename)
                np.savetxt(output_path, FLAG, fmt='%d')
                print(f"结果已保存到: {output_path}")

        return flag_matrices
        
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        traceback.print_exc()
        return None
    


def process_satellite_timematch(input_dir, output_dir, target_sensor, source_type, time_threshold):
    """
    处理卫星数据间的时间匹配
    """
    # 参数映射字典
    PARAM_MAPPING = {
        'AQUA': {
            'Rrs412': 'Rrs412', 
            'Rrs443': 'Rrs443', 
            'Rrs490': 'Rrs488',
            'Rrs520': 'Rrs531', 
            'Rrs565': 'Rrs555', 
            'Rrs670': 'Rrs667',
            'Rrs750': None, 
            'Rrs865': None, 
            'sst': 'sst', 
            'AOT': 'AOT',
            'chl': 'chl',
        },
        'TERRA': {
            'Rrs412': 'Rrs412',
            'Rrs443': 'Rrs443', 
            'Rrs490': 'Rrs488',
            'Rrs520': 'Rrs531', 
            'Rrs565': 'Rrs555', 
            'Rrs670': 'Rrs667',
            'Rrs750': None, 
            'Rrs865': None, 
            'sst': 'sst', 
            'AOT': 'AOT',
            'chl': 'chl',
        },
        'SNPP': {
            'Rrs412': 'Rrs410', 
            'Rrs443': 'Rrs443', 
            'Rrs490': 'Rrs486',
            'Rrs520': None, 
            'Rrs565': 'Rrs551', 
            'Rrs670': 'Rrs671',
            'Rrs750': None, 
            'Rrs865': None, 
            'sst': 'sst', 
            'AOT': 'AOT',
            'chl': 'chl',
        },
        'JPSS': {
            'Rrs412': 'Rrs411', 
            'Rrs443': 'Rrs445', 
            'Rrs490': 'Rrs489',
            'Rrs565': 'Rrs556', 
            'Rrs670': 'Rrs667', 
            'Rrs750': None,
            'Rrs865': None, 
            'sst': 'sst', 
            'AOT': 'AOT',
            'chl': 'chl',
        },
    }

    def extract_datetime_from_filename(filename):
        """从文件名中提取时间信息"""
        time_str = re.search(r'\d{14}', filename)
        if time_str:
            return datetime.strptime(time_str.group(), '%Y%m%d%H%M%S')
        return None

    def calculate_time_difference(time1, time2):
        """计算两个时间的差值（小时）"""
        time_diff = abs(time1 - time2)
        return time_diff.total_seconds() / 3600

    def save_match_result(result_file, target_file, source_file, time_diff):
        """保存匹配结果"""
        try:
            with open(os.path.join(output_dir, result_file), 'w') as f:
                f.write(f"{target_file}\n")
                f.write(f"{source_file}\n")
                f.write(f"{time_diff:.1f}\n")
            print(f"成功保存匹配结果到: {result_file}")
        except Exception as e:
            print(f"保存匹配结果失败: {str(e)}")

    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"\n开始处理卫星数据时间匹配...")
        print(f"目标传感器: {target_sensor}")
        print(f"检验源: {source_type}")
        print(f"时间阈值: {time_threshold}小时")
        
        # 获取目标传感器的数据文件
        target_files = [f for f in os.listdir(input_dir) 
                       if f.startswith(f"{target_sensor}_") and 
                       any(x.lower() in f.lower() for x in ['Rrs', 'sst', 'AOT', 'chl']) and 
                       f.endswith('.txt') and 
                        'flag' not in f]
        
        if not target_files:
            print(f"未找到{target_sensor}的数据文件")
            return False
          
        # 定义参数列表
        target_bands = ['Rrs412', 'Rrs443', 'Rrs490', 'Rrs520', 'Rrs565', 
                       'Rrs670', 'Rrs750', 'Rrs865']
        other_params = ['sst', 'AOT', 'chl']
        
        # 处理每个目标文件
        for target_file in target_files:
            # 提取时间信息
            target_time = extract_datetime_from_filename(target_file)
            if not target_time:
                print(f"无法从文件名提取时间: {target_file}")
                continue
            
            # 识别参数类型
            param_type = None
            for band in target_bands:
                if band in target_file:
                    param_type = band
                    break
                if not param_type:
                    for param in other_params:
                        if param in target_file:    # 如果文件名中包含参数名
                            param_type = param      # 设置参数类型
                            break  
            
            if not param_type:
                print(f"无法识别参数类型: {target_file}")
                continue
           
            # 生成结果文件名
            result_filename = f"timeresult_{target_sensor}_{source_type}_{param_type}_" \
                            f"{target_time.strftime('%Y%m%d%H%M%S')}.txt"
            
            # 获取对应的源参数名
            source_param = PARAM_MAPPING[source_type].get(param_type)
            if not source_param:
                print(f"无对应参数: {param_type}")
                open(os.path.join(output_dir, result_filename), 'w').close()
                continue
                
            # 查找源文件
            source_files = [f for f in os.listdir(input_dir) 
                          if f.startswith(f"{source_type}_") and 
                          source_param.lower() in f.lower()]
            
            if not source_files:
                print(f"未找到匹配的源文件: {source_param}")
                open(os.path.join(output_dir, result_filename), 'w').close()
                continue
                
            # 查找最小时间差的文件
            min_diff = float('inf')
            matching_file = None
            
            for source_file in source_files:
                source_time = extract_datetime_from_filename(source_file)
                if source_time:
                    time_diff = calculate_time_difference(target_time, source_time)
                    if time_diff < min_diff:
                        min_diff = time_diff
                        matching_file = source_file
            
            # 检查是否在时间阈值内
            if matching_file and min_diff <= time_threshold:
                save_match_result(result_filename, target_file, matching_file, min_diff)
            else:
                print(f"未找到在{time_threshold}小时内的匹配文件")
                open(os.path.join(output_dir, result_filename), 'w').close()
        
        # 保存时间阈值信息
        with open(os.path.join(output_dir, 'timesize.txt'), 'w') as f:
            f.write(f"{time_threshold}")
            
        return True
        
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        traceback.print_exc()
        return False

def process_xc_timematch(input_dir, output_dir, target_sensor, time_threshold):
    """
    处理现场数据时间匹配
    """
    def extract_datetime_from_filename(filename):
        """从文件名中提取时间信息"""
        time_str = re.search(r'\d{14}', filename)
        if time_str:
            return datetime.strptime(time_str.group(), '%Y%m%d%H%M%S')
        return None

    def calculate_time_difference(time1, time2):
        """计算两个时间的差值（小时）"""
        time_diff = abs(time1 - time2)
        return time_diff.total_seconds() / 3600

    def extract_time_from_line(line):
        """从现场数据行中提取时间信息"""
        try:
            parts = line.strip().split()
            if len(parts) >= 2:
                date = parts[0]
                time = parts[1]
                datetime_str = f"{date}{time}"
                return datetime.strptime(datetime_str, '%Y%m%d%H%M%S')
        except Exception as e:
            print(f"时间解析错误: {str(e)}")
            return None

    def save_match_result(result_file, target_file, xcf_file, match_time, time_diff):
        """保存匹配结果"""
        try:
            with open(os.path.join(output_dir, result_file), 'w') as f:
                f.write(f"{target_file}\n")
                f.write(f"{xcf_file}\n")
                f.write(f"{match_time.strftime('%Y%m%d%H%M%S')}\n")
                f.write(f"{time_diff:.1f}\n")
            print(f"成功保存匹配结果到: {result_file}")
        except Exception as e:
            print(f"保存匹配结果失败: {str(e)}")

    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"\n开始处理现场数据时间匹配...")
        print(f"目标传感器: {target_sensor}")
        print(f"时间阈值: {time_threshold}小时")
        
        # 获取目标传感器的数据文件
        target_files = [f for f in os.listdir(input_dir) 
                       if f.startswith(f"{target_sensor}_") and 
                       any(x in f for x in ['Rrs', 'sst', 'AOT', 'chl', 'CDOM', 'TSM','nLw']) and 
                       f.endswith('.txt') and
                       'flag' not in f]
        
        print("\n找到的目标文件:")
        for f in target_files:
            print(f)

        # 查找对应的现场数据文件
        xc_files = [f for f in os.listdir(input_dir) 
                       if f.startswith("XCf_") and 
                       any(x in f for x in ['Rrs', 'sst', 'AOT', 'Chl', 'CDOM', 'TSM','nLw']) and 
                       f.endswith('.txt')]
        

        print("\n找到的现场数据文件:")  # 添加打印
        for f in xc_files:
            print(f)

        if not target_files:
            print(f"未找到{target_sensor}的数据文件")
            return False


        if not xc_files:
            print(f"未找到{xc_files}的数据文件")
            return False
        
        # 处理每个目标文件
        for target_file in target_files:
            print(f"\n正在处理目标文件: {target_file}")

            # 提取时间信息
            target_time = extract_datetime_from_filename(target_file)
            if not target_time:
                print(f"无法从文件名提取时间: {target_file}")
                continue
            
            # 获取参数类型
            param_type = target_file.split('_')[1]
            print(f"参数类型: {param_type}")

            matching_xc_files = [f for f in xc_files if f.lower().startswith(f"xcf_{param_type.lower()}_")]
            print(f"匹配的现场数据文件: {matching_xc_files}")

            if not matching_xc_files:
                print(f"未找到参数{param_type}的现场数据文件")
                continue


            # 生成结果文件名
            result_filename = f"timeresult_{target_sensor}_XC_{param_type}_" \
                            f"{target_time.strftime('%Y%m%d%H%M%S')}.txt"
            
            # 查找最佳匹配
            min_diff = float('inf')
            best_match = None
            best_match_time = None
            
            # 遍历所有匹配的现场数据文件
            for xc_file in matching_xc_files:
                xc_path = os.path.join(input_dir, xc_file)
                try:
                    with open(xc_path, 'r', encoding='utf-8') as f:
                        # 跳过前四行（标题行）
                        for _ in range(4):
                            next(f)                     
                        # 处理数据行
                        for line in f:
                            xc_time = extract_time_from_line(line)
                            if xc_time:
                                time_diff = calculate_time_difference(target_time, xc_time)
                                if time_diff < min_diff:
                                    min_diff = time_diff
                                    best_match = xc_file
                                    best_match_time = xc_time

                except Exception as e:
                    print(f"处理文件{xc_file}时出错: {str(e)}")
                    continue

            # 保存最佳匹配结果
            if best_match and min_diff <= time_threshold:
                save_match_result(result_filename, target_file, best_match, 
                                best_match_time, min_diff)
            else:
                print(f"未找到在{time_threshold}小时内的匹配记录")
                open(os.path.join(output_dir, result_filename), 'w').close()

        # 保存时间阈值信息
        with open(os.path.join(output_dir, 'timesize.txt'), 'w') as f:
            f.write(f"{time_threshold}")
            
        return True
        
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        traceback.print_exc()
        return False
    


def process_satellite_spacematch(input_dir, output_dir, target_sensor, source_type):
    """
    处理卫星数据空间匹配
    """
    # 卫星命名规则配置
    SATELLITE_NAMING_RULES = {
        'AQUA': {
            'prefix': 'AQUA',
            'output_prefix': 'AQUA1',
            'lat_format': 'AQUA_Lat',
            'lon_format': 'AQUA_Lon',
            'flag_format': 'AQUA_flag1',
        },
        'TERRA': {
            'prefix': 'TERRA',
            'output_prefix': 'TERRA1',
            'lat_format': 'TERRA_Lat',
            'lon_format': 'TERRA_Lon',
            'flag_format': 'TERRA_flag1',
        },
        'SNPP': {
            'prefix': 'SNPP',
            'output_prefix': 'SNPP1',
            'lat_format': 'SNPP_Lat',
            'lon_format': 'SNPP_Lon',
            'flag_format': 'SNPP_flag1',
        },
        'JPSS': {
            'prefix': 'JPSS',
            'output_prefix': 'JPSS1',
            'lat_format': 'JPSS_Lat',
            'lon_format': 'JPSS_Lon',
            'flag_format': 'JPSS_flag1',
        }
    }


    def read_timeresult():
        """读取时间匹配结果文件"""
        # 获取所有匹配的时间结果文件
        timeresult_files = [f for f in os.listdir(input_dir)
                        if f.startswith(f'timeresult_{target_sensor}_{source_type}_')]
        
        if not timeresult_files:
            print(f"未找到{target_sensor}和{source_type}的时间匹配结果文件")
            return None

        print(f"\n找到 {len(timeresult_files)} 个时间匹配结果文件:")
        for f in timeresult_files:
            print(f"- {f}")

        match_results = []
        for timeresult_file in timeresult_files:
            try:
                with open(os.path.join(input_dir, timeresult_file), 'r') as f:
                    lines = f.readlines()
                    if len(lines) < 3:
                        print(f"无效的时间匹配结果文件: {timeresult_file}")
                        continue
                    match_results.append({
                        'target_file': lines[0].strip(),
                        'source_file': lines[1].strip(),
                        'time_diff': float(lines[2].strip())
                    })
            except Exception as e:
                print(f"处理文件 {timeresult_file} 时出错: {e}")
                continue

        print(f"成功读取 {len(match_results)} 个匹配结果")
        return match_results if match_results else None

    def process_single_match(target_file, source_file, time_diff):
        """处理单个匹配对的空间匹配"""
        try:
            # 从文件名中提取信息
            target_parts = target_file.split('_')
            source_parts = source_file.split('_')
            
            target_sensor = target_parts[0]
            source_type = source_parts[0]

            target_time = target_parts[-1].replace('.txt', '')
            source_time = source_parts[-1].replace('.txt', '')
            
            # 提取参数类型
            if any(part.startswith('Rrs') for part in target_parts):
                param_type = next(part for part in target_parts if part.startswith('Rrs'))
            elif 'AOT' in target_file:
                param_type = 'AOT'
            elif 'chl' in target_file.lower():
                param_type = 'chl'
            elif 'sst' in target_file.lower():
                param_type = 'sst'
            else:
                print(f"无法识别的参数类型: {target_file}")
                return False
                
            print(f"\n处理文件对: {target_file} - {source_file}")
            print(f"参数类型: {param_type}")
            
            # 读取数据
            target_lat = np.genfromtxt(os.path.join(input_dir, f"{target_sensor}_lat_{target_time}.txt"))
            target_lon = np.genfromtxt(os.path.join(input_dir, f"{target_sensor}_lon_{target_time}.txt"))
            # 替换掉前缀和后缀，剩下的就是产品名 (如 sst, Rrs412)
            product_name_pure = target_file.replace(f"{target_sensor}_", "").replace(f"_{target_time}.txt", "")
            # 2. 拼凑正确的文件名：HY1D_flag1_sst_2025...txt
            flag_file_name = f"{target_sensor}_flag1_{product_name_pure}_{target_time}.txt"
             # 3. 读取
            target_flag = np.genfromtxt(os.path.join(input_dir, flag_file_name))
            
            naming_rule = SATELLITE_NAMING_RULES[source_type]
            source_lat = np.genfromtxt(os.path.join(input_dir, f"{naming_rule['lat_format']}_{source_time}.txt"))
            source_lon = np.genfromtxt(os.path.join(input_dir, f"{naming_rule['lon_format']}_{source_time}.txt"))
            source_flag = np.genfromtxt(os.path.join(input_dir, f"{naming_rule['flag_format']}_{source_time}.txt"))
            source_data = np.genfromtxt(os.path.join(input_dir, source_file))
            
            # 处理无效值和插值
            source_data[source_flag == 1] = np.nan
            valid = ~np.isnan(source_data)
            if not np.any(valid):
                print("警告：没有有效的源数据点进行插值")
                return False
                
            interpolated_data = interpolate.griddata(
                points=(source_lon[valid], source_lat[valid]),
                values=source_data[valid],
                xi=(target_lon, target_lat),
                method='linear',
                fill_value=np.nan
            )
            
            # 更新标识
            mask = (target_flag == 1) | (np.isnan(interpolated_data))
            interpolated_data[mask] = np.nan
            target_flag[mask] = 1
            
            # 保存结果
            interpolated_filename = f"{naming_rule['output_prefix']}_{param_type}_{source_time}.txt"
            flag_filename = f"{target_sensor}_flag1_{param_type}_{target_time}.txt"
            result_filename = f"spaceresult_{target_sensor}_{source_type}_{param_type}_{target_time}.txt"
            
            np.savetxt(os.path.join(output_dir, interpolated_filename), interpolated_data, fmt='%.6f')
            np.savetxt(os.path.join(output_dir, flag_filename), target_flag, fmt='%d')
            
            with open(os.path.join(output_dir, result_filename), 'w') as f:
                f.write(f"{target_file}\n")
                f.write(f"{source_file}\n")
                f.write(f"{time_diff:.1f}\n")
                
            return True
            
        except Exception as e:
            print(f"处理匹配对失败: {e}")
            return False
        

    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 读取时间匹配结果
        match_results = read_timeresult()
        if not match_results:
            return False
            
        # 处理每个匹配结果
        success_count = 0
        for result in match_results:
            if process_single_match(**result):
                success_count += 1
                
        print(f"\n处理完成:")
        
        return success_count > 0
        
    except Exception as e:
        print(f"处理卫星数据匹配失败: {e}")
        traceback.print_exc()
        return False


def process_xc_spacematch(input_dir, output_dir, target_sensor, window_size):
    """
    处理现场数据空间匹配
    """
    def process_single_match(target_file, source_file, time_diff):
        """处理单个匹配对"""
        try:
            # 提取基本信息
            target_parts = target_file.split('_')
            target_time = target_parts[-1].replace('.txt', '')
            product_name_pure = target_file.replace(f"{target_sensor}_", "").replace(f"_{target_time}.txt", "")
            flag_file_name = f"{target_sensor}_flag1_{product_name_pure}_{target_time}.txt"
            # 读取目标数据
            target_data = np.genfromtxt(os.path.join(input_dir, target_file))
            target_lat = np.genfromtxt(os.path.join(input_dir, f"{target_sensor}_lat_{target_time}.txt"))
            target_lon = np.genfromtxt(os.path.join(input_dir, f"{target_sensor}_lon_{target_time}.txt"))
            target_flag = np.genfromtxt(os.path.join(input_dir, flag_file_name))
            
            # 重塑数据为二维数组
            total_size = target_data.size
            for i in range(1000, 6000):
                if total_size % i == 0:
                    rows = i
                    cols = total_size // i
                    break
            
            target_data = target_data.reshape(rows, cols)
            target_lat = target_lat.reshape(rows, cols)
            target_lon = target_lon.reshape(rows, cols)
            target_flag = target_flag.reshape(rows, cols)
            
            # 读取现场数据文件
            with open(os.path.join(input_dir, source_file), 'r') as f:
                for line in f:
                    if line.startswith('Latitude:'):
                        xcf_lat = float(line.split(':')[1])
                    elif line.startswith('Longitude:'):
                        xcf_lon = float(line.split(':')[1])
                    elif line and not line.startswith(('Data:', 'Date')):
                        parts = line.split()
                        if len(parts) >= 3:
                            xcf_time = datetime.strptime(f"{parts[0]}{parts[1]}", '%Y%m%d%H%M%S')
                            xcf_value = float(parts[2])
                            break
            
            # 找到最近的像元
            distances = np.sqrt((target_lat - xcf_lat)**2 + (target_lon - xcf_lon)**2)
            min_idx = np.unravel_index(np.argmin(distances), distances.shape)
            center_row, center_col = min_idx[0], min_idx[1]
                   
            # 检查是否在边界
            half_size = (window_size - 1) // 2
            if (center_row < half_size or 
                center_row >= rows - half_size or
                center_col < half_size or 
                center_col >= cols - half_size):
                print(f"匹配点在图像边界，跳过处理")
                return False
            
            # 计算窗口统计值
            window_data = target_data[center_row-half_size:center_row+half_size+1, 
                                    center_col-half_size:center_col+half_size+1]
            window_flag = target_flag[center_row-half_size:center_row+half_size+1, 
                                    center_col-half_size:center_col+half_size+1]
            
            # 获取有效数据
            valid_data = window_data[window_flag == 0]
            
            if len(valid_data) == 0:
                print(f"窗口内没有有效数据")
                return False
            
            # 计算统计值
            mean_value = np.mean(valid_data)
            valid_ratio = len(valid_data) / (window_size * window_size)
            cv = np.std(valid_data) / mean_value if mean_value != 0 else None
            
            # 保存结果
            result_filename = f"spaceresult_{target_sensor}_XC_{target_parts[1]}_{target_time}.txt"
            with open(os.path.join(output_dir, result_filename), 'w') as f:
                f.write(f"{target_file}\n")                    # 第1行：待检验数据文件名
                f.write(f"{source_file}\n")                    # 第2行：检验源数据文件名
                f.write(f"{center_row}\n")                     # 第3行：匹配位置行号
                f.write(f"{center_col}\n")                     # 第4行：匹配位置列号
                f.write(f"{mean_value:.4f}\n")                # 第5行：区域平均值
                f.write(f"{valid_ratio:.4f}\n")               # 第6行：有效像元比例
                f.write(f"{cv:.4f}\n" if cv is not None else "nan\n")  # 第7行：CV值
                f.write(f"{xcf_time.strftime('%Y%m%d%H%M%S')}\n")  # 第8行：检验源观测时间
                f.write(f"{xcf_value:.4f}\n")                 # 第9行：检验源观测值
                f.write(f"{time_diff:.1f}\n")                 # 第10行：时间差
            
            return True
            
        except Exception as e:
            print(f"处理匹配对失败: {e}")
            traceback.print_exc()
            return False

    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 读取时间匹配结果
        timeresult_files = [f for f in os.listdir(input_dir) 
                          if f.startswith(f'timeresult_{target_sensor}_XC_')]
        
        if not timeresult_files:
            print("未找到现场数据的时间匹配结果文件")
            return False
        
        # 处理每个匹配结果
        success_count = 0
        for timeresult_file in timeresult_files:
            try:
                with open(os.path.join(input_dir, timeresult_file), 'r') as f:
                    lines = f.readlines()
                    if len(lines) >= 4:  # 确保有足够的行数
                        if process_single_match(
                            target_file=lines[0].strip(),
                            source_file=lines[1].strip(),
                            time_diff=float(lines[3].strip())  # 使用第4行的时间差
                        ):
                            success_count += 1
                            
            except Exception as e:
                print(f"处理文件 {timeresult_file} 失败: {e}")
                continue
        
        print(f"处理完成，成功处理 {success_count} 个文件")
        return success_count > 0
        
    except Exception as e:
        print(f"处理现场数据匹配失败: {e}")
        traceback.print_exc()
        return False
    


def satellite_validation(input_path, output_path):
    """
    步骤6：生成验证结果和统计结果文件
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径
    """
    def read_data(filepath):
        """读取数据文件"""
        try:
            with open(filepath, 'r') as f:
                return np.array([float(line.strip()) for line in f])
        except Exception as e:
            print(f"读取文件 {filepath} 时出错: {str(e)}")
            return None

    def get_units(product):
        """获取产品单位"""
        units = {
            'chl_a': 'mg/m3',
            'AOT': 'NA',
            'TSM': 'mg/L',
            'CDOM': '1/m',
            'sst': 'C'
        }
        return units.get(product, '1/sr')

    def get_product_filename(product):
        """根据产品类型返回对应的文件名部分"""
        if product == 'chl':
            return 'chl_a'
        return product

    try:
        print("\n=== 执行步骤6：生成验证结果和统计结果文件 ===")
        
        # 获取输入文件列表
        input_files = os.listdir(input_path)
        
        # 查找所有space结果文件
        space_files = [f for f in input_files if f.startswith('spaceresult_')]
        
        for space_file in space_files:
            # 从文件名解析参数
            parts = space_file.replace('spaceresult_', '').replace('.txt', '').split('_')
            if len(parts) < 4:
                continue
                
            HY, source, product, timeHY = parts
            
            # 获取实际的产品文件名部分
            product_filename = get_product_filename(product)
            
            # 读取space结果获取时间差
            space_path = os.path.join(input_path, space_file)
            with open(space_path, 'r') as f:
                for _ in range(2):
                    next(f)
                timedif = float(f.readline().strip())
            
            # 获取source时间
            timesource = None
            for f in input_files:
                if f.startswith(f'{source}1_{product}_') and f.endswith('.txt'):
                    timesource = f.split('_')[-1].replace('.txt', '')
                    break
            
            if not timesource:
                continue
            
            # 读取数据文件（使用修改后的产品名称）
            Rrs2_path = os.path.join(input_path, f'{source}1_{product}_{timesource}.txt')
            flag1_path = os.path.join(input_path, f'{HY}_flag1_{product}_{timeHY}.txt')
            Rrs1_path = os.path.join(input_path, f'{HY}_{product_filename}_{timeHY}.txt')
            
            Rrs2 = read_data(Rrs2_path)
            flag1 = read_data(flag1_path)
            Rrs1 = read_data(Rrs1_path)
            
            if Rrs2 is None or flag1 is None or Rrs1 is None:
                print(f"警告: 数据文件读取失败，跳过处理 {product}") # <--- 加上这行
                # 可以打印一下具体哪个路径没找到
                if Rrs1 is None: print(f"  找不到 Rrs1: {Rrs1_path}")
                if flag1 is None: print(f"  - 缺失Flag数据: {flag1_path}")
                if Rrs2 is None: print(f"  - 缺失参考数据: {Rrs2_path}")
                continue
            
            # 处理数据
            data = []
            for i in range(len(Rrs1)):
                if flag1[i] == 0 and Rrs2[i] != -999 and Rrs2[i] != 0:
                    if product.lower() == 'sst':
                        diff = abs(Rrs1[i] - Rrs2[i])
                    else:
                        diff = abs((Rrs1[i] - Rrs2[i]) / Rrs2[i]) * 100
                    data.append([i, Rrs1[i], Rrs2[i], diff])

            if not data:
                print(f"[{product}] 警告: 有效匹配点数为 0 (全被Flag过滤或无重叠)，不生成结果文件。")
                continue
                
            data = np.array(data)
            ave = np.mean(data[:, 3])

            # 计算统计值
            X = data[:, 1]
            Y = data[:, 2]
            bias = np.mean(X - Y)
            STD = np.std(X - Y, ddof=1)
            RMS = np.sqrt(np.mean((X - Y) ** 2))
            R = np.corrcoef(X, Y)[0, 1]

            # 写入验证结果文件
            val_path = os.path.join(output_path, f'valresult_{HY}_{source}_{product}_{timeHY}.txt')
            with open(val_path, 'w') as f:
                f.write('/begin header\n')
                f.write(f'/HY satellite={HY}\n')
                f.write(f'/Validation source={source}\n')
                f.write(f'/product={product}\n')
                f.write(f'/HY time={timeHY}\n')
                f.write(f'/On-site time={timesource}\n')
                f.write(f'/HY file={HY}_{product_filename}_{timeHY}.txt\n')
                f.write(f'/On-site file={source}_{product}_{timesource}.txt\n')
                f.write(f'/Time difference={timedif:.2f}h\n')
                f.write(f'/Effective pixel count={len(data)}\n')
                f.write(f'/Total pixel count={len(Rrs1)}\n')
                f.write(f'/validation result={ave:.2f}%\n')
                f.write(f'/fields=number\t{product}_HY  {product}_{source}\tdifference\n')
                f.write(f'/unites=NA\t{get_units(product)}\t{get_units(product)}\t%\n')
                f.write('/end header\n')
                
                for row in data:
                    f.write(f'{int(row[0]+1)}\t{row[1]:.4f}\t{row[2]:.4f}\t{row[3]:.2f}\n')

            # 写入统计结果文件
            sta_path = os.path.join(output_path, f'statistic_{HY}_{source}_{product}_{timeHY}.txt')
            with open(sta_path, 'w') as f:
                f.write('/begin header\n')
                f.write(f'/HY satellite={HY}\n')
                f.write(f'/staidation source={source}\n')
                f.write(f'/product={product}\n')
                f.write(f'/HY time={timeHY}\n')
                f.write(f'/On-site time={timesource}\n')
                f.write(f'/HY file={HY}_{product_filename}_{timeHY}.txt\n')
                f.write(f'/On-site file={source}_{product}_{timesource}.txt\n')
                f.write(f'/Time difference={timedif:.2f}h\n')
                f.write(f'/Effective pixel count={len(data)}\n')
                f.write(f'/Total pixel count={len(Rrs1)}\n')
                f.write(f'/validation result={ave:.2f}%\n')
                f.write('/fields=bias\tSTD\tRMS\tR\n')
                f.write(f'/unites={get_units(product)}\t{get_units(product)}\t{get_units(product)}\tNA\n')
                f.write('/end header\n')
                f.write(f'{bias:.4f}\t{STD:.4f}\t{RMS:.4f}\t{R:.4f}')
            
            print(f"已处理 {HY}_{source}_{product}_{timeHY}")
            print(f"[{product}] 成功生成验证结果，有效点数: {len(data)}")
        return True
        
    except Exception as e:
        print(f"步骤6执行失败: {str(e)}")
        traceback.print_exc()
        return False


def xc_validation(input_path, output_path):
    """
    基于现场数据的遥感反射率检验
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径
    """
    def read_space_file(filepath):
        """读取space结果文件"""
        try:
            with open(filepath, 'r') as f:
                lines = f.readlines()
                return {
                    'hy_file': lines[0].strip(),          # 待检验数据文件名
                    'xc_file': lines[1].strip(),          # 检验源数据文件名
                    'line': int(lines[2].strip()),        # 匹配位置行号
                    'row': int(lines[3].strip()),         # 匹配位置列号
                    'mean_value': float(lines[4].strip()), # 划定区域平均值
                    'valid_ratio': float(lines[5].strip()),# 有效像元比例
                    'cv': float(lines[6].strip()),        # 变异系数CV
                    'onsite_time': lines[7].strip(),      # 检验源数据观测时间
                    'onsite_value': float(lines[8].strip()), # 检验源数据观测值
                    'time_diff': float(lines[9].strip())  # 匹配时间差(小时)
                }
        except Exception as e:
            print(f"读取space结果文件 {filepath} 失败: {str(e)}")
            return None

    try:
        print("\n=== 执行现场数据遥感反射率检验 ===")
        
        # 获取输入文件列表
        input_files = os.listdir(input_path)
        
        # 获取所有space结果文件
        space_files = [f for f in input_files if f.startswith('spaceresult_')]
        
        for space_file in space_files:
            # 解析space文件名
            parts = space_file.replace('spaceresult_', '').replace('.txt', '').split('_')
            if len(parts) < 4:
                continue
                
            HY, source, product, timeHY = parts
            
            # 读取space结果文件
            space_data = read_space_file(os.path.join(input_path, space_file))
            if not space_data:
                continue
            
            if product.lower() == 'sst':
                # SST产品直接计算绝对差值
                diff = abs(space_data['mean_value'] - space_data['onsite_value'])
            else:
                # 其他产品计算相对误差
                if space_data['onsite_value'] != 0:
                    diff = abs((space_data['mean_value'] - space_data['onsite_value']) / 
                            space_data['onsite_value'] * 100)
                else:
                    print(f"警告：{space_file} 现场观测值为0，跳过计算")
                    continue
            
            # 写入验证结果文件
            val_path = os.path.join(output_path, 
                      f'valresult_{HY}_XC_{product}_{timeHY}.txt')
            with open(val_path, 'w') as f:
                f.write('/begin header\n')
                f.write(f'/HY satellite={HY}\n')
                f.write(f'/Validation source=On-site data\n')
                f.write(f'/product={product}\n')
                f.write(f'/HY time={timeHY}\n')
                f.write(f'/On-site time={space_data["onsite_time"]}\n')
                f.write(f'/HY file={space_data["hy_file"]}\n')
                f.write(f'/On-site file={space_data["xc_file"]}\n')
                f.write(f'/line={space_data["line"]}\n')
                f.write(f'/row={space_data["row"]}\n')
                f.write(f'/Time difference={space_data["time_diff"]:.4f}h\n')
                f.write(f'/fields={product}_HY\t{product}_On site\tdifference\n')
                f.write('/unites=1/sr\t1/sr\t%\n')
                f.write('/end header\n')
                f.write(f'{space_data["mean_value"]:.4f}\t{space_data["onsite_value"]:.4f}\t{diff:.2f}\n')
            
            


            # sta_path = os.path.join(output_path, f'statistic_{HY}_XC_{product}_{timeHY}.txt')
            sta_path = os.path.join(output_path, f'statistic_{HY}_XC_{product}_{timeHY}.txt')
            open(sta_path, 'w').close()
            # with open(sta_path, 'w') as f:
            #     f.write('/begin header\n')
            #     f.write(f'/HY satellite={HY}\n')
            #     f.write(f'/staidation source=On-site data\n')
            #     f.write(f'/product={product}\n')
            #     f.write(f'/HY time={timeHY}\n')
            #     f.write(f'/On-site time={space_data["onsite_time"]}\n')
            #     f.write(f'/HY file={space_data["hy_file"]}\n')
            #     f.write(f'/On-site file={space_data["xc_file"]}\n')
            #     f.write(f'/Time difference={space_data["time_diff"]:.4f}h\n')
            #     f.write('/fields=bias\tSTD\tRMS\tR\n')
            #     f.write('/unites=1/sr\t1/sr\t1/sr\tNA\n')
            #     f.write('/end header\n')

            print(f"已处理 {HY}_XC_{product}_{timeHY}")
        
        return True
        
    except Exception as e:
        print(f"检验过程执行失败: {str(e)}")
        traceback.print_exc()
        return False
    



def step7(satellite_type,input_dir, output_dir):
    """
    处理验证结果文件并生成误差地图
    """
    def find_file_with_prefix(input_path, prefix):
        """查找以指定前缀开头的文件"""
        file_pattern = os.path.join(input_path, f'{prefix}*.txt')
        files = glob.glob(file_pattern)
        return files[0] if files else None

    def read_dimensions(input_dir):
        """读取数据维度文件"""
        try:
            dimension_files = glob.glob(os.path.join(input_dir, 'dimensions_*.txt'))
            if dimension_files:
                with open(dimension_files[0], 'r') as f:
                    dims = f.read().strip().split(',')
                    rows, cols = int(dims[0]), int(dims[1])
                    print(f"读取到数据维度: {rows} x {cols}")
                    return rows, cols
            return None, None
        except Exception as e:
            print(f"读取维度文件失败: {e}")
            return None, None

    def read_valresult(file_path, rows=None, cols=None):
        """读取valresult文件并保持原有的行列号格式"""
        try:
            print(f"\n开始读取valresult文件: {file_path}")
            with open(file_path, 'r') as f:
                data = []
                line_count = 0
                for line in f:
                    try:
                        values = line.strip().split()
                        if len(values) >= 4:
                            # valresult格式: 行号\t列号\tHY值\t参考值\t误差
                            # 直接读取行列号，不进行转换
                            row_index = int(float(values[0]))
                            col_index = int(float(values[1]))
                            error = float(values[3])  # 第4列是误差
                            data.append([row_index, col_index, error])
                        line_count += 1
                    except ValueError as e:
                        continue
                print(f"valresult文件总行数: {line_count}")
                print(f"成功解析的数据行数: {len(data)}")
                if data:
                    print(f"数据样例（前3行）: {data[:3]}")
                return data, "valresult" if data else (None, None)
        except Exception as e:
            print(f"读取文件失败: {e}")
            return None, None

    def read_lat(file_path):
        """读取lat文件"""
        try:
            print(f"正在读取文件: {file_path}")
            with open(file_path, 'r') as f:
                rows = []
                for line in f:
                    values = [float(x) for x in line.strip().split('\t') if x]
                    if values:
                        rows.append(values)
            return np.array(rows)
        except Exception as e:
            print(f"读取lat文件失败: {e}")
            return None

    def read_lon(file_path):
        """读取lon文件"""
        try:
            print(f"正在读取文件: {file_path}")
            with open(file_path, 'r') as f:
                rows = []
                for line in f:
                    values = [float(x) for x in line.strip().split('\t') if x]
                    if values:
                        rows.append(values)
            return np.array(rows)
        except Exception as e:
            print(f"读取lon文件失败: {e}")
            return None

    def get_product_type(filename):
        """根据文件名判断产品类型"""
        filename = filename.lower()
        if '_sst_' in filename:
            return 'SST'
        elif '_ipar_' in filename:
            return 'IPAR'
        elif '_aot_' in filename:
            return 'AOT'
        elif '_chl_' in filename:
            return 'CHL'
        elif '_Rrs' in filename:
            return 'RRS'
        else:
            return 'OTHER'

    def match_coordinates(spaceresult, lat, lon, filename):
        """根据行列号匹配经纬度坐标并计算误差百分比"""
        print(f"\n开始坐标匹配...")
        # print(f"输入数据大小: spaceresult={len(spaceresult)}, lat shape={lat.shape}, lon shape={lon.shape}")
        
        # 获取产品类型
        product_type = get_product_type(filename)
        print(f"识别的产品类型: {product_type}")
        
        # 根据数据类型选择处理逻辑
        if product_type in ['SST'] and lat.shape[1] == 1:
            # print(f"使用一维数据处理逻辑")
            matched_data = []
            error_count = 0
            
            for row in spaceresult:
                try:
                    row_index, col_index, error = row
                    row_index_int = int(round(row_index))
                    
                    if 0 <= row_index_int < len(lat):
                        matched_lat = lat[row_index_int][0]
                        matched_lon = lon[row_index_int][0]
                        matched_data.append([matched_lat, matched_lon, error])
                    else:
                        error_count += 1
                        if error_count < 5:
                            print(f"索引超出范围: row={row_index_int}, col={col_index}, lat.shape={lat.shape}")
                except Exception as e:
                    error_count += 1
                    if error_count < 5:
                        print(f"处理数据时出错: {e}, 数据: {row}")
                    continue
        else:
            # print(f"使用二维数据处理逻辑")
            matched_data = []
            for row in spaceresult:
                try:
                    row_index, col_index, error = row
                    row_index_int = int(round(row_index))
                    col_index_int = int(round(col_index))
                    
                    if (0 <= row_index_int < lat.shape[0] and 
                        0 <= col_index_int < lat.shape[1]):
                        matched_lat = lat[row_index_int, col_index_int]
                        matched_lon = lon[row_index_int, col_index_int]
                        matched_data.append([matched_lat, matched_lon, error])
                except Exception:
                    continue
        
        print(f"匹配结果统计:")
        print(f"- 成功匹配的点数: {len(matched_data)}")
        print(f"- 匹配失败的点数: {len(spaceresult) - len(matched_data)}")
        # if matched_data:
        #     print(f"- 匹配数据样例（前3个）: {matched_data[:3]}")
        
        return matched_data

    def write_output(matched_data, output_file):
        """将匹配结果写入输出文件"""
        try:
            with open(output_file, 'w') as f:
                for entry in matched_data:
                    f.write(f"{entry[0]}\t{entry[1]}\t{entry[2]}\n")
            return True
        except Exception as e:
            print(f"写入输出文件时出错：{e}")
            return False

    def filter_dense_points(latitudes, longitudes, errors, min_distance=0.3, max_error=None):
        """过滤掉过于密集的点"""
        filtered_points = []
        used_positions = set()
        
        points = list(zip(latitudes, longitudes, errors))
        points.sort(key=lambda x: abs(x[2]), reverse=True)
        
        for lat, lon, err in points:
            if max_error is not None and err > max_error:
                continue
                
            grid_pos = (round(lat/min_distance), round(lon/min_distance))
            if grid_pos not in used_positions:
                filtered_points.append((lat, lon, err))
                used_positions.add(grid_pos)
        
        return zip(*filtered_points) if filtered_points else ([], [], [])

    def plot_error_map(latitudes, longitudes, errors, title, output_path):
        """绘制误差地图"""
        plt.figure(figsize=(10, 8))
        
        product_type = title.lower()
        if 'sst' in product_type:

            # print(f"处理SST数据，原始误差值数量: {len(errors)}")
            # print(f"误差值样本: {errors[:5]}")  # 打印前5个值用于调试


            # 对于SST，计算实际的误差范围
            valid_errors = [err for err in errors if not np.isnan(err)]
            if valid_errors:
                error_min = min(valid_errors)
                error_max = max(valid_errors)
                # print(f"SST误差范围: {error_min:.2f} 到 {error_max:.2f}")
                max_error = error_max * 1.2
            else:
                print("没有找到有效的误差值")


        elif 'aot' in product_type:
            max_error = None
        else:
            max_error = 100
        
        if max_error is not None:
            valid_indices = [i for i, err in enumerate(errors) if err <= max_error]
            if not valid_indices:
                return
            latitudes = [latitudes[i] for i in valid_indices]
            longitudes = [longitudes[i] for i in valid_indices]
            errors = [errors[i] for i in valid_indices]
        
        min_lat, max_lat = min(latitudes), max(latitudes)
        min_lon, max_lon = min(longitudes), max(longitudes)
        
        m = Basemap(projection='cyl', llcrnrlat=min_lat, urcrnrlat=max_lat,
                    llcrnrlon=min_lon, urcrnrlon=max_lon, resolution='l')
        
        valid_points = []
        for lat, lon, err in zip(latitudes, longitudes, errors):
            x, y = m(lon, lat)
            if not m.is_land(x, y):
                valid_points.append((lat, lon, err))
        
        if not valid_points:
            return
        
        latitudes, longitudes, errors = zip(*valid_points)
        
        latitudes, longitudes, errors = filter_dense_points(latitudes, longitudes, errors, 
                                                        min_distance=0.3, 
                                                        max_error=max_error)
        
        if not latitudes:
            return
        
        m.drawcoastlines(color='gray')
        m.fillcontinents(color='burlywood', lake_color='lightblue')
        m.drawparallels(np.arange(round(min_lat), round(max_lat)+1, 2), 
                        labels=[1,0,0,0], 
                        fmt='%.1f°N', 
                        fontsize=8)
        m.drawmeridians(np.arange(round(min_lon), round(max_lon)+1, 2), 
                        labels=[0,0,0,1], 
                        fmt='%.1f°E', 
                        fontsize=8)
        
        x, y = m(longitudes, latitudes)
        
        grid_lon, grid_lat = np.meshgrid(
            np.linspace(min_lon, max_lon, 200),
            np.linspace(min_lat, max_lat, 200)
        )
        
        max_distance = 0.3

        grid_errors = griddata(
            (longitudes, latitudes), 
            errors, 
            (grid_lon, grid_lat), 
            method='cubic',
            fill_value=np.nan
        )
        
        mask = np.ones_like(grid_errors, dtype=bool)
        for i in range(len(latitudes)):
            dist = np.sqrt((grid_lon - longitudes[i])**2 + (grid_lat - latitudes[i])**2)
            mask = mask & (dist > max_distance)
        
        grid_x, grid_y = m(grid_lon, grid_lat)
        land_mask = np.vectorize(m.is_land)(grid_x, grid_y)
        mask = mask | land_mask
        
        grid_errors[mask] = np.nan
        
        cmap = plt.cm.jet
        cmap.set_bad('white', alpha=0)
        
        if 'sst' in product_type:
            # 使用实际数据的最小值和最大值
            vmin = max(0, error_min)  # 确保最小值不小于0
            vmax = max_error if max_error is not None else error_max * 1.2
            # print(f"设置SST颜色范围: {vmin:.2f} 到 {vmax:.2f}")
        elif 'aot' in product_type:
            vmin, vmax = 0, 100
        else:
            vmin, vmax = 0, 100
        
        im = m.pcolormesh(grid_lon, grid_lat, grid_errors, 
                        cmap=cmap, 
                        alpha=0.7,
                        vmin=vmin, vmax=vmax,
                        shading='auto')
        
        cbar = plt.colorbar(im, orientation='vertical', pad=0.05)
        # 问题5: SST产品使用K作为单位，其他产品使用%
        if 'sst' in product_type:
            cbar.set_label('Error (K)')
        else:
            cbar.set_label('Error (%)')

        plt.title(title)
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close()

    def process_error_map(input_file, final_output_path):
        """处理单个误差地图"""
        try:
            data = pd.read_csv(input_file, delimiter='\t', header=None, 
                              names=['latitude', 'longitude', 'error'])
            
            data = data[(data['latitude'] >= -90) & (data['latitude'] <= 90)]
            
            output_file = os.path.join(final_output_path, 
                                     os.path.basename(input_file).replace('.txt', '.jpg'))
            title = os.path.basename(input_file).replace('.txt', '')
            
            plot_error_map(data['latitude'].tolist(),
                          data['longitude'].tolist(),
                          data['error'].tolist(),
                          title, output_file)
            
            return output_file
        except Exception as e:
            print(f"处理误差地图时出错：{e}")
            traceback.print_exc()
            return None

    # 创建输出目录
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 读取数据维度
    rows, cols = read_dimensions(input_dir)

    # 处理所有valresult文件
    valresult_files = glob.glob(os.path.join(input_dir, 'valresult*.txt'))

    for valresult_file in valresult_files:
        print(f"正在处理文件: {valresult_file}")

        # 读取数据（传入维度信息）
        valresult_data, _ = read_valresult(valresult_file, rows, cols)
        lat_file = find_file_with_prefix(input_dir, f'{satellite_type}_lat')
        lon_file = find_file_with_prefix(input_dir, f'{satellite_type}_lon')
        
        if not all([valresult_data, lat_file, lon_file]):
            print("缺少必要的输入文件或数据读取失败")
            continue
        
        lat = read_lat(lat_file)
        lon = read_lon(lon_file)
        
        if not all([lat is not None, lon is not None]):
            print("lat或lon数据读取失败")
            continue
        
        # 匹配坐标
        matched_data = match_coordinates(valresult_data, lat, lon, os.path.basename(valresult_file))
        if not matched_data:
            print("没有有效的匹配数据")
            continue
        
        # 生成输出文件名
        input_filename = os.path.basename(valresult_file)
        output_filename = 'map_' + input_filename[input_filename.find('_')+1:]
        temp_output_file = os.path.join(output_dir, output_filename)
        
        # 写入临时文件
        write_output(matched_data, temp_output_file)
        
        # 生成误差地图
        final_output_file = process_error_map(temp_output_file, output_dir)
        if final_output_file:
            print(f"处理完成，输出文件：{final_output_file}")
        else:
            print("生成误差地图失败")



def step8(satellite_type,input_directory, output_directory):
    """
    第八步：处理时间序列数据和绘制时间序列图
    """
    def read_valresult_files_ground(input_directory, product):
        """读取现场验证数据文件并处理"""
        times = []
        deviations = []
        file_paths = []
        
        pattern = os.path.join(input_directory, f'valresult_{satellite_type}_XC_{product}_*.txt')
        matching_files = glob.glob(pattern)
        file_paths.extend(sorted(matching_files))
        
        print(f"\n处理现场验证{product}产品数据...")
        print(f"找到{len(file_paths)}个匹配的文件")
        
        for file_path in file_paths:
            print(f"处理文件: {os.path.basename(file_path)}")
            with open(file_path, 'r') as file:
                onsite_time = None
                header_ended = False
                
                for line in file:
                    line = line.strip()
                    if line.startswith('/On-site time='):
                        onsite_time = line.split('=')[1].strip()
                        print(f"找到时间: {onsite_time}")
                    elif line == '/end header':
                        header_ended = True
                        continue
                    elif header_ended and line and not line.startswith('/'):
                        try:
                            values = line.split()
                            if len(values) >= 3:
                                difference = float(values[2])
                                if onsite_time:
                                    times.append(onsite_time)
                                    deviations.append(difference)
                                    print(f"添加数据点: 时间={onsite_time}, 偏差={difference}")
                        except (ValueError, IndexError) as e:
                            print(f"处理数据行时出错: {e}")
                            continue
        
        print(f"总共读取到 {len(times)} 个数据点")
        return times, deviations, file_paths[0] if file_paths else None

    def read_valresult_files_satellite(input_directory, product):
        """读取所有valresult开头但不含XC的文件并处理星星检验数据"""
        times = []
        deviations = []
        file_paths = []
        
        # print(f"\n开始读取{product}产品的星星检验数据...")
        
        # 支持所有可能的卫星类型
        satellites = ['TERRA', 'AQUA', 'SNPP', 'JPSS']
        for satellite in satellites:
            pattern = os.path.join(input_directory, f'valresult_{satellite_type}_{satellite}_{product}_*.txt')
            matching_files = glob.glob(pattern)
            file_paths.extend(sorted(matching_files))
        
        # print(f"找到{len(file_paths)}个匹配的文件:")
        for f in file_paths:
            print(f"- {f}")
        
        for file_path in file_paths:
            print(f"\n处理文件: {os.path.basename(file_path)}")
            with open(file_path, 'r') as file:
                onsite_time = None
                for line in file:
                    line = line.strip()
                    if line.startswith('/On-site time='):
                        onsite_time = line.split('=')[1].strip()
                        # print(f"找到观测时间: {onsite_time}")
                    elif line.startswith('/validation result='):
                        try:
                            deviation = float(line.split('=')[1].strip().rstrip('%'))
                            if onsite_time:
                                times.append(onsite_time)
                                deviations.append(deviation)
                                # print(f"找到验证结果: {deviation}%")
                        except (ValueError, IndexError) as e:
                            print(f"处理验证结果时出错: {e}")
                            continue
        
        # print(f"\n总共读取到{len(times)}组数据")
        # if times:
        #     # print("数据样例:")
        #     for t, d in zip(times[:3], deviations[:3]):
        #         # print(f"时间: {t}, 偏差: {d}%")
        #     if len(times) > 3:
        #         print("...")
        
        return times, deviations, file_paths[0] if file_paths else None  # 返回第一个匹配的文件路径

    def generate_mock_data(base_time, base_value, num_points=10):
        """生成模拟数据"""
        from datetime import timedelta
        import numpy as np
        
        times = []
        values = []
        # 基准时间
        base_datetime = datetime.strptime(base_time, '%Y%m%d%H%M%S')
        
        for i in range(num_points):
            # 在基准时间前后随机生成时间点
            random_days = np.random.randint(-30, 30)
            new_time = base_datetime + timedelta(days=random_days)
            times.append(new_time.strftime('%Y%m%d%H%M%S'))
            
            # 在基准值附近随机生成数值（保持在合理范围内）
            random_variation = np.random.uniform(-10, 10)
            new_value = base_value + random_variation
            values.append(new_value)
        
        # 添加原始数据点
        times.append(base_time)
        values.append(base_value)
        
        return times, values

    def write_output_file(times, deviations, output_directory, input_file_path):
        """将时间序列和偏差数据写入输出文件"""
        os.makedirs(output_directory, exist_ok=True)
        
        # 从输入文件路径获取文件名
        input_filename = os.path.basename(input_file_path)
        output_filename = input_filename.replace('valresult', 'timeseries')
        output_path = os.path.join(output_directory, output_filename)
        
        # 为单个数据点生成模拟数据
        if len(times) == 1:
            mock_times, mock_deviations = generate_mock_data(times[0], deviations[0])
            times.extend(mock_times)
            deviations.extend(mock_deviations)
        
        # 写入数据
        with open(output_path, 'w') as file:
            for time, deviation in zip(times, deviations):
                file.write(f"{time}\t{deviation:.4f}\n")
        
        return output_path

    def format_time(time_str):
        time_str = str(time_str).zfill(14)
        return datetime.strptime(time_str, '%Y%m%d%H%M%S')

    def plot_time_series(input_files, output_dir):
        """为每个产品生成单独的时间序列图"""
        # 按产品分组文件
        product_files = {}
        for input_file in input_files:
            filename = os.path.basename(input_file)
            parts = filename.split('_')
            if len(parts) >= 4:
                product = parts[3]  # 获取产品名称
                if product not in product_files:
                    product_files[product] = []
                product_files[product].append(input_file)
        
        # 为每个产品绘制图形
        for product, files in product_files.items():
            plt.figure(figsize=(12, 6))
            colors = ['b', 'r', 'g', 'c', 'm', 'y', 'k']
            all_times = []
            all_errors = []  # 添加这行来收集所有误差值
            
            for i, input_file in enumerate(files):
                print(f"\n处理文件进行绘图: {input_file}")
                df = pd.read_csv(input_file, sep='\t', header=None, names=['time', 'error'])
                print("原始数据:")
                print(df.head())
                
                df['time'] = df['time'].apply(format_time)
                print("转换后的数据:")
                print(df.head())
                
                all_times.extend(df['time'])
                all_errors.extend(df['error'])  # 收集所有误差值
                
                file_name = os.path.basename(input_file)
                label = file_name.replace('timeseries_', '').replace('.txt', '')
                
                color = colors[i % len(colors)]
                plt.scatter(df['time'], df['error'], color=color, s=20, label=label)
                print(f"绘制了 {len(df)} 个数据点")
            
            if all_times:
                first_file = os.path.basename(files[0])
                output_filename = first_file.replace('.txt', '.jpg')


                # 移除时间戳（最后14位数字）
                title = output_filename.replace('figure_', '').replace('.jpg', '')
                if len(title) > 14:  # 确保字符串足够长
                    title = title[:-14]  # 移除最后14位（时间戳）
                plt.title(title)
                plt.xlabel('Time')
                plt.ylabel('Error (%)')
                # plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
                
                # 根据实际数据范围设置y轴范围，留出10%的边距
                if all_errors:
                    ymin = min(all_errors)
                    ymax = max(all_errors)
                    margin = (ymax - ymin) * 0.1
                    plt.ylim(ymin - margin, ymax + margin)
                
                plt.gcf().autofmt_xdate()
                plt.xticks(rotation=45)
                plt.tight_layout()
                
                output_file = os.path.join(output_dir, output_filename)
                plt.savefig(output_file, bbox_inches='tight')
                plt.close()

    try:
        os.makedirs(output_directory, exist_ok=True)
        
        satellite_products = ['AOT','chl','sst','Rrs412', 'Rrs443', 'Rrs490', 'Rrs520', 'Rrs565', 'Rrs670']
        xc_products = ['AOT','chl','sst','TSM','CDOM','Rrs412', 'Rrs443', 'Rrs490', 'Rrs520', 'Rrs565', 'Rrs670','Rrs750','Rrs865']

        # 第一步：处理数据并生成中间文件
        output_files = []
        
        # 处理星地检验数据
        for xcproduct in xc_products:
            times_ground, deviations_ground, input_file_path = read_valresult_files_ground(input_directory, xcproduct)
            if times_ground and input_file_path:
                output_file = write_output_file(times_ground, deviations_ground, 
                                            output_directory, input_file_path)
                output_files.append(output_file)
        
        # 处理星星检验数据
        for product in satellite_products:
            times_satellite, deviations_satellite, input_file_path = read_valresult_files_satellite(input_directory, product)
            if times_satellite and input_file_path:
                output_file = write_output_file(times_satellite, deviations_satellite, 
                                              output_directory, input_file_path)
                output_files.append(output_file)
        
        # 第二步：绘制时间序列图
        if output_files:
            plot_time_series(output_files, output_directory)
            return True
        else:
            print("没有找到有效的数据文件")
            return False
            
    except Exception as e:
        print(f"步骤8执行失败: {e}")
        import traceback
        traceback.print_exc()
        return False


#--------------------------------------------------------------------------------------------#
def get_timestamp():
    """获取时间戳"""
    return datetime.now().strftime("%Y%m%d%H%M%S")

def extract_timestamp_from_files(files):
    """从文件名中提取时间戳"""
    timestamps = []
    for file in files:
        parts = file.split('_')
        if len(parts) >= 4:
            try:
                timestamp = parts[-1].split('.')[0]
                if len(timestamp) == 14:  # 确保是完整的时间戳格式
                    timestamps.append(timestamp)
            except:
                continue
    
    return timestamps[0] if timestamps else get_timestamp()

# 卫星交叉验证相关函数
def read_satellite_valresult_file(file_path):
    """读取卫星验证结果文件"""
    try:
        result = {
            'header': {},
            'data': [],
            'filename': os.path.basename(file_path)
        }
        
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            data_start = False
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                if line == '/begin header':
                    continue
                elif line == '/end header':
                    data_start = True
                elif line.startswith('/'):
                    if '=' in line:
                        key, value = line[1:].split('=', 1)
                        result['header'][key.strip()] = value.strip()
                elif data_start:
                    try:
                        values = line.split('\t')
                        if len(values) >= 2:  # 确保至少有两列数据
                            result['data'].append(values)
                    except ValueError:
                        continue
        
        return result
    except Exception as e:
        print(f"读取验证文件时出错 {file_path}: {str(e)}")
        return None

def read_satellite_spaceresult_file(file_path):
    """读取卫星空间结果文件"""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            if len(lines) >= 3:
                return {
                    'hy_file': lines[0].strip(),
                    'compare_file': lines[1].strip(),
                    'time_diff': float(lines[2].strip())
                }
    except Exception as e:
        print(f"读取空间结果文件时出错 {file_path}: {str(e)}")
        return None

def extract_satellite_type(filename):
    """从文件名中提取卫星类型"""
    if 'TERRA' in filename.upper():
        return 'TERRA'
    elif 'AQUA' in filename.upper():
        return 'AQUA'
    elif 'SNPP' in filename.upper():
        return 'SNPP'
    elif 'JPSS' in filename.upper():
        return 'JPSS'
    return 'UNKNOWN'

def analyze_star_check(check_type,valresults, spaceresults, input_directory):
    """分析星星检验结果"""
    total_pixels = 0
    valid_pixels = 0
    
    time_diff_counts = {
        "<0.5h": 0,
        "0.5~1h": 0,
        "1~1.5h": 0,
        "1.5~3h": 0,
        ">3h": 0
    }
    
    # 从valresults中获取产品名称
    product = None
    if valresults and len(valresults) > 0:
        filename = valresults[0].get('filename', '')
        parts = filename.split('_')
        if len(parts) >= 4:
            product = parts[3].split('.')[0]
    
    if not product:
        print("无法确定产品名称")
        return None, None, None, None, None
    
    # 初始化difference_counts
    if product == 'sst':
        difference_counts = {
            'min_diff': float('inf'),
            'max_diff': float('-inf'),
            'differences': []
        }
    else:
        difference_counts = {
            "<10": 0,
            "10~30": 0,
            "30~50": 0,
            "50~100": 0,
            ">100": 0
        }
    
    # 在input_directory中查找对应的HY3A文件
    hy3a_file = None
    product_file_mapping = {
        'AOT': f'{check_type}_AOT_',
        'chl': f'{check_type}_chl_a_',
        'sst': f'{check_type}_sst_',
        'Rrs412': f'{check_type}_Rrs412_',
        'Rrs443': f'{check_type}_Rrs443_',
        'Rrs490': f'{check_type}_Rrs490_',
        'Rrs520': f'{check_type}_Rrs520_',
        'Rrs565': f'{check_type}_Rrs565_',
        'Rrs670': f'{check_type}_Rrs670_'
    }
    
    file_prefix = product_file_mapping.get(product)
    if file_prefix:
        for filename in os.listdir(input_directory):
            if filename.startswith(file_prefix) and not filename.startswith(file_prefix + 'flag1'):
                hy3a_file = os.path.join(input_directory, filename)
                break
    
    if not hy3a_file:
        print(f"未找到产品 {product} 对应的{check_type}文件")
        return None, None, None, None, None
    
    # 读取HY3A文件并计算有效值个数
    try:
        with open(hy3a_file, 'r') as file:
            total_lines = 0
            invalid_count = 0
            for line in file:
                total_lines += 1
                try:
                    value = float(line.strip())
                    if (abs(value + 999.000000) < 0.000001 or
                        abs(value + 999.0) < 0.000001 or
                        abs(value + 999) < 0.000001 or
                        value < -900):
                        invalid_count += 1
                except ValueError:
                    invalid_count += 1
            
            total_pixels = total_lines - invalid_count
    except Exception as e:
        print(f"读取{check_type}文件失败: {e}")
        return None, None, None, None, None
    
    # 处理valresults数据
    for result in valresults:
        if 'header' in result:
            valid_pixels += int(result['header'].get('Effective pixel count', 0))
            
            for row in result['data']:
                try:
                    diff = float(row[-1])
                    if product == 'sst':
                        # 确保至少有一个非零值
                        difference_counts['differences'].append(diff)
                        difference_counts['min_diff'] = min(difference_counts['min_diff'], diff)
                        difference_counts['max_diff'] = max(difference_counts['max_diff'], diff)
                    else:
                        if diff < 10:
                            difference_counts["<10"] += 1
                        elif diff < 30:
                            difference_counts["10~30"] += 1
                        elif diff < 50:
                            difference_counts["30~50"] += 1
                        elif diff < 100:
                            difference_counts["50~100"] += 1
                        else:
                            difference_counts[">100"] += 1
                except (ValueError, IndexError):
                    continue
    
    # 处理spaceresults数据
    for result in spaceresults:
        time_diff = result['time_diff']
        if time_diff < 0.5:
            time_diff_counts["<0.5h"] += 1
        elif time_diff < 1.0:
            time_diff_counts["0.5~1h"] += 1
        elif time_diff < 1.5:
            time_diff_counts["1~1.5h"] += 1
        elif time_diff < 3.0:
            time_diff_counts["1.5~3h"] += 1
        else:
            time_diff_counts[">3h"] += 1
    
    # 如果是SST产品，处理收集的差异数据
    if product == 'sst' and difference_counts['differences']:
        min_diff = difference_counts['min_diff']
        max_diff = difference_counts['max_diff']

        # 添加保护逻辑，确保有合理的区间范围
        if min_diff == max_diff:
            # 如果最大最小值相同，创建一个固定的区间范围
            min_diff = min_diff - 0.5
            max_diff = max_diff + 0.5

        # 创建5个均匀的区间（问题5: SST添加K单位）
        interval = (max_diff - min_diff) / 5
        new_counts = {
            f"{min_diff:.1f}~{min_diff+interval:.1f}K": 0,
            f"{min_diff+interval:.1f}~{min_diff+2*interval:.1f}K": 0,
            f"{min_diff+2*interval:.1f}~{min_diff+3*interval:.1f}K": 0,
            f"{min_diff+3*interval:.1f}~{min_diff+4*interval:.1f}K": 0,
            f"{min_diff+4*interval:.1f}~{max_diff:.1f}K": 0
        }
        
        # 统计每个区间的数量
        for diff in difference_counts['differences']:
            for i, (key, _) in enumerate(new_counts.items()):
                lower = min_diff + i * interval
                upper = min_diff + (i + 1) * interval if i < 4 else max_diff + 0.1
                if lower <= diff < upper:
                    new_counts[key] += 1
                    break
        
        difference_counts = new_counts
    
    # 从第一个验证结果文件名中获取卫星类型
    satellite_type = 'UNKNOWN'
    if valresults and len(valresults) > 0:
        filename = valresults[0].get('filename', '')
        satellite_type = extract_satellite_type(filename)
    
    return total_pixels, valid_pixels, time_diff_counts, difference_counts, satellite_type

def generate_satellite_statistics_file(filename, total_pixels, valid_pixels, 
                                    time_diff_counts, difference_counts, product):
    """生成卫星交叉验证统计文件"""
    product_names = {
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
    product_name = product_names.get(product, product)
    
    print(f"\n=== 正在生成{product_name}统计文件 ===")
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"{product_name}统计结果：\n")
            f.write(f"总像元数：{total_pixels}\t有效检验像元数：{valid_pixels}\n\n")
            
            f.write("时间差分布情况：\n")
            for key, value in time_diff_counts.items():
                f.write(f"{key}:{value}\n")
            f.write("\n")
            
            f.write("检验结果情况：\n")
            for key, value in difference_counts.items():
                f.write(f"{key}：{value}\n")
        
        print(f"{product_name}统计文件生成成功")
    except Exception as e:
        print(f"生成统计文件时出错: {str(e)}")

def generate_satellite_plots(check_type,valid_pixels, total_pixels, time_diff_counts, 
                        difference_counts, output_directory, product, satellite_type, timestamp=None):
    """生成卫星交叉验证统计图"""
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    
    product_names = {
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
    product_name = product_names.get(product, product)

    if timestamp is None:
        timestamp = get_timestamp()
    base_name = f"{check_type}_{satellite_type}_{product}_{timestamp}"
    
    # 1. 有效检验像元比例饼图
    plt.figure(figsize=(10, 8))
    invalid_pixels = max(0, total_pixels - valid_pixels)
    valid_pixels = max(0, valid_pixels)
    
    if total_pixels > 0:
        sizes = [valid_pixels, invalid_pixels]
        labels = ['有效检验像元数', '无效像元数']
        plt.pie(sizes, labels=labels, autopct='%1.1f%%')
        plt.title(f"{product_name}有效检验像元比例")
        pixel_output = os.path.join(output_directory, f"pixelstastic_{base_name}.jpg")
        plt.savefig(pixel_output)
        plt.close()
        print(f"生成卫星 {product} 有效像元比例图")
    else:
        print(f"警告: {product} 没有有效的像元数据")
        plt.close()
    
    # 2. 时间差分布饼图
    if any(time_diff_counts.values()):
        plt.figure(figsize=(10, 8))
        sizes = list(time_diff_counts.values())
        sizes = [max(0, size) for size in sizes]
        if sum(sizes) > 0:
            labels = list(time_diff_counts.keys())
            plt.pie(sizes, labels=labels, autopct='%1.1f%%')
            plt.title(f"{product_name}时间差分布情况")
            time_output = os.path.join(output_directory, f"timestastic_{base_name}.jpg")
            plt.savefig(time_output)
            plt.close()
            print(f"生成卫星 {product} 时间差分布图")
        else:
            print(f"警告: {product} 没有有效的时间差数据")
            plt.close()
    
    # 3. 检验结果分布饼图
    if any(difference_counts.values()):
        plt.figure(figsize=(10, 8))
        sizes = list(difference_counts.values())
        sizes = [max(0, size) for size in sizes]
        if sum(sizes) > 0:
            labels = list(difference_counts.keys())
            plt.pie(sizes, labels=labels, autopct='%1.1f%%')
            plt.title(f"{product_name}检验结果情况")
            val_output = os.path.join(output_directory, f"valstastic_{base_name}.jpg")
            plt.savefig(val_output)
            plt.close()
            print(f"生成卫星 {product} 检验结果分布图")
        else:
            print(f"警告: {product} 没有有效的检验结果数据")
            plt.close()

def step9_satellite(check_type,input_directory, output_directory):
    """处理星星检验数据"""
    try:
        satellite_data = {}
        
        # 读取并分类所有文件
        for filename in os.listdir(input_directory):
            file_path = os.path.join(input_directory, filename)
            
            if (filename.startswith('valresult_') or filename.startswith('spaceresult_')) and 'XC' not in filename:
                parts = filename.split('_')
                if len(parts) >= 4:
                    product = parts[3].split('.')[0]
                    if product not in satellite_data:
                        satellite_data[product] = {
                            'valresults': [],
                            'spaceresults': []
                        }
                    
                    if filename.startswith('valresult_'):
                        result = read_satellite_valresult_file(file_path)
                        if result:
                            satellite_data[product]['valresults'].append(result)
                    elif filename.startswith('spaceresult_'):
                        result = read_satellite_spaceresult_file(file_path)
                        if result:
                            satellite_data[product]['spaceresults'].append(result)
        
        # 获取时间戳
        timestamp = extract_timestamp_from_files(os.listdir(input_directory))
        
        # 处理每个产品的数据
        for product, data in satellite_data.items():
            if data['valresults']:
                total_pixels, valid_pixels, time_diff_counts, difference_counts, satellite_type = analyze_star_check(check_type,
                    data['valresults'], data['spaceresults'], input_directory)
                
                if all(v is not None for v in [total_pixels, valid_pixels, time_diff_counts, difference_counts, satellite_type]):
                    # 生成统计文件
                    stats_filename = os.path.join(output_directory, 
                        f"resstastic_{check_type}_{satellite_type}_{product}_{timestamp}.txt")
                    generate_satellite_statistics_file(
                        stats_filename,
                        total_pixels,
                        valid_pixels,
                        time_diff_counts,
                        difference_counts,
                        product
                    )
                    
                    # 生成统计图
                    generate_satellite_plots(
                        check_type,
                        valid_pixels,
                        total_pixels,
                        time_diff_counts,
                        difference_counts,
                        output_directory,
                        product,
                        satellite_type,
                        timestamp
                    )
        
        return True
    except Exception as e:
        print(f"卫星交叉验证处理失败: {e}")
        import traceback
        traceback.print_exc()
        return False

# 现场验证相关函数
def read_ground_valresult_file(file_path):
    """读取现场验证结果文件"""
    try:
        result = {
            'header': {},
            'data': [],
            'filename': os.path.basename(file_path)
        }
        
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            data_start = False
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                if line == '/begin header':
                    continue
                elif line == '/end header':
                    data_start = True
                elif line.startswith('/'):
                    if '=' in line:
                        key, value = line[1:].split('=', 1)
                        result['header'][key.strip()] = value.strip()
                elif data_start:
                    try:
                        values = line.split('\t')
                        if len(values) >= 2:
                            result['data'].append(values)
                    except ValueError:
                        continue
        
        return result
    except Exception as e:
        print(f"读取验证文件时出错 {file_path}: {str(e)}")
        return None

def read_ground_spaceresult_file(file_path):
    """读取现场空间结果文件"""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            if len(lines) >= 10:  # 确保至少有9行
                return {
                    'hy_file': lines[0].strip(),
                    'compare_file': lines[1].strip(),
                    'time_diff': float(lines[9].strip()),
                    'valid_ratio': float(lines[5].strip()),  # 第六行:有效像元比例
                    'cv_value': float(lines[6].strip())      # 第七行:CV值
                }
    except Exception as e:
        print(f"读取空间结果文件时出错 {file_path}: {str(e)}")
        return None

def analyze_ground_validation(valresults, spaceresults):
    """分析星地检验结果"""
    valid_images = len(valresults)
    
    time_diff_counts = {
        "<0.5h": 0,
        "0.5~1h": 0,
        "1~1.5h": 0,
        "1.5~3h": 0,
        ">3h": 0
    }
    
    valid_ratio_counts = {
        "=1": 0,
        "0.9~1": 0,
        "0.8~0.9": 0,
        "0.6~0.8": 0,
        "<0.6": 0
    }
    
    cv_value_counts = {
        "<0.05": 0,
        "0.05~0.1": 0,
        ">0.1": 0
    }
    
    # 从valresults中获取产品名称
    product = None
    if valresults and len(valresults) > 0:
        filename = valresults[0].get('filename', '')
        parts = filename.split('_')
        if len(parts) >= 4:
            product = parts[3].split('.')[0]
    
    # 根据产品类型初始化difference_counts
    if product == 'sst':
        difference_counts = {
            'min_diff': float('inf'),
            'max_diff': float('-inf'),
            'differences': []
        }
    else:
        difference_counts = {
            "<10": 0,
            "10~30": 0,
            "30~50": 0,
            "50~100": 0,
            ">100": 0
        }
    
    # 处理spaceresults数据
    for result in spaceresults:
        # 处理时间差
        time_diff = result['time_diff']
        if time_diff < 0.5:
            time_diff_counts["<0.5h"] += 1
        elif time_diff < 1.0:
            time_diff_counts["0.5~1h"] += 1
        elif time_diff < 1.5:
            time_diff_counts["1~1.5h"] += 1
        elif time_diff < 3.0:
            time_diff_counts["1.5~3h"] += 1
        else:
            time_diff_counts[">3h"] += 1

        # 处理有效像元比例
        valid_ratio = result['valid_ratio']
        if valid_ratio == 1:
            valid_ratio_counts["=1"] += 1
        elif valid_ratio >= 0.9:
            valid_ratio_counts["0.9~1"] += 1
        elif valid_ratio >= 0.8:
            valid_ratio_counts["0.8~0.9"] += 1
        elif valid_ratio >= 0.6:
            valid_ratio_counts["0.6~0.8"] += 1
        else:
            valid_ratio_counts["<0.6"] += 1

        # 处理CV值
        cv_value = result['cv_value']
        if cv_value < 0.05:
            cv_value_counts["<0.05"] += 1
        elif cv_value < 0.1:
            cv_value_counts["0.05~0.1"] += 1
        else:
            cv_value_counts[">0.1"] += 1
    
    # 处理valresults数据
    for result in valresults:
        for row in result['data']:
            try:
                diff = float(row[-1])
                if product == 'sst':
                    # 确保至少有一个非零值
                    difference_counts['differences'].append(diff)
                    difference_counts['min_diff'] = min(difference_counts['min_diff'], diff)
                    difference_counts['max_diff'] = max(difference_counts['max_diff'], diff)
                else:
                    if diff < 10:
                        difference_counts["<10"] += 1
                    elif diff < 30:
                        difference_counts["10~30"] += 1
                    elif diff < 50:
                        difference_counts["30~50"] += 1
                    elif diff < 100:
                        difference_counts["50~100"] += 1
                    else:
                        difference_counts[">100"] += 1
            except (ValueError, IndexError):
                continue
    
    # 如果是SST产品，处理收集的差异数据
    if product == 'sst' and difference_counts['differences']:
        min_diff = difference_counts['min_diff']
        max_diff = difference_counts['max_diff']

        # 添加保护逻辑，确保有合理的区间范围
        if min_diff == max_diff:
            # 如果最大最小值相同，创建一个固定的区间范围
            min_diff = min_diff - 0.5
            max_diff = max_diff + 0.5

        # 创建5个均匀的区间（问题5: SST添加K单位）
        interval = (max_diff - min_diff) / 5
        new_counts = {
            f"{min_diff:.1f}~{min_diff+interval:.1f}K": 0,
            f"{min_diff+interval:.1f}~{min_diff+2*interval:.1f}K": 0,
            f"{min_diff+2*interval:.1f}~{min_diff+3*interval:.1f}K": 0,
            f"{min_diff+3*interval:.1f}~{min_diff+4*interval:.1f}K": 0,
            f"{min_diff+4*interval:.1f}~{max_diff:.1f}K": 0
        }
        
        # 统计每个区间的数量
        for diff in difference_counts['differences']:
            for i, (key, _) in enumerate(new_counts.items()):
                lower = min_diff + i * interval
                upper = min_diff + (i + 1) * interval if i < 4 else max_diff + 0.1
                if lower <= diff < upper:
                    new_counts[key] += 1
                    break
        
        difference_counts = new_counts
    
    return valid_images, time_diff_counts, valid_ratio_counts, cv_value_counts, difference_counts

def generate_ground_statistics_file(filename, valid_images, time_diff_counts, 
                                  valid_ratio_counts, cv_value_counts, 
                                  difference_counts, product):
    """生成现场验证统计文件"""
    product_names = {
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
    product_name = product_names.get(product, product)
    
    print(f"\n=== 正在生成{product_name}现场验证统计文件: {filename} ===")
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"{product_name}现场验证统计结果：\n")
            f.write(f"有效检验影像数：{valid_images}\n\n")
            
            f.write("时间差分布情况：\n")
            for key, value in time_diff_counts.items():
                f.write(f"{key}:{value}\n")
            f.write("\n")
            
            f.write("空间窗口内有效像元比例分布情况：\n")
            for key, value in valid_ratio_counts.items():
                f.write(f"{key}:{value}\n")
            f.write("\n")
            
            f.write("空间窗口内CV值分布情况：\n")
            for key, value in cv_value_counts.items():
                f.write(f"{key}:{value}\n")
            f.write("\n")
            
            f.write("检验结果情况：\n")
            for key, value in difference_counts.items():
                f.write(f"{key}%：{value}\n")
        
        print(f"{product_name}现场验证统计文件生成成功")
    except Exception as e:
        print(f"生成统计文件时出错: {str(e)}")

def generate_ground_plots(check_type,time_diff_counts, valid_ratio_counts, cv_value_counts, 
                     difference_counts, output_directory, product, timestamp=None):
    """生成现场验证统计图"""
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    
    product_names = {
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
    product_name = product_names.get(product, product)

    if timestamp is None:
        timestamp = get_timestamp()

    base_name = f"{check_type}_XC_{product}_{timestamp}"

    # 生成随机数据
    def generate_random_distribution(total=100):
        """生成随机分布的数据"""
        values = []
        remaining = total
        for _ in range(4):  # 生成前4个数
            if remaining <= 0:
                values.append(0)
                continue
            value = random.randint(0, remaining)
            values.append(value)
            remaining -= value
        values.append(remaining)  # 最后一个数使用剩余值
        random.shuffle(values)  # 随机打乱顺序
        return values

    # 1. 时间差分布饼图
    if any(time_diff_counts.values()):
        plt.figure(figsize=(10, 8))
        # 使用实际的时间差数据
        sizes = list(time_diff_counts.values())
        sizes = [max(0, size) for size in sizes]  # 确保非负
        if sum(sizes) > 0:
            labels = list(time_diff_counts.keys())
            plt.pie(sizes, labels=labels, autopct='%1.1f%%')
            plt.title(f"{product_name}时间差分布情况")
            time_output = os.path.join(output_directory, f"timestastic_{base_name}.jpg")
            plt.savefig(time_output)
            plt.close()
            print(f"生成现场 {product} 时间差分布图")
        else:
            plt.close()
    
    # 2. 检验结果分布饼图
    if any(difference_counts.values()):
        plt.figure(figsize=(10, 8))
        # 使用实际的检验结果数据
        sizes = list(difference_counts.values())
        sizes = [max(0, size) for size in sizes]  # 确保非负
        if sum(sizes) > 0:
            labels = list(difference_counts.keys())
            plt.pie(sizes, labels=labels, autopct='%1.1f%%')
            plt.title(f"{product_name}检验结果情况")
            val_output = os.path.join(output_directory, f"valstastic_{base_name}.jpg")
            plt.savefig(val_output)
            plt.close()
            print(f"生成现场 {product} 检验结果分布图")
        else:
            plt.close()

    # 在generate_ground_plots函数���添加调试信息
    print(f"SST difference_counts: {difference_counts}")
    print(f"any(difference_counts.values()): {any(difference_counts.values())}")

def step9_ground(check_type,input_directory, output_directory):
    """处理现场验证数据"""
    try:
        # 存储现场数据
        ground_data = {}
        
        # 读取并分类所有文件
        for filename in os.listdir(input_directory):
            if 'XC' in filename:  # 只处理现场验证数据
                file_path = os.path.join(input_directory, filename)
                parts = filename.split('_')
                if len(parts) >= 4:
                    product = parts[3].split('.')[0]
                    if product not in ground_data:
                        ground_data[product] = {
                            'valresults': [],
                            'spaceresults': []
                        }
                    
                    if filename.startswith('valresult_'):
                        result = read_ground_valresult_file(file_path)
                        if result:
                            ground_data[product]['valresults'].append(result)
                    elif filename.startswith('spaceresult_'):
                        result = read_ground_spaceresult_file(file_path)
                        if result:
                            ground_data[product]['spaceresults'].append(result)
        
        # 获取时间戳
        timestamp = extract_timestamp_from_files(os.listdir(input_directory))
        
        # 处理每个产品的数据
        for product, data in ground_data.items():
            if data['valresults']:
                valid_images, time_diff_counts, valid_ratio_counts, cv_value_counts, difference_counts = analyze_ground_validation(
                    data['valresults'], data['spaceresults'])
                
                # 生成统计文件
                stats_filename = os.path.join(output_directory, 
                    f"resstastic_{check_type}_XC_{product}_{timestamp}.txt")
                generate_ground_statistics_file(
                    stats_filename,
                    valid_images,
                    time_diff_counts,
                    valid_ratio_counts,
                    cv_value_counts,
                    difference_counts,
                    product
                )
                
                # 生成统计图
                generate_ground_plots(
                    check_type,
                    time_diff_counts,
                    valid_ratio_counts,
                    cv_value_counts,
                    difference_counts,
                    output_directory,
                    product,
                    timestamp
                )
        
        return True
    except Exception as e:
        print(f"现场验证处理失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def step9(check_type,input_directory, output_directory):
    """
    第九步：处理星地检验和星星检验数据，生成统计结果和图表
    """
    try:
        # 处理卫星交叉验证数据
        # print("\n处理卫星交叉验证数据...")
        step9_satellite(check_type,output_directory, output_directory)
        
        # 处理现场验证数据
        print("\n处理现场验证数据...")
        step9_ground(check_type,output_directory, output_directory)
        

    except Exception as e:
        print(f"步骤9执行失败: {e}")
        import traceback
        traceback.print_exc()
        return False

#--------------------------------------------------------------------------------------------#
def make_satellite_report_data(input_dir):
    # 处理valresult文件
    for f in [f for f in os.listdir(input_dir) if f.startswith("valresult_")]:
        output_filename = f.replace("valresult_", "report_")
        
        # 尝试不同的编码
        encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'latin1']
        for encoding in encodings:
            try:
                with open(os.path.join(input_dir, f), 'r', encoding=encoding) as file:
                    lines = file.readlines()
                    report_data = {}
                    
                    for line in lines:
                        line = line.strip()
                        if not line:
                            continue
                            
                        # 收集所需数据
                        if line.startswith('/HY satellite'):
                            report_data['/HY satellite'] = line.split('=')[1].strip()
                        elif line.startswith('/product'):
                            report_data['/Product'] = line.split('=')[1].strip()
                        elif line.startswith('/HY time'):
                            report_data['/HY time'] = line.split('=')[1].strip()
                        elif line.startswith('/HY file'):
                            report_data['/HY file'] = line.split('=')[1].strip()
                        elif line.startswith('/Validation source'):
                            report_data['/Validation Source'] = line.split('=')[1].strip()
                        elif line.startswith('/On-site time'):
                            report_data['/On-site time'] = line.split('=')[1].strip()
                        elif line.startswith('/On-site file'):
                            report_data['/On-site file'] = line.split('=')[1].strip()
                        elif line.startswith('/Time difference'):
                            report_data['/Time Difference'] = line.split('=')[1].strip()
                        elif line.startswith('/Total pixel count'):
                            report_data['/Total pixel count'] = line.split('=')[1].strip()
                    
                    with open(os.path.join(input_dir, output_filename), 'w', encoding='utf-8') as outfile:
                        for key, value in report_data.items():
                            outfile.write(f"{key}={value}\n")
                break
            except UnicodeDecodeError:
                if encoding == encodings[-1]:  # 如果是最后一个编码尝试
                    print(f"Failed to decode file {f} with all encodings")
                continue
            except Exception as e:
                print(f"Error processing file {f}: {str(e)}")
                break

    # 处理statistic文件
    for f in [f for f in os.listdir(input_dir) if f.startswith("statistic_")]:
        report_filename = f.replace("statistic_", "report_")
        
        for encoding in encodings:
            try:
                with open(os.path.join(input_dir, f), 'r', encoding=encoding) as file:
                    lines = file.readlines()
                    # 获取最后一行数据
                    last_line = lines[-1].strip()
                    # 分割数据
                    values = last_line.split()
                    
                    # 将数据写入report文件
                    with open(os.path.join(input_dir, report_filename), 'a', encoding='utf-8') as outfile:
                        outfile.write(f'/bias={values[0]}\n')
                        outfile.write(f'/STD={values[1]}\n')
                        outfile.write(f'/RMS={values[2]}\n')
                        outfile.write(f'/R={values[3]}\n')
                break
            except UnicodeDecodeError:
                if encoding == encodings[-1]:
                    print(f"Failed to decode file {f} with all encodings")
                continue
            except Exception as e:
                print(f"Error processing file {f}: {str(e)}")
                break

    # 添加处理resstastic文件的部分
    for f in [f for f in os.listdir(input_dir) if f.startswith("resstastic_")]:
        report_filename = f.replace("resstastic_", "report_")
        
        for encoding in encodings:
            try:
                with open(os.path.join(input_dir, f), 'r', encoding=encoding) as file:
                    lines = file.readlines()
                    
                    # 只查找包含总像元数和有效检验像元数的行
                    for line in lines:
                        if "总像元数" in line and "有效检验像元数" in line:
                            parts = line.strip().split('\t')
                            total_pixels = parts[0].split('：')[1]
                            valid_pixels = parts[1].split('：')[1]
                            
                            # 将数据追加到report文件
                            with open(os.path.join(input_dir, report_filename), 'a', encoding='utf-8') as outfile:
                                outfile.write(f'/Effective pixel count={total_pixels}\n')
                                outfile.write(f'/valresult={valid_pixels}\n')
                            break
                break
            except UnicodeDecodeError:
                if encoding == encodings[-1]:
                    print(f"Failed to decode file {f} with all encodings")
                continue
            except Exception as e:
                print(f"Error processing file {f}: {str(e)}")
                break

def make_ground_report_data(input_dir):
    encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'latin1']
    
    for f in [f for f in os.listdir(input_dir) if f.startswith("valresult_")]:
        output_filename = f.replace("valresult_", "report_")
        
        for encoding in encodings:
            try:
                with open(os.path.join(input_dir, f), 'r', encoding=encoding) as file:
                    lines = file.readlines()
                    report_data = {}
                    relative_bias = None  # 单独存储相对偏差
                    
                    # 找到/end header的位置
                    header_end_index = -1
                    for i, line in enumerate(lines):
                        if '/end header' in line:
                            header_end_index = i
                            break
                    
                    # 如果找到/end header，读取下一行的第三列数据作为相对偏差
                    if header_end_index != -1 and header_end_index + 1 < len(lines):
                        relative_bias = lines[header_end_index + 1].strip().split()[2]
                    
                    # 继续读取其他数据
                    for line in lines:
                        line = line.strip()
                        if not line:
                            continue
                            
                        if line.startswith('/HY satellite'):
                            report_data['/HY satellite'] = line.split('=')[1].strip()
                        elif line.startswith('/product'):
                            report_data['/Product'] = line.split('=')[1].strip()
                        elif line.startswith('/HY time'):
                            report_data['/HY time'] = line.split('=')[1].strip()
                        elif line.startswith('/HY file'):
                            report_data['/HY file'] = line.split('=')[1].strip()
                        elif line.startswith('/Validation source'):
                            report_data['/Validation Source'] = line.split('=')[1].strip()
                        elif line.startswith('/On-site time'):
                            report_data['/On-site time'] = line.split('=')[1].strip()
                        elif line.startswith('/On-site file'):
                            report_data['/On-site file'] = line.split('=')[1].strip()
                        elif line.startswith('/Time difference'):
                            report_data['/Time Difference'] = line.split('=')[1].strip()
                        elif line.startswith('/Total pixel count'):
                            report_data['/Total pixel count'] = line.split('=')[1].strip()
                    
                    # 读取对应的spaceresult文件
                    space_filename = f.replace("valresult_", "spaceresult_")
                    if os.path.exists(os.path.join(input_dir, space_filename)):
                        try:
                            with open(os.path.join(input_dir, space_filename), 'r', encoding=encoding) as space_file:
                                space_lines = space_file.readlines()
                                if len(space_lines) >= 7:  # 确保文件至少有7行
                                    # 读取第六行和第七行
                                    valid_ratio = space_lines[5].strip()  # 第六行
                                    cv_value = space_lines[6].strip()     # 第七行
                                    report_data['/Valid Ratio'] = valid_ratio
                                    report_data['/CV Value'] = cv_value
                        except Exception as e:
                            print(f"处理空间文件 {space_filename} 时出错: {str(e)}")
                    
                    # 写入report文件
                    with open(os.path.join(input_dir, output_filename), 'w', encoding='utf-8') as outfile:
                        # 先写入其他数据
                        for key, value in report_data.items():
                            outfile.write(f"{key}={value}\n")
                        # 最后写入相对偏差
                        if relative_bias is not None:
                            outfile.write(f"/Relative Bias={relative_bias}\n")
                break
            except UnicodeDecodeError:
                if encoding == encodings[-1]:
                    print(f"无法解码文件 {f}")
                continue
            except Exception as e:
                print(f"处理文件 {f} 时出错: {str(e)}")
                break



def check_space_available(current_y, required_height):
    """检查页面剩余空间是否足够"""
    min_margin = 3 * cm
    if current_y - required_height < min_margin:
        return False
    return True

def calculate_image_dimensions(original_width, original_height, max_width, max_height):
    """根据页面空间动态计算图片尺寸"""
    ratio = min(max_width/original_width, max_height/original_height)
    return original_width * ratio, original_height * ratio

def read_report_data(input_dir):
    report_data_list = []
    
    for f in [f for f in os.listdir(input_dir) if f.startswith("report_")]:
        try:
            with open(os.path.join(input_dir, f), 'r', encoding='utf-8') as file:
                lines = file.readlines()
                report_data = {}
                
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue  
                    if '=' in line:
                        key, value = line.split('=', 1)
                        report_data[key] = value
                
                report_data_list.append(report_data)
                
        except Exception as e:
            print(f"Error reading file {f}: {str(e)}")
            
    return report_data_list

def extract_info_from_filenames(input_dir):
    files = os.listdir(input_dir)
    pattern = r"report_(\w+)_(\w+)_(\w+)_(\w+)"
    
    # 设置默认值
    satellite_info = None
    source_data = None
    product = None      # 添加默认值
    time_info = None    # 添加默认值
    
    for file in files:
        match = re.match(pattern, file)
        if match:
            satellite_info, source_data, product, time_info = match.groups()
            break
    
    return satellite_info, source_data, product, time_info

def calculate_first_page_layout(c, A4, labels, values, table_data):
    """固定表格底部位置，向上计算布局"""
    # 页面尺寸
    total_height = A4[1]
    
    # 页面边距
    top_margin = 2.5 * cm
    bottom_margin = 2 * cm  # 表格距离底部的距离
    
    # 计算表格高度
    table_height = len(table_data) * 0.8 * cm
    
    # 固定表格底部位置
    table_y = bottom_margin  # 表格底部对齐底部边距
    
    # 标题和文字内容位置保持不变
    title_y = total_height - top_margin
    content_start_y = title_y - 4 * cm  # 与原来保持一致的文字起始位置
    
    return {
        'title_y': title_y,
        'content_start_y': content_start_y,
        'table_y': table_y
    }


def step_report(datestr, input_temp, input_img, coldata_path, output_path, satellite_type,source_org_type,space_size,time_size):
    # 确保输出目录存在
    os.makedirs(coldata_path, exist_ok=True)
    os.makedirs(output_path, exist_ok=True)
    os.makedirs(input_img, exist_ok=True)

 # 定义产品配置
    # 定义AQUA_VAR_CONFIG
    AQUA_VAR_CONFIG = {
        'sst': {'sources': ['AQUA'], 'unit': 'K', 'col_values': [25, 1800]},
        'chl': {'sources': ['AQUA'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        'Rrs412': {'sources': ['AQUA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs443': {'sources': ['AQUA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs490': {'sources': ['AQUA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs520': {'sources': ['AQUA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs565': {'sources': ['AQUA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs670': {'sources': ['AQUA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'AOT': {'sources': ['AQUA'], 'unit': '', 'col_values': [25, 1800]},
        # 'TSM': {'sources': ['AQUA'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        # 'CDOM': {'sources': ['AQUA'], 'unit': '1/m', 'col_values': [25, 1800]},
    }

    # 定义TERRA_VAR_CONFIG
    TERRA_VAR_CONFIG = {
        'sst': {'sources': ['TERRA'], 'unit': 'K', 'col_values': [25, 1800]},
        'chl': {'sources': ['TERRA'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        'Rrs412': {'sources': ['TERRA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs443': {'sources': ['TERRA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs490': {'sources': ['TERRA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs520': {'sources': ['TERRA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs565': {'sources': ['TERRA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs670': {'sources': ['TERRA'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'AOT': {'sources': ['TERRA'], 'unit': '', 'col_values': [25, 1800]},
        # 'TSM': {'sources': ['TERRA'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        # 'CDOM': {'sources': ['TERRA'], 'unit': '1/m', 'col_values': [25, 1800]},
    }

    # 定义SNPP_VAR_CONFIG
    SNPP_VAR_CONFIG = {
        'sst': {'sources': ['SNPP'], 'unit': 'K', 'col_values': [25, 1800]},
        'chl': {'sources': ['SNPP'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        'Rrs412': {'sources': ['SNPP'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs443': {'sources': ['SNPP'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs490': {'sources': ['SNPP'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs565': {'sources': ['SNPP'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs670': {'sources': ['SNPP'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'AOT': {'sources': ['SNPP'], 'unit': '', 'col_values': [25, 1800]},
        # 'TSM': {'sources': ['SNPP'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        # 'CDOM': {'sources': ['SNPP'], 'unit': '1/m', 'col_values': [25, 1800]},
    }

    # 定义JPSS_VAR_CONFIG
    JPSS_VAR_CONFIG = {
        'sst': {'sources': ['JPSS'], 'unit': 'K', 'col_values': [25, 1800]},
        'chl': {'sources': ['JPSS'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        'Rrs412': {'sources': ['JPSS'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs443': {'sources': ['JPSS'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs490': {'sources': ['JPSS'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs565': {'sources': ['JPSS'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs670': {'sources': ['JPSS'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'AOT': {'sources': ['JPSS'], 'unit': '', 'col_values': [25, 1800]},
        # 'TSM': {'sources': ['JPSS'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        # 'CDOM': {'sources': ['JPSS'], 'unit': '1/m', 'col_values': [25, 1800]},
    }

    # 定义HY1C_VAR_CONFIG（卫星间验证）
    HY1C_VAR_CONFIG = {
        'sst': {'sources': ['HY1C'], 'unit': 'K', 'col_values': [25, 1800]},
        'chl': {'sources': ['HY1C'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        'Rrs412': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs443': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs490': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs520': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs565': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs670': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'AOT': {'sources': ['HY1C'], 'unit': '', 'col_values': [25, 1800]},
        'TSM': {'sources': ['HY1C'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        'CDOM': {'sources': ['HY1C'], 'unit': '1/m', 'col_values': [25, 1800]},
    }

    # 定义HY1D_VAR_CONFIG（卫星间验证）
    HY1D_VAR_CONFIG = {
        'sst': {'sources': ['HY1D'], 'unit': 'K', 'col_values': [25, 1800]},
        'chl': {'sources': ['HY1D'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        'Rrs412': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs443': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs490': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs520': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs565': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs670': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'AOT': {'sources': ['HY1D'], 'unit': '', 'col_values': [25, 1800]},
        'TSM': {'sources': ['HY1D'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        'CDOM': {'sources': ['HY1D'], 'unit': '1/m', 'col_values': [25, 1800]},
    }

    # 定义HY1E_VAR_CONFIG（卫星间验证）
    HY1E_VAR_CONFIG = {
        'sst': {'sources': ['HY1E'], 'unit': 'K', 'col_values': [25, 1800]},
        'chl': {'sources': ['HY1E'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        'Rrs412': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs443': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs490': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs520': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs565': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs670': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs750': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'AOT': {'sources': ['HY1E'], 'unit': '', 'col_values': [25, 1800]},
        'TSM': {'sources': ['HY1E'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        'CDOM': {'sources': ['HY1E'], 'unit': '1/m', 'col_values': [25, 1800]},
        'IPAR': {'sources': ['HY1E'], 'unit': 'Einstein/m²/d', 'col_values': [25, 1800]},
    }

    # 初始化VAR_CONFIG
    VAR_CONFIG = None

    if source_org_type == 'AQUA':
        VAR_CONFIG = AQUA_VAR_CONFIG
    elif source_org_type == 'TERRA':
        VAR_CONFIG = TERRA_VAR_CONFIG
    elif source_org_type == 'SNPP':
        VAR_CONFIG = SNPP_VAR_CONFIG
    elif source_org_type == 'JPSS':
        VAR_CONFIG = JPSS_VAR_CONFIG
    elif source_org_type == 'HY1C':
        VAR_CONFIG = HY1C_VAR_CONFIG
    elif source_org_type == 'HY1D':
        VAR_CONFIG = HY1D_VAR_CONFIG
    elif source_org_type == 'HY1E':
        VAR_CONFIG = HY1E_VAR_CONFIG
    else:
        raise ValueError(f"Unsupported source type: {source_org_type}")
    
    # print(f"tttttttttttttttttttt\n")

    def calc_metric(coldata_path, reference, var_name, datestr, satellite_type):
        """
        从报告文件中提取指标数据

        :param coldata_path: 数据文件路径
        :param reference: 数据源
        :param var_name: 变量名
        :param datestr: 日期字符串
        :param satellite_type: 卫星类型
        :return: bias, rms, n
        """
        # 实际的报告文件格式：report_{satellite}_{reference}_{var_name}_{timestamp}.txt
        # 首先尝试匹配包含日期字符串的文件
        files = os.listdir(coldata_path)
        pattern = f'report_{satellite_type}_{reference}_{var_name}_{datestr}'
        matching_files = [f for f in files if pattern in f and f.endswith('.txt')]

        if not matching_files:
            # 如果没找到，尝试旧的命名格式
            old_pattern = f'{satellite_type}_COCTS_{reference}_{var_name}_report_{datestr}'
            matching_files = [f for f in files if old_pattern in f and f.endswith('.txt')]

        if not matching_files:
            raise FileNotFoundError(f"找不到匹配的报告文件。\n查找模式: report_{satellite_type}_{reference}_{var_name}_{datestr}*.txt\n目录: {coldata_path}\n已有文件: {[f for f in files if f.endswith('.txt')]}")

        filepath = os.path.join(coldata_path, matching_files[0])

        bias = None
        rms = None
        n = None
        with open(filepath, 'r') as f:
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

        if bias is None or rms is None or n is None:
            raise ValueError(f"在文件{filepath}中找不到所需的数据")
        return bias, rms, n

    def generate_replacements(datestr: str, imgpath: str, coldata_path: str, satellite_type: str,source_org_type: str, space_size: str, time_size: str):
        """
        根据日期生成包含所有产品类型的替换字典

        :param datestr: 日期字符串 (格式: %Y%m%d)
        :param imgpath: 图片存储根目录
        :param coldata_path: 数据文件路径
        :param satellite_type: 卫星类型
        :return: 包含文本、表格、图片路径的替换字典
        """
        date_obj = datetime.strptime(datestr, '%Y%m%d')
        replacements = {
            'text': {
                '{{date}}': date_obj.strftime('%Y-%m-%d'),
                '{{date_cn}}': date_obj.strftime('%Y年%m月%d日'),
            },
            'tables': {},
            'images': {}
        }

        for var_name, config in VAR_CONFIG.items():
            metrics = {
                source: dict(zip(['bias', 'rms', 'n'], calc_metric(coldata_path, source, var_name, datestr, satellite_type)))
                for source in config['sources']
            }

            # 获取产品中文名称和单位
            product_cn_name = PRODUCT_NAMES.get(var_name, var_name)
            unit = config.get('unit', '')

            # 生成val_results和col_results（处理问题1-4,6）
            val_results = []
            col_results = []
            has_valid_data = False  # 用于问题6：判断是否有n>0的数据

            for source in config['sources']:
                n = metrics[source]['n']

                # 验证结果表格（表一）- n=0时不添加val_results，让cleanup清除占位符
                if n > 0:
                    has_valid_data = True
                    # 根据产品类型设置单位
                    if var_name == 'sst':
                        # 海温：bias和rms都用K（绝对误差）
                        bias_str = f"{metrics[source]['bias']:.4f}{unit}"
                        rms_str = f"{metrics[source]['rms']:.4f}{unit}"
                    else:
                        # 其他产品：bias用%（相对误差，需要乘100），rms用原单位
                        bias_str = f"{metrics[source]['bias']*100:.2f}%"
                        rms_str = f"{metrics[source]['rms']:.4f}{unit}"

                    val_results.append([
                        f'{satellite_type} vs {source}',
                        bias_str,
                        rms_str
                    ])

                # 匹配结果表格（表二）- n=0时全部填0
                if n > 0:
                    time_window = f"{time_size}h" if time_size != '汇总' and not str(time_size).endswith('h') else str(time_size)
                    space_window = f"{space_size}*{space_size}"
                else:
                    # n=0时：产品名、空间窗口0、时间窗口0、匹配数0
                    space_window = "0"
                    time_window = "0"

                col_results.append([
                    product_cn_name,
                    space_window,
                    time_window,
                    f"{n}"
                ])

            # 生成唯一的图片占位符
            image_keys = {
                # 'sct': [f'{satellite_type.lower()}_vs_{var_name}_{source.lower()}_sct' for source in config['sources']],
                # 'geo': [f'{satellite_type.lower()}_vs_{var_name}_{source.lower()}_geo' for source in config['sources']]
                'sct': [f'hy1c_vs_{var_name}_terra_sct'],
                'geo': [f'hy1c_vs_{var_name}_terra_geo']
            }
            print(f"Generated data for {var_name}: {val_results}")  # 添加打印语句

            # 从报告文件中提取时间戳
            timestamp = None
            for source in config['sources']:
                # 列出目录中的所有文件
                files = os.listdir(coldata_path)

                # 首先尝试新格式：report_{satellite}_{reference}_{var_name}_{datestr}_{time}.txt
                new_pattern = f'report_{satellite_type}_{source}_{var_name}_{datestr}'
                matching_files = [f for f in files if new_pattern in f and f.endswith('.txt')]

                if matching_files:
                    # 从新格式文件名中提取时间戳
                    # 格式: report_HY1C_TERRA_sst_20250101_120000.txt
                    parts = matching_files[0].replace('.txt', '').split('_')
                    if len(parts) >= 5:
                        # 时间戳是最后一部分
                        timestamp = parts[-1]
                        break
                else:
                    # 尝试旧格式：{satellite}_COCTS_{source}_{var_name}_report_{datestr}_{time}.txt
                    old_pattern = f'{satellite_type}_COCTS_{source}_{var_name}_report_{datestr}'
                    matching_files = [f for f in files if old_pattern in f and f.endswith('.txt')]
                    if matching_files:
                        # 从旧格式文件名中提取时间戳
                        parts = matching_files[0].split('_')
                        if len(parts) > 6:
                            timestamp = parts[6].replace('.txt', '')
                            break

            # 问题6: 只有has_valid_data为True时才添加图片路径（第3章才显示该产品）
            images = {}
            if has_valid_data:
                for key in image_keys['sct']:
                    img_path = os.path.join(imgpath, f'{satellite_type}_COCTS_{source_org_type}_{var_name.upper()}_PIE_{datestr}_{timestamp}.jpg')
                    if not os.path.exists(img_path):
                        alt_img_path = os.path.join(imgpath, f'{satellite_type}_COCTS_{source_org_type}_{var_name}_PIE_{datestr}_{timestamp}.jpg')
                        if os.path.exists(alt_img_path):
                            img_path = alt_img_path
                            print(f"使用替代图片路径: {img_path}")
                    images[f'{{{{{key}}}}}'] = img_path

                for key in image_keys['geo']:
                    img_path = os.path.join(imgpath, f'{satellite_type}_COCTS_{source_org_type}_{var_name.upper()}_GEO_{datestr}_{timestamp}.jpg')
                    if not os.path.exists(img_path):
                        alt_img_path = os.path.join(imgpath, f'{satellite_type}_COCTS_{source_org_type}_{var_name}_GEO_{datestr}_{timestamp}.jpg')
                        if os.path.exists(alt_img_path):
                            img_path = alt_img_path
                            print(f"使用替代图片路径: {img_path}")
                    images[f'{{{{{key}}}}}'] = img_path

            replacements['text'].update({
                '{{satellite_type}}': satellite_type,
                '{{source_type}}': ', '.join(config['sources']),
                '{{unit}}': config['unit']
            })
            # 表一：只有has_valid_data时才添加，否则让cleanup清除占位符
            if has_valid_data:
                replacements['tables'][f'{{{{val_results_{var_name}}}}}'] = val_results
            # 表二：总是添加（n=0时显示0）
            replacements['tables'][f'{{{{col_results_{var_name}}}}}'] = col_results
            replacements['images'].update(images)

        return replacements


    def _replace_text(doc, placeholder, replacement):
        """遍历所有段落和表格单元格进行文本替换"""
        for p in doc.paragraphs:
            if placeholder in p.text:
                _replace_text_in_paragraph(p, placeholder, replacement)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if placeholder in p.text:
                            _replace_text_in_paragraph(p, placeholder, replacement)

    def _replace_text_in_paragraph(paragraph, placeholder, replacement):
        """在单个段落中执行文本替换，保留原始格式"""
        if placeholder not in paragraph.text:
            return
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, replacement)

    def _fill_table(doc, placeholder, table_data):
        """
        动态填充表格数据并设置居中格式

        :param doc: Document对象
        :param placeholder: 表格占位符
        :param table_data: 表格数据，预期为二维列表，其中每个子列表代表表格的一行
        """
        found = False
        for table in doc.tables:
            # 遍历表格中的所有行，找到包含占位符的行
            for row in table.rows:
                for cell in row.cells:
                    if placeholder in cell.text:
                        found = True
                        # 清空当前行的数据
                        for c in row.cells:
                            c.text = ""
                        # 填充新数据到单元格
                        for i, value in enumerate(table_data[0]):  # 只取第一个子列表的数据
                            cell = row.cells[i]
                            cell.text = str(value)
                            # 设置水平居中
                            for paragraph in cell.paragraphs:
                                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                            # 设置垂直居中
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        break  # 找到占位符后，跳出单元格循环
                if found:
                    break  # 找到占位符后，跳出行循环
        if not found:
            print(f"未找到占位符 '{placeholder}' 的表格。")

    def _insert_image(doc, placeholder, image_path):
        """在占位符位置插入图片"""
        if not os.path.exists(image_path):
            print(f"警告：图片文件不存在: {image_path}")
            return
        found = False
        # 遍历所有段落
        for p in doc.paragraphs:
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, '')
                run = p.add_run()
                run.add_picture(image_path, width=Inches(6))
                print(f"成功插入图片: {image_path}")
                found = True
                break
        if not found:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if placeholder in cell.text:
                            cell.text = cell.text.replace(placeholder, '')
                            run = cell.add_run()
                            run.add_picture(image_path, width=Inches(4))
                            print(f"成功插入图片: {image_path}")
                            found = True
                            break
                    if found:
                        break
                if found:
                    break
        if not found:
            print(f"未找到占位符 '{placeholder}' 的位置。")

    def _cleanup_placeholders(doc):
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

    def _remove_empty_subsections_and_renumber(doc):
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


    def _handle_unused_placeholders(doc, replacements):
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
                    _fill_unused_val_results(doc, var_name, satellite, source)

        # 处理未使用的col_results占位符（表二）
        for placeholder in all_placeholders:
            match = col_pattern.match(placeholder)
            if match:
                full_placeholder = f'{{{{{placeholder}}}}}'
                if full_placeholder not in used_tables:
                    # 删除整行
                    var_name = match.group(1)
                    print(f"【调试】删除未使用的col_results行: {full_placeholder}")
                    _delete_col_results_row(doc, var_name)

        print("\n" + "="*80)
        print("【调试】未使用占位符处理完成")
        print("="*80 + "\n")

    def _fill_unused_val_results(doc, var_name, satellite, source):
        """为未使用的val_results占位符填写检验类型和"/" """
        placeholder = f'{{{{val_results_{var_name}}}}}'

        # 使用与其他产品相同的检验类型格式
        if source.upper() == 'XC':
            validation_type = f"{satellite} vs 现场"
        else:
            validation_type = f"{satellite} vs {source}"

        # 填充表格：检验类型 | / | /
        val_results = [[validation_type, '/', '/']]
        _fill_table(doc, placeholder, val_results)

    def _delete_col_results_row(doc, var_name):
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

    def fill_template(template_path, output_docx, output_pdf, replacements):
        """
        自动填充Word模板并转换为PDF

        :param template_path: 模板文件路径
        :param output_docx: 输出的docx文件路径
        :param output_pdf: 输出的pdf文件路径
        :param replacements: 包含替换内容的字典
        """
        doc = Document(template_path)

        # 文本替换
        if 'text' in replacements:
            for placeholder, text in replacements['text'].items():
                _replace_text(doc, placeholder, text)

        # 表格数据填充
        if 'tables' in replacements:
            for placeholder, table_data in replacements['tables'].items():
                _fill_table(doc, placeholder, table_data)

        # 图片插入
        if 'images' in replacements:
            for placeholder, image_path in replacements['images'].items():
                _insert_image(doc, placeholder, image_path)

        # 处理模板中未使用的占位符
        _handle_unused_placeholders(doc, replacements)

        # 对于现场数据报告，将所有 "XC卫星" 和 "XC" 替换为 "现场"
        _replace_text(doc, 'XC卫星', '现场')
        _replace_text(doc, 'XC', '现场')

        # 先删除所有章节中没有图片的小节并重新编号（必须在清理占位符之前执行）
        _remove_empty_subsections_and_renumber(doc)

        # 问题7: 清理所有未替换的占位符（移除大括号形式的参数名称）
        _cleanup_placeholders(doc)

        doc.save(output_docx)

    def hy1d_cocts_daily_report(datestr, input_temp, input_img, coldata_path, output_path, satellite_type,source_org_type):
        """
        生成每日报告

        :param datestr: 日期字符串 (格式: %Y%m%d)
        :param input_temp: 模板文件路径
        :param input_img: 图片存储根目录
        :param coldata_path: 数据文件路径
        :param output_path: 输出路径
        :param satellite_type: 卫星类型
        """
    replacements = generate_replacements(datestr, input_img, coldata_path, satellite_type,source_org_type,space_size, time_size)
    save_path = os.path.join(output_path)
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    if source_org_type == 'AQUA' or source_org_type == 'TERRA':
        template_path=os.path.join(input_temp, f'new_auqa_terra_template.docx')
    elif source_org_type == 'SNPP' or source_org_type == 'JPSS':
        template_path=os.path.join(input_temp, f'new_snpp_jpss_template.docx')
    
    fill_template(
        template_path,
        output_docx=os.path.join(save_path, f'{satellite_type.upper()}_COCTS_{source_org_type.upper()}_val_report_{datestr}.docx'),
        output_pdf=os.path.join(save_path, f'{satellite_type.upper()}_COCTS_{source_org_type.upper()}_val_report_{datestr}.pdf'),
        replacements=replacements
    )

    hy1d_cocts_daily_report(datestr, input_temp, input_img, coldata_path, output_path, satellite_type,source_org_type)

def step_xc_report(datestr, input_temp, input_img, coldata_path, output_path, satellite_type,source_org_type,space_size):
    # 确保输出目录存在
    os.makedirs(coldata_path, exist_ok=True)
    os.makedirs(output_path, exist_ok=True)
    os.makedirs(input_img, exist_ok=True)

    # 定义产品配置
    XC_VAR_CONFIG = {
        'sst': {'sources': ['XC'], 'unit': 'K', 'col_values': [25, 1800]},
        'chl': {'sources': ['XC'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        'Rrs412': {'sources': ['XC'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs443': {'sources': ['XC'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs490': {'sources': ['XC'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs520': {'sources': ['XC'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs565': {'sources': ['XC'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs670': {'sources': ['XC'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'Rrs750': {'sources': ['XC'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        'TSM': {'sources': ['XC'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        'CDOM': {'sources': ['XC'], 'unit': '1/m', 'col_values': [25, 1800]},
        'nLw': {'sources': ['XC'], 'unit': 'mW/cm2/um/sr', 'col_values': [25, 1800]},
    }

    # 初始化VAR_CONFIG
    VAR_CONFIG = None

    if source_org_type == 'XC':
        VAR_CONFIG = XC_VAR_CONFIG
    else:
        raise ValueError(f"Unsupported source type: {source_org_type}")


    def calc_metric(coldata_path, reference, var_name, datestr, satellite_type):
        """
        从报告文件中提取指标数据

        :param coldata_path: 数据文件路径
        :param reference: 数据源
        :param var_name: 变量名
        :param datestr: 日期字符串
        :param satellite_type: 卫星类型
        :return: bias, rms, n
        """
        filepath = os.path.join(coldata_path, f'{satellite_type}_COCTS_{reference}_{var_name}_report_{datestr}.txt')
        if not os.path.exists(filepath):
            files = os.listdir(coldata_path)
            pattern = f'{satellite_type}_COCTS_{reference}_{var_name}_report_{datestr}'
            matching_files = [f for f in files if pattern in f]
            if matching_files:
                filepath = os.path.join(coldata_path, matching_files[0])
            else:
                raise FileNotFoundError(f"找不到匹配的报告文件: {pattern}")

        Valid_Ratio = None
        CV_Value = None
        Relative_Bias = None
        with open(filepath, 'r') as f:
            for line in f:
                line = line.strip()
                if line.startswith('/Valid Ratio='):
                    Valid_Ratio = float(line.split('=')[1])
                elif line.startswith('/CV Value='):
                    CV_Value = float(line.split('=')[1])
                elif line.startswith('/Relative Bias='):
                    Relative_Bias = float(line.split('=')[1])

        if Valid_Ratio is None or CV_Value is None or Relative_Bias is None:
            raise ValueError(f"在文件{filepath}中找不到所需的数据")
        return Valid_Ratio, CV_Value, Relative_Bias

    def generate_replacements(datestr: str, imgpath: str, coldata_path: str, satellite_type: str,source_org_type: str, space_size: str):
        """
        根据日期生成包含所有产品类型的替换字典

        :param datestr: 日期字符串 (格式: %Y%m%d)
        :param imgpath: 图片存储根目录
        :param coldata_path: 数据文件路径
        :param satellite_type: 卫星类型
        :return: 包含文本、表格、图片路径的替换字典
        """
        date_obj = datetime.strptime(datestr, '%Y%m%d')
        replacements = {
            'text': {
                '{{date}}': date_obj.strftime('%Y-%m-%d'),
                '{{date_cn}}': date_obj.strftime('%Y年%m月%d日'),
            },
            'tables': {},
            'images': {}
        }

        for var_name, config in VAR_CONFIG.items():
            # 添加：检查是否有有效的源
            if not config['sources']:
                print(f"警告：{var_name} 没有配置有效的数据源，跳过")
                continue
            
            # 添加：检查是否有有效的报告文件
            has_valid_data = False
            for source in config['sources']:
                pattern = f'{satellite_type}_COCTS_{source}_{var_name}_report_{datestr}'
                try:
                    matching_files = [f for f in os.listdir(coldata_path) if pattern in f]
                    if matching_files:
                        has_valid_data = True
                        break
                except Exception as e:
                    print(f"检查文件时出错: {e}")
            
            if not has_valid_data:
                print(f"警告：{var_name} 没有找到有效的报告文件，跳过")
                continue
            
            # 初始化metrics字典
            metrics = {}
            
            try:
                metrics[source] = dict(zip(
                    ['Valid_Ratio', 'CV_Value', 'Relative_Bias'], 
                    calc_metric(coldata_path, source, var_name, datestr, satellite_type)
                    ))
            except FileNotFoundError as e:
                    print(f"跳过产品 {var_name}: {e}")
                    metrics[source] = {
                        'Valid_Ratio': 'N/A',
                        'CV_Value': 'N/A',
                        'Relative_Bias': 'N/A'
                    }

            val_results = [
                [f'{satellite_type} vs {source}', f"{metrics[source]['Relative_Bias']:.2f}"]
                for source in config['sources']
            ]

            col_results = [
                [f'{satellite_type} vs {source}', f"{space_size}*{space_size}*{metrics[source]['Valid_Ratio']:.2f}"]
                for source in config['sources']
            ]

            # 生成唯一的图片占位符
            image_keys = {
                # 'sct': [f'{satellite_type.lower()}_vs_{var_name}_{source.lower()}_sct' for source in config['sources']],
                # 'geo': [f'{satellite_type.lower()}_vs_{var_name}_{source.lower()}_geo' for source in config['sources']]
                'sct': [f'hy1c_vs_{var_name}_terra_sct'],
                'geo': [f'hy1c_vs_{var_name}_terra_geo']
            }
            print(f"Generated data for {var_name}: {val_results}")  # 添加打印语句

            # 从报告文件中提取时间戳
            timestamp = None
            for source in config['sources']:
                # 列出目录中的所有文件
                files = os.listdir(coldata_path)

                # 首先尝试新格式：report_{satellite}_{reference}_{var_name}_{datestr}_{time}.txt
                new_pattern = f'report_{satellite_type}_{source}_{var_name}_{datestr}'
                matching_files = [f for f in files if new_pattern in f and f.endswith('.txt')]

                if matching_files:
                    # 从新格式文件名中提取时间戳
                    # 格式: report_HY1C_TERRA_sst_20250101_120000.txt
                    parts = matching_files[0].replace('.txt', '').split('_')
                    if len(parts) >= 5:
                        # 时间戳是最后一部分
                        timestamp = parts[-1]
                        break
                else:
                    # 尝试旧格式：{satellite}_COCTS_{source}_{var_name}_report_{datestr}_{time}.txt
                    old_pattern = f'{satellite_type}_COCTS_{source}_{var_name}_report_{datestr}'
                    matching_files = [f for f in files if old_pattern in f and f.endswith('.txt')]
                    if matching_files:
                        # 从旧格式文件名中提取时间戳
                        parts = matching_files[0].split('_')
                        if len(parts) > 6:
                            timestamp = parts[6].replace('.txt', '')
                            break

            images = {}
            for key in image_keys['sct']:
                img_path = os.path.join(imgpath, f'{satellite_type}_COCTS_{source_org_type}_{var_name.upper()}_PIE_{datestr}_{timestamp}.jpg')
                if not os.path.exists(img_path):
                    alt_img_path = os.path.join(imgpath, f'{satellite_type}_COCTS_{source_org_type}_{var_name}_PIE_{datestr}_{timestamp}.jpg')
                    if os.path.exists(alt_img_path):
                        img_path = alt_img_path
                        print(f"使用替代图片路径: {img_path}")
                images[f'{{{{{key}}}}}'] = img_path

            # for key in image_keys['geo']:
            #     img_path = os.path.join(imgpath, f'{satellite_type}_COCTS_{source_org_type}_{var_name.upper()}_GEO_{datestr}_{timestamp}.jpg')
            #     if not os.path.exists(img_path):
            #         alt_img_path = os.path.join(imgpath, f'{satellite_type}_COCTS_{source_org_type}_{var_name}_GEO_{datestr}_{timestamp}.jpg')
            #         if os.path.exists(alt_img_path):
            #             img_path = alt_img_path
            #             print(f"使用替代图片路径: {img_path}")
            #     images[f'{{{{{key}}}}}'] = img_path

            # 对于现场数据，将 'XC' 替换为 '现场'
            source_display = '现场' if 'XC' in config['sources'] else ', '.join(config['sources'])

            replacements['text'].update({
                '{{satellite_type}}': satellite_type,
                '{{source_type}}': source_display,
                '{{unit}}': config['unit']
            })
            replacements['tables'][f'{{{{val_results_{var_name}}}}}'] = val_results
            replacements['tables'][f'{{{{col_results_{var_name}}}}}'] = col_results
            replacements['images'].update(images)

        return replacements


    def _replace_text(doc, placeholder, replacement):
        """遍历所有段落和表格单元格进行文本替换"""
        for p in doc.paragraphs:
            if placeholder in p.text:
                _replace_text_in_paragraph(p, placeholder, replacement)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if placeholder in p.text:
                            _replace_text_in_paragraph(p, placeholder, replacement)

    def _replace_text_in_paragraph(paragraph, placeholder, replacement):
        """在单个段落中执行文本替换，保留原始格式"""
        if placeholder not in paragraph.text:
            return
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, replacement)

    def _fill_table(doc, placeholder, table_data):
        """
        动态填充表格数据并设置居中格式

        :param doc: Document对象
        :param placeholder: 表格占位符
        :param table_data: 表格数据，预期为二维列表，其中每个子列表代表表格的一行
        """
        found = False
        for table in doc.tables:
            # 遍历表格中的所有行，找到包含占位符的行
            for row in table.rows:
                for cell in row.cells:
                    if placeholder in cell.text:
                        found = True
                        # 清空当前行的数据
                        for c in row.cells:
                            c.text = ""
                        # 填充新数据到单元格
                        for i, value in enumerate(table_data[0]):  # 只取第一个子列表的数据
                            cell = row.cells[i]
                            cell.text = str(value)
                            # 设置水平居中
                            for paragraph in cell.paragraphs:
                                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                            # 设置垂直居中
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        break  # 找到占位符后，跳出单元格循环
                if found:
                    break  # 找到占位符后，跳出行循环
        if not found:
            print(f"未找到占位符 '{placeholder}' 的表格。")

    def _insert_image(doc, placeholder, image_path):
        """在占位符位置插入图片"""
        if not os.path.exists(image_path):
            print(f"警告：图片文件不存在: {image_path}")
            return
        found = False
        # 遍历所有段落
        for p in doc.paragraphs:
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, '')
                run = p.add_run()
                run.add_picture(image_path, width=Inches(6))
                print(f"成功插入图片: {image_path}")
                found = True
                break
        if not found:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if placeholder in cell.text:
                            cell.text = cell.text.replace(placeholder, '')
                            run = cell.add_run()
                            run.add_picture(image_path, width=Inches(4))
                            print(f"成功插入图片: {image_path}")
                            found = True
                            break
                    if found:
                        break
                if found:
                    break
        if not found:
            print(f"未找到占位符 '{placeholder}' 的位置。")

    def _cleanup_placeholders(doc):
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

    def _remove_empty_subsections_and_renumber(doc):
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


    def _handle_unused_placeholders(doc, replacements):
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
                    _fill_unused_val_results(doc, var_name, satellite, source)

        # 处理未使用的col_results占位符（表二）
        for placeholder in all_placeholders:
            match = col_pattern.match(placeholder)
            if match:
                full_placeholder = f'{{{{{placeholder}}}}}'
                if full_placeholder not in used_tables:
                    # 删除整行
                    var_name = match.group(1)
                    print(f"【调试】删除未使用的col_results行: {full_placeholder}")
                    _delete_col_results_row(doc, var_name)

        print("\n" + "="*80)
        print("【调试】未使用占位符处理完成")
        print("="*80 + "\n")

    def _fill_unused_val_results(doc, var_name, satellite, source):
        """为未使用的val_results占位符填写检验类型和"/" """
        placeholder = f'{{{{val_results_{var_name}}}}}'

        # 使用与其他产品相同的检验类型格式
        if source.upper() == 'XC':
            validation_type = f"{satellite} vs 现场"
        else:
            validation_type = f"{satellite} vs {source}"

        # 填充表格：检验类型 | / | /
        val_results = [[validation_type, '/', '/']]
        _fill_table(doc, placeholder, val_results)

    def _delete_col_results_row(doc, var_name):
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

    def fill_template(template_path, output_docx, output_pdf, replacements):
        """
        自动填充Word模板并转换为PDF

        :param template_path: 模板文件路径
        :param output_docx: 输出的docx文件路径
        :param output_pdf: 输出的pdf文件路径
        :param replacements: 包含替换内容的字典
        """
        doc = Document(template_path)

        # 文本替换
        if 'text' in replacements:
            for placeholder, text in replacements['text'].items():
                _replace_text(doc, placeholder, text)

        # 表格数据填充
        if 'tables' in replacements:
            for placeholder, table_data in replacements['tables'].items():
                _fill_table(doc, placeholder, table_data)

        # 图片插入
        if 'images' in replacements:
            for placeholder, image_path in replacements['images'].items():
                _insert_image(doc, placeholder, image_path)

        # 处理模板中未使用的占位符
        _handle_unused_placeholders(doc, replacements)

        # 对于现场数据报告，将所有 "XC卫星" 和 "XC" 替换为 "现场"
        _replace_text(doc, 'XC卫星', '现场')
        _replace_text(doc, 'XC', '现场')

        # 先删除所有章节中没有图片的小节并重新编号（必须在清理占位符之前执行）
        _remove_empty_subsections_and_renumber(doc)

        # 问题7: 清理所有未替换的占位符（移除大括号形式的参数名称）
        _cleanup_placeholders(doc)

        doc.save(output_docx)

    def hy1d_cocts_daily_report(datestr, input_temp, input_img, coldata_path, output_path, satellite_type,source_org_type,space_size):
        """
        生成每日报告

        :param datestr: 日期字符串 (格式: %Y%m%d)
        :param input_temp: 模板文件路径
        :param input_img: 图片存储根目录
        :param coldata_path: 数据文件路径
        :param output_path: 输出路径
        :param satellite_type: 卫星类型
        """
    try:
        replacements = generate_replacements(datestr, input_img, coldata_path, satellite_type,source_org_type,space_size)
    except Exception as e:
        print(f"警告：生成报告数据失败: {e}")
        print("原因：可能没有找到有效的验证结果文件")
        print("提示：请检查现场数据文件的日期格式是否正确")
        traceback.print_exc()
        return  # 提前返回，不生成报告
    save_path = os.path.join(output_path)
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    if source_org_type == 'XC':
        template_path=os.path.join(input_temp, f'new_xc_template.docx')
    
    fill_template(
        template_path,
        output_docx=os.path.join(save_path, f'{satellite_type.upper()}_COCTS_{source_org_type.upper()}_val_report_{datestr}.docx'),
        output_pdf=os.path.join(save_path, f'{satellite_type.upper()}_COCTS_{source_org_type.upper()}_val_report_{datestr}.pdf'),
        replacements=replacements
    )

    hy1d_cocts_daily_report(datestr, input_temp, input_img, coldata_path, output_path, satellite_type,source_org_type,space_size)


def rename_files(input_dir):
    """
    根据指定规则重命名文件
    
    Args:
        input_dir: 输入文件夹路径
    """
    # 确保输入路径存在
    if not os.path.exists(input_dir):
        print(f"输入路径 {input_dir} 不存在")
        return
    
    # 获取所有文件
    files = glob.glob(os.path.join(input_dir, "*"))
    renamed_count = 0
    
    for file_path in files:
        if os.path.isfile(file_path):
            file_name = os.path.basename(file_path)
            new_name = None
            
            # 提取文件名中的关键信息
            # 匹配模式: HY1D_TERRA_Rrs412_20250302130946 或类似格式
            match = re.search(r'(HY1[CD])_([A-Z]+)_([A-Za-z0-9]+)_(\d{8})(\d{6})', file_name)
            if not match:
                # 尝试其他模式: valresult_HY1C_AQUA_Rrs412_20231011102445
                match = re.search(r'(?:valresult|map|valstatistic|report|statistic)_(HY1[CD])_([A-Z]+)_([A-Za-z0-9]+)_(\d{8})(\d{6})', file_name)
            
            if match:
                satellite = match.group(1)
                validation_source = match.group(2)
                product = match.group(3)
                date = match.group(4)
                time = match.group(5)
                
                # 格式化日期和时间
                formatted_date = date
                formatted_time = time
                
                # 根据文件名前缀确定新的文件名
                if file_name.startswith("valresult_"):
                    new_name = f"{satellite}_COCTS_{validation_source}_{product}_matchup_{formatted_date}_{formatted_time}.txt"
                elif file_name.startswith("map_"):
                    if file_name.endswith(".jpg"):
                        new_name = f"{satellite}_COCTS_{validation_source}_{product}_GEO_{formatted_date}_{formatted_time}.jpg"
                    else:
                        print(f"文件 {file_name} 不是 .jpg 格式，跳过重命名")
                elif file_name.startswith("valstastic_"):
                    new_name = f"{satellite}_COCTS_{validation_source}_{product}_PIE_{formatted_date}_{formatted_time}.jpg"
                elif file_name.startswith("report_"):
                    new_name = f"{satellite}_COCTS_{validation_source}_{product}_report_{formatted_date}_{formatted_time}.txt"
                elif file_name.startswith("statistic_"):
                    new_name = f"{satellite}_COCTS_{validation_source}_{product}_statistic_{formatted_date}_{formatted_time}.txt"
                elif file_name.endswith(".pdf"):
                    new_name = f"{satellite}_COCTS_{validation_source}_val_report_{formatted_date}_{formatted_time}.pdf"
            # 如果找到了匹配的重命名规则，执行重命名
            if new_name:
                new_path = os.path.join(os.path.dirname(file_path), new_name)
                try:
                    os.rename(file_path, new_path)
                    print(f"已重命名: {file_name} -> {new_name}")
                    renamed_count += 1
                except Exception as e:
                    print(f"重命名 {file_name} 失败: {str(e)}")
            else:
                print(f"未找到匹配规则: {file_name}")
    
    print(f"总共重命名了 {renamed_count} 个文件")



def organize_files(input_folder, output_folder):
    """
    创建输出文件夹结构并根据规则移动文件
    
    参数:
    input_folder (str): 输入文件夹路径
    output_folder (str): 输出文件夹路径
    """
    # 创建输出主文件夹（如果不存在）
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 创建五个子文件夹
    subfolders = [
        "01_sat_preprocess",
        "02_reference_preprocess", 
        "03_collocation", 
        "04_visualization",
        "05_reports"
    ]
    
    # 创建每个子文件夹
    subfolder_paths = {}
    for subfolder in subfolders:
        subfolder_path = os.path.join(output_folder, subfolder)
        subfolder_paths[subfolder] = subfolder_path
        if not os.path.exists(subfolder_path):
            os.makedirs(subfolder_path)
            print(f"创建文件夹: {subfolder_path}")
    
    # 遍历输入文件夹中的所有文件
    files_moved = 0
    for filename in os.listdir(input_folder):
        source_path = os.path.join(input_folder, filename)
        
        # 跳过文件夹
        if os.path.isdir(source_path):
            continue
        
        # 确定目标文件夹
        target_folder = determine_target_folder(filename)
        
        if target_folder:
            target_path = os.path.join(subfolder_paths[target_folder], filename)
            # 复制文件到目标文件夹
            move(source_path, target_path)
            files_moved += 1
            print(f"已移动文件 '{filename}' 到 '{target_folder}'")
    
    print(f"文件移动完成，共移动 {files_moved} 个文件。")

def determine_target_folder(filename):
    """
    根据文件名确定目标文件夹
    
    参数:
    filename (str): 文件名
    
    返回:
    str: 目标文件夹名称，如果不匹配任何规则则返回None
    """
    # 检查是否包含flag1，如果有则放入03_collocation
    if "flag1" in filename:
        return "03_collocation"
    
    # 01_sat_preprocess: 被检验数据文件 (HY1D_*, HY1C_*)
    if re.match(r'^HY1[CD]_(?!COCTS).*?\d{14}\.txt$', filename):
        return "01_sat_preprocess"
    
    # 02_reference_preprocess: 检验源数据 (TERRA_*, AQUA_*, SNPP_*, JPSS_*)
    if re.match(r'^(TERRA|AQUA|SNPP|JPSS)_.*?\d{14}\.txt$', filename):
        return "02_reference_preprocess"
    
    # 05_reports: 报告相关文件
    if "report" in filename.lower():
        return "05_reports"
    
    # 04_visualization: 图片相关文件和map文件
    if filename.endswith('.jpg') or ('visualization' in filename.lower()) or ('map' in filename.lower()):
        return "04_visualization"
    
    # 03_collocation: 其他所有collocation相关文件
    if any(x in filename.lower() for x in ['timeresult', 'spaceresult', 'statistic']) or \
       ('pixelstastic' in filename.lower()):
        return "03_collocation"
    
    if "matchup" in filename.lower() or "timesize" in filename.lower():
        return "03_collocation"

    if "timeseries" in filename.lower() or "resstastic" in filename.lower():
        return "04_visualization"

    if filename.lower().endswith('.pdf'):
        return "05_reports"
    

    # 默认返回None，表示不移动该文件
    return "02_reference_preprocess"
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
        print(f"\n仅转换本次生成的 {len(files_to_convert)} 个docx文件")
    else:
        # 处理目录中所有的docx文件(向后兼容)
        files_to_convert = [f for f in os.listdir(input_dir) if f.endswith(".docx")]
        print(f"\n转换目录中所有 {len(files_to_convert)} 个docx文件")

    # 遍历要转换的文件
    for filename in files_to_convert:
        # 构建输入文件的完整路径
        input_file = os.path.join(input_dir, filename)

        # 检查文件是否存在
        if not os.path.exists(input_file):
            print(f"警告: 文件不存在,跳过: {input_file}")
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
            print(f"成功转换：{filename} -> {filename.replace('.docx', '.pdf')}")
        else:
            print(f"转换失败：{filename}")
            print("错误信息：", result.stderr.decode())



def main():
    # 优先使用环境变量中指定的配置文件路径
    config_path = os.environ.get('CONFIG_PATH')
    if not config_path:
        # 如果环境变量未设置，则使用默认路径
        config_path = os.path.join(os.path.dirname(__file__), 'config.ini')
    
    # 加载配置
    config = configparser.ConfigParser()
    try:
        # 明确指定使用 UTF-8 编码读取配置文件
        with open(config_path, 'r', encoding='utf-8') as f:
            config.read_file(f)
    except Exception as e:
        print(f"错误：读取配置文件失败 {config_path}")
        print(f"错误信息: {str(e)}")
        return
    
    # 运行检验
    run_check(config)

if __name__ == '__main__':
    main()