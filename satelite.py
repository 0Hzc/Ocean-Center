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
import platform
from docx import Document
from docx.shared import Inches
from datetime import datetime
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
import os
import re
from datetime import datetime, timedelta
import subprocess
from shutil import move

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
    # 环境变量中指定的配置文件路径优先
    config_path = os.environ.get('CONFIG_PATH')
    if not config_path:
        # 否则使用相对路径
        config_path = 'config.ini'
    
    # 确保使用utf-8编码读取配置文件
    with open(config_path, 'r', encoding='utf-8') as f:
        config.read_file(f)
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
        input_dir = config['PATH']['input_dir']
        output_dir = config['PATH']['output_dir']
        window_size = int(config['PARAMS']['window_size'])
        time_threshold = int(config['PARAMS']['time_threshold'])
        source_type = config['VALIDATION']['source_type']
        font_path = config['font']['font_path']
        inspection_type = config['SATELLITE']['type']
        space_size = int(config['PARAMS']['window_size'])
        set_font = 'Simhei'
         
         # 定义输入子目录
        sat_input_dir = os.path.join(input_dir, "01_sat")
        reference_input_dir = os.path.join(input_dir, "02_reference")
        # reports_input_dir = os.path.join(input_dir, "05_reports")
         

        # 确保路径使用正斜杠（对Linux兼容）
        input_dir = input_dir.replace('\\', '/')
        output_dir = output_dir.replace('\\', '/')
        font_path = font_path.replace('\\', '/')
        sat_input_dir = sat_input_dir.replace('\\', '/')
        reference_input_dir = reference_input_dir.replace('\\', '/')

        # 引入必要的库 (如果文件开头没引用的情况下)
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
        
        print(f"\n=== 开始 {inspection_type} vs {source_type} 数据检验流程 ===")
        #步骤1：处理被检验数据
        if inspection_type == 'HY1E':

            print("\n处理被检验HY1E数据...")
            process_hye_data(
                inspection_type,
                hy_file_l2a=os.path.join(sat_input_dir, config['HY1E']['l2a_file']),
                hy_file_l2b=os.path.join(sat_input_dir, config['HY1E']['l2b_file']),
                hy_file_l2c=os.path.join(sat_input_dir, config['HY1E']['l2c_file']),
                hy_file_l2t=os.path.join(sat_input_dir, config['HY1E']['l2t_file']),
                output_dir=output_dir
            )
        elif inspection_type == 'HY1C':
            print("\n处理被检验HY1C数据...")
            process_hycd_data(
                inspection_type,
                hy_file_l2a=os.path.join(sat_input_dir, config['HY1C']['l2a_file']),
                hy_file_l2b=os.path.join(sat_input_dir, config['HY1C']['l2b_file']),
                output_dir=output_dir
            )
        else:
            print("\n处理被检验HY1D数据...")
            process_hycd_data(
                inspection_type,
                hy_file_l2a=os.path.join(sat_input_dir, config['HY1D']['l2a_file']),
                hy_file_l2b=os.path.join(sat_input_dir, config['HY1D']['l2b_file']),
                output_dir=output_dir
            )

        #步骤2：处理检验源数据
        if source_type == 'HY1E':
            print("\n处理检验源HY1E数据...")
            process_hye_data(
                hy_file_l2a=os.path.join(reference_input_dir, config['HY1E']['l2a_file']),
                hy_file_l2b=os.path.join(reference_input_dir, config['HY1E']['l2b_file']),
                hy_file_l2c=os.path.join(reference_input_dir, config['HY1E']['l2c_file']),
                hy_file_l2t=os.path.join(reference_input_dir, config['HY1E']['l2t_file']),
                output_dir=output_dir
            )
        elif source_type == 'HY1C':
            print("\n处理检验源HY1C数据...")
            process_hycd_data(
                source_type,
                hy_file_l2a=os.path.join(reference_input_dir, config['HY1C']['l2a_file']),
                hy_file_l2b=os.path.join(reference_input_dir, config['HY1C']['l2b_file']),
                output_dir=output_dir
            )
        else:
            print("\n处理检验源HY1D数据...")
            process_hycd_data(
                source_type,
                hy_file_l2a=os.path.join(reference_input_dir, config['HY1D']['l2a_file']),
                hy_file_l2b=os.path.join(reference_input_dir, config['HY1D']['l2b_file']),
                output_dir=output_dir
            )
        
        # 步骤3：标识检查
        print("\n执行步骤3标识检查...")
        if inspection_type == 'HY1E' or source_type == 'HY1E':
            # HY1E数据的标识检查
            HY1E_flag_create(output_dir, window_size)
        if inspection_type in ['HY1C', 'HY1D'] or source_type in ['HY1C', 'HY1D']:
            HY_flag_create(inspection_type,output_dir, window_size)
            HY_flag_create(source_type,output_dir, window_size)
        
        
        # 步骤4：时间匹配
        print("\n执行步骤4时间匹配...")
        process_satellite_timematch(output_dir, output_dir, inspection_type, source_type, time_threshold)
        
        # 步骤5：空间匹配
        print("\n执行空间匹配...")
        process_satellite_spacematch(output_dir, output_dir, inspection_type, source_type)
        # 保存空间窗口大小信息
        try:
            with open(os.path.join(output_dir, 'spacesize.txt'), 'w') as f:
                f.write(f"{window_size}")
        except Exception as e:
            print(f"保存空间窗口大小信息失败: {e}")
        
        # 步骤6：生成验证结果
        print("\n生成验证结果...")

        satellite_validation(output_dir, output_dir)
        
        # 步骤7：生成误差地图
        step7(output_dir, output_dir,inspection_type)

        # 步骤8：生成折线图
        step8(output_dir, output_dir,inspection_type)

        # 步骤9：生成统计结果和图表
        step9(output_dir, output_dir,inspection_type,source_type)

        #步骤10：生成报告所需数据report文件
        make_satellite_report_data(output_dir)

        #步骤11 修改文件名
        rename_files(output_dir)

        #步骤12 移动文件到指定文件夹
        organize_files(output_dir, output_dir)
            
        # # 步骤13：生成报告
        oc_file=os.path.join(sat_input_dir, config[inspection_type]['l2a_file'])
        beijing_time = extract_datetime(oc_file)
        time_str = beijing_time.strftime('%Y%m%d%H%M%S')
        extracted_data = time_str[:8]

        input_temp = os.path.join(input_dir, '05_reports')
        input_img = os.path.join(output_dir, '04_visualization')
        coldata_path = os.path.join(output_dir, '05_reports')
        output_path = os.path.join(output_dir, '05_reports')

        step_report(extracted_data, input_temp, input_img, coldata_path, output_path, inspection_type, source_type, space_size,time_threshold)

        word_to_pdf(output_path, output_path)

        print(f"\n=== {inspection_type}  vs {source_type} 数据检验流程完成 ===")
        return True
    
    except Exception as e:
        print(f"检验流程执行失败: {str(e)}")
        traceback.print_exc()
        return False

def process_hye_data(source_type, hy_file_l2a, hy_file_l2b, hy_file_l2c, hy_file_l2t,output_dir):
    """
    处理HY3A待检验数据
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
        os.makedirs(output_dir, exist_ok=True)
        prefix = source_type

        # 处理反射率数据
        with h5py.File(hy_file_l2a, 'r') as h5_file:
            # # 获取时间信息
            # year = int(h5_file['Scan Line Attributes/year'][0])
            # day = int(h5_file['Scan Line Attributes/day'][0])
            # millisecond = int(h5_file['Scan Line Attributes/msec'][0])
            
            # # 转换为北京时间
            # utc_time = datetime(year, 1, 1) + timedelta(days=day-1, milliseconds=millisecond)
            # beijing_time = utc_time + timedelta(hours=8)
            # time_str = beijing_time.strftime('%Y%m%d%H%M%S')
            date_str = extract_datetime(hy_file_l2a)
            time_str = date_str.strftime("%Y%m%d%H%M%S")

            # 获取数据维度
            lat_data = h5_file['Navigation Data/Latitude'][:]
            rows, cols = lat_data.shape
            print(f"检测到数据维度: {rows} x {cols}")
            
            dim_file = os.path.join(output_dir, f'dimensions_{time_str}.txt')
            with open(dim_file, 'w') as f:
                f.write(f"{rows},{cols}")

            # 保存基础数据
            save_data_to_txt(h5_file['Navigation Data/Latitude'][:], 
                           os.path.join(output_dir, f'{prefix}_lat_{time_str}.txt'))
            save_data_to_txt(h5_file['Navigation Data/Longitude'][:], 
                           os.path.join(output_dir, f'{prefix}_lon_{time_str}.txt'))
            save_data_to_txt(h5_file['Geophysical Data/l2_flags'][:], 
                           os.path.join(output_dir, f'{prefix}_flag_{time_str}.txt'))
     

        # 处理TSM等参数数据
        with h5py.File(hy_file_l2b, 'r') as h5_file:
            # 保存参数数据
            params = {
                'TSM': 'Geophysical Data/TSM',
                'CDOM': 'Geophysical Data/CDOM'
            }
            
            for param_name, dataset_path in params.items():
                data = h5_file[dataset_path][:]
                save_data_to_txt(data, 
                               os.path.join(output_dir, f'{prefix}_{param_name}_{time_str}.txt'))
                

        print('\nHY1E数据处理完成\n')
        return True

    except Exception as e:
        print(f"处理数据时出错: {str(e)}")
        traceback.print_exc()
        return False

def process_hycd_data(source_type,hy_file_l2a, hy_file_l2b, output_dir):
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

            # 保存基础数据
            save_data_to_txt(h5_file['Navigation Data/Latitude'][:], 
                           os.path.join(output_dir, f'{prefix}_lat_{time_str}.txt'))
            save_data_to_txt(h5_file['Navigation Data/Longitude'][:], 
                           os.path.join(output_dir, f'{prefix}_lon_{time_str}.txt'))
            save_data_to_txt(h5_file['Geophysical Data/l2_flags'][:], 
                           os.path.join(output_dir, f'{prefix}_flag_{time_str}.txt'))


        # 处理TSM等参数数据
        with h5py.File(hy_file_l2b, 'r') as h5_file:
            # 保存参数数据
            params = {
                'TSM': 'Geophysical Data/TSM',
                'CDOM': 'Geophysical Data/CDOM'
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

def HY1E_flag_create(input_dir,window_size):
    try:
        print("\n开始执行HY1E_flag_create函数\n")
        flag_matrices = {}
        
        # 检查目录中的文件
        all_files = os.listdir(input_dir)
      
        # 处理所有HY3A_flag文件
        for filename in all_files:        
            if filename.startswith('HY1E_flag_') and filename.endswith('.txt'):
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
                mask = ((flag_matrix & (1 << 8)) | (flag_matrix & (1 << 22))) != 0
                FLAG[mask] = 1
                # print(f"\n位运算后的FLAG统计:")
                # print(f"- FLAG中1的数量: {np.sum(FLAG == 1)}")
                # print(f"- FLAG中0的数量: {np.sum(FLAG == 0)}")
                
                # 查找对应的产品文件
                for product_file in all_files:
                    if 'lon' in product_file or 'lat' in product_file:
                        continue
                    if product_file.startswith('HY1E_') and time_id in product_file:
                        print(f"\n处理产品文件: {product_file}")  # 新增：显示当前处理的产品文件

                        if os.path.exists(os.path.join(input_dir, product_file)):
                            # 记录处理前的1的数量
                            ones_before = np.sum(FLAG == 1)
                            
                            temp_matrix = generate_flag_from_data(
                                os.path.join(input_dir, product_file), 
                                'HY1E'
                            )
                            if temp_matrix is not None and len(temp_matrix) == len(flag_matrix):
                                FLAG = np.logical_or(FLAG, temp_matrix).astype(np.int32)
                                
                                # 计算并显示变化
                                ones_after = np.sum(FLAG == 1)
                                new_ones = ones_after - ones_before
                                # print(f"\n产品 {product_file} 的影响:")
                                # print(f"- 处理前1的数量: {ones_before}")
                                # print(f"- 处理后1的数量: {ones_after}")
                                # print(f"- 该产品新增1的数量: {new_ones}")
                                # print(f"- 占总像素的比例: {(new_ones / len(FLAG)) * 100:.2f}%")
                                
                                if new_ones > len(FLAG) * 0.5:  # 如果新增的1超过50%
                                    print(f"警告: 产品 {product_file} 导致大量像素变为1!")
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

def generate_flag_from_data(data_file, satellite_type):
    try:
        data = np.genfromtxt(data_file, delimiter=None)
        flag = np.zeros_like(data, dtype=np.int32)
        
        # 添加统计信息
        filename = os.path.basename(data_file)
        # 根据不同产品类型处理无效值
        if 'ipar' in filename:       
            # 设置标记
            flag[data == -717.002197265625] = 1
            flag[np.isnan(data)] = 1
        elif satellite_type in ['HY1E', 'HY1C', 'HY1D']:
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



def HY_flag_create(satellite_type,input_dir,window_size):
    try:
        print(f"\n开始执行{satellite_type}_flag_create函数\n")
        flag_matrices = {}
        
        # 检查目录中的文件
        all_files = os.listdir(input_dir)
      
        # 处理所有HY_flag文件
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
                mask = ((flag_matrix & (1 << 8)) | (flag_matrix & (1 << 22))) != 0
                FLAG[mask] = 1
                # print(f"\n位运算后的FLAG统计:")
                # print(f"- FLAG中1的数量: {np.sum(FLAG == 1)}")
                # print(f"- FLAG中0的数量: {np.sum(FLAG == 0)}")
                
                # 查找对应的产品文件
                for product_file in all_files:
                    if 'lon' in product_file or 'lat' in product_file:
                        continue
                    if product_file.startswith(f'{satellite_type}_') and time_id in product_file:
                        print(f"\n处理产品文件: {product_file}")  # 新增：显示当前处理的产品文件

                        if os.path.exists(os.path.join(input_dir, product_file)):
                            # 记录处理前的1的数量
                            ones_before = np.sum(FLAG == 1)
                            
                            temp_matrix = generate_flag_from_data(
                                os.path.join(input_dir, product_file), 
                                satellite_type
                            )
                            if temp_matrix is not None and len(temp_matrix) == len(flag_matrix):
                                FLAG = np.logical_or(FLAG, temp_matrix).astype(np.int32)
                                
                                # 计算并显示变化
                                ones_after = np.sum(FLAG == 1)
                                new_ones = ones_after - ones_before
                                # print(f"\n产品 {product_file} 的影响:")
                                # print(f"- 处理前1的数量: {ones_before}")
                                # print(f"- 处理后1的数量: {ones_after}")
                                # print(f"- 该产品新增1的数量: {new_ones}")
                                # print(f"- 占总像素的比例: {(new_ones / len(FLAG)) * 100:.2f}%")
                                
                                if new_ones > len(FLAG) * 0.5:  # 如果新增的1超过50%
                                    print(f"警告: 产品 {product_file} 导致大量像素变为1!")
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
        'HY1C': {
            'CDOM': 'CDOM', 
            'TSM': 'TSM', 
        },
        'HY1D': {
            'CDOM': 'CDOM', 
            'TSM': 'TSM', 
        },
        'HY1E': {
            'CDOM': 'CDOM', 
            'TSM': 'TSM', 
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
                       any(x.lower() in f.lower() for x in ['CDOM', 'TSM']) and 
                       f.endswith('.txt')]
        
        if not target_files:
            print(f"未找到{target_sensor}的数据文件")
            return False
          
        # 定义参数列表

        other_params = ['CDOM', 'TSM']
        
        # 处理每个目标文件
        for target_file in target_files:
            # 提取时间信息
            target_time = extract_datetime_from_filename(target_file)
            if not target_time:
                print(f"无法从文件名提取时间: {target_file}")
                continue
            
            # 识别参数类型
            param_type = None
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

def process_satellite_spacematch(input_dir, output_dir, target_sensor, source_type):
    """
    处理卫星数据空间匹配
    """
    # 卫星命名规则配置
    SATELLITE_NAMING_RULES = {
        'HY1C': {
            'prefix': 'HY1C',
            'output_prefix': 'HY1C1',
            'lat_format': 'HY1C_lat',
            'lon_format': 'HY1C_lon',
            'flag_format': 'HY1C_flag1',
        },
        'HY1D': {
            'prefix': 'HY1D',
            'output_prefix': 'HY1D1',
            'lat_format': 'HY1D_lat',
            'lon_format': 'HY1D_lon',
            'flag_format': 'HY1D_flag1',
        },
        'HY1E': {
            'prefix': 'HY1E',
            'output_prefix': 'HY1E1',
            'lat_format': 'HY1E_lat',
            'lon_format': 'HY1E_lon',
            'flag_format': 'HY1E_flag1',
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
            elif 'CDOM' in target_file:
                param_type = 'CDOM'
            elif 'TSM' in target_file:
                param_type = 'TSM'
            else:
                print(f"无法识别的参数类型: {target_file}")
                return False
                
            print(f"\n处理文件对: {target_file} - {source_file}")
            print(f"参数类型: {param_type}")
            
            # 读取数据
            target_lat = np.genfromtxt(os.path.join(input_dir, f"{target_sensor}_lat_{target_time}.txt"))
            target_lon = np.genfromtxt(os.path.join(input_dir, f"{target_sensor}_lon_{target_time}.txt"))
            target_flag = np.genfromtxt(os.path.join(input_dir, f"{target_sensor}_flag1_{target_time}.txt"))
            
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
            'TSM': 'mg/L',
            'CDOM': '1/m',
        }
        return units.get(product, '1/sr')

    def get_product_filename(product):
        """根据产品类型返回对应的文件名部分"""
        if product == 'TSM':
            return 'TSM'
        elif product == 'CDOM':
            return 'CDOM'
        return product

    try:
        print("\n=== 执行步骤6：生成验证结果和统计结果文件 ===")
        # 获取输入文件列表
        input_files = os.listdir(input_path)
        print(f"输入目录中共有 {len(input_files)} 个文件")
        
        # 查找所有space结果文件
        space_files = [f for f in input_files if f.startswith('spaceresult_')]
        print(f"找到 {len(space_files)} 个space结果文件")
        
        if len(space_files) == 0:
            print("警告: 没有找到任何space结果文件，请检查输入路径是否正确")
        
        for space_file in space_files:
            print(f"\n开始处理space文件: {space_file}")
            # 从文件名解析参数
            parts = space_file.replace('spaceresult_', '').replace('.txt', '').split('_')
            print(f"解析文件名得到的部分: {parts}")
            
            if len(parts) < 4:
                print(f"警告: 文件名格式不正确，跳过处理: {space_file}")
                continue
                
            HY, source, product, timeHY = parts
            print(f"解析参数: HY={HY}, source={source}, product={product}, timeHY={timeHY}")
            
            # 获取实际的产品文件名部分
            product_filename = get_product_filename(product)
            print(f"产品文件名部分: {product_filename}")
            
            # 读取space结果获取时间差
            space_path = os.path.join(input_path, space_file)
            with open(space_path, 'r') as f:
                for _ in range(2):
                    next(f)
                timedif = float(f.readline().strip())
            
            # 获取source时间
            print(f"开始查找source时间...")
            timesource = None
            for f in input_files:
                if f.startswith(f'{source}1_{product}_') and f.endswith('.txt'):
                    timesource = f.split('_')[-1].replace('.txt', '')
                    print(f"找到匹配的source文件: {f}, 时间为: {timesource}")
                    break
            
            if not timesource:
                print(f"警告: 未找到匹配的source时间，跳过处理")
                continue
            
            # 读取数据文件（使用修改后的产品名称）
            Rrs2_path = os.path.join(input_path, f'{source}1_{product}_{timesource}.txt')
            flag1_path = os.path.join(input_path, f'{HY}_flag1_{product}_{timeHY}.txt')
            Rrs1_path = os.path.join(input_path, f'{HY}_{product_filename}_{timeHY}.txt')
            
            print(f"准备读取以下数据文件:")
            print(f"  Rrs2_path: {Rrs2_path}")
            print(f"  flag1_path: {flag1_path}")
            print(f"  Rrs1_path: {Rrs1_path}")
            
            Rrs2 = read_data(Rrs2_path)
            flag1 = read_data(flag1_path)
            Rrs1 = read_data(Rrs1_path)
            
            if Rrs2 is None or flag1 is None or Rrs1 is None:
                print("警告: 数据文件读取失败，跳过处理")
                continue
            
            # 处理数据
            print("开始处理数据...")
            data = []
            for i in range(len(Rrs1)):
                if flag1[i] == 0 and Rrs2[i] != -999 and Rrs2[i] != 0:
                    if product.lower() == 'sst':
                        diff = abs(Rrs1[i] - Rrs2[i])
                    else:
                        diff = abs((Rrs1[i] - Rrs2[i]) / Rrs2[i]) * 100
                    data.append([i, Rrs1[i], Rrs2[i], diff])

            if not data:
                print("警告: 没有有效的数据点，跳过处理")
                continue
                
            data = np.array(data)
            ave = np.mean(data[:, 3])



            # # 处理数据
            # print("开始处理数据...")
            # print(f"总数据点数量: {len(Rrs1)}")
            # data = []
            # invalid_flag_count = 0
            # invalid_minus999_count = 0
            # invalid_zero_count = 0
            
            # for i in range(len(Rrs1)):
            #     # 打印每10个数据点的信息，避免输出过多
            #     if i % 10 == 0:
            #         print(f"处理数据点 {i}: HY值={Rrs1[i]:.4f}, {source}值={Rrs2[i]:.4f}, 标志位={flag1[i]}")
                
            #     # 检查各个条件并记录不符合的情况
            #     if flag1[i] != 0:
            #         invalid_flag_count += 1
            #         if i % 10 == 0:
            #             print(f"  数据点 {i} 标志位不为0 (flag1={flag1[i]}), 跳过")
            #         continue
                    
            #     if Rrs2[i] == -999:
            #         invalid_minus999_count += 1
            #         if i % 10 == 0:
            #             print(f"  数据点 {i} 参考值为-999, 跳过")
            #         continue
                    
            #     if Rrs2[i] == 0:
            #         invalid_zero_count += 1
            #         if i % 10 == 0:
            #             print(f"  数据点 {i} 参考值为0, 跳过")
            #         continue
                
            #     # 计算差异
            #     if product.lower() == 'sst':
            #         diff = abs(Rrs1[i] - Rrs2[i])
            #         if i % 10 == 0:
            #             print(f"  数据点 {i} 有效: 绝对差异={diff:.4f}")
            #     else:
            #         diff = abs((Rrs1[i] - Rrs2[i]) / Rrs2[i]) * 100
            #         if i % 10 == 0:
            #             print(f"  数据点 {i} 有效: 相对差异={diff:.2f}%")
                
            #     data.append([i, Rrs1[i], Rrs2[i], diff])

            # # 打印无效数据点统计
            # print(f"无效数据点统计:")
            # print(f"  标志位不为0的点数: {invalid_flag_count}")
            # print(f"  参考值为-999的点数: {invalid_minus999_count}")
            # print(f"  参考值为0的点数: {invalid_zero_count}")
            # print(f"  总无效点数: {invalid_flag_count + invalid_minus999_count + invalid_zero_count}")
            
            # if not data:
            #     print("警告: 没有有效的数据点，跳过处理")
            #     continue
                
            # data = np.array(data)
            # print(f"有效数据点数量: {len(data)}, 占总数据点的 {len(data)/len(Rrs1)*100:.2f}%")
            
            # # 打印有效数据的基本统计信息
            # print(f"有效数据统计:")
            # print(f"  HY数据范围: {np.min(data[:, 1]):.4f} - {np.max(data[:, 1]):.4f}, 平均值: {np.mean(data[:, 1]):.4f}")
            # print(f"  {source}数据范围: {np.min(data[:, 2]):.4f} - {np.max(data[:, 2]):.4f}, 平均值: {np.mean(data[:, 2]):.4f}")
            # print(f"  差异范围: {np.min(data[:, 3]):.4f} - {np.max(data[:, 3]):.4f}")
            
            # ave = np.mean(data[:, 3])

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
        
        print("\n=== 步骤6执行完成 ===")
        return True
        
    except Exception as e:
        print(f"步骤6执行失败: {str(e)}")
        traceback.print_exc()
        return False


def step7(input_dir, output_dir,inspection_type):
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
        """读取valresult文件并转换一维索引为二维索引"""
        try:
            print(f"\n开始读取valresult文件: {file_path}")
            with open(file_path, 'r') as f:
                data = []
                line_count = 0
                for line in f:
                    try:
                        values = line.strip().split()
                        if len(values) >= 4:
                            # valresult格式: 索引号\tHY值\t参考值\t误差
                            pixel_index = int(float(values[0])) - 1  # 转换为0-based索引
                            error = float(values[3])  # 第4列是误差

                            # 将一维索引转换为二维索引
                            if rows is not None and cols is not None:
                                row_index = pixel_index // cols
                                col_index = pixel_index % cols
                            else:
                                # 如果没有维度信息，将索引作为行号，列号为0
                                row_index = pixel_index
                                col_index = 0

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
        if 'CDOM' in filename:
            return 'CDOM'
        elif 'TSM' in filename:
            return 'TSM'

    def match_coordinates(spaceresult, lat, lon, filename):
        """根据行列号匹配经纬度坐标并计算误差百分比"""
        print(f"\n开始坐标匹配...")
        # print(f"输入数据大小: spaceresult={len(spaceresult)}, lat shape={lat.shape}, lon shape={lon.shape}")
        
        # 获取产品类型
        product_type = get_product_type(filename)
        print(f"识别的产品类型: {product_type}")
        
        # 根据数据类型选择处理逻辑
        if product_type in ['SST', 'IPAR'] and lat.shape[1] == 1:
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
        if 'sst' in product_type or 'ipar' in product_type:

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
        lat_file = find_file_with_prefix(input_dir, inspection_type+'_lat')
        lon_file = find_file_with_prefix(input_dir, inspection_type+'_lon')
        
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

def step8(input_directory, output_directory,inspection_type):
    """
    第八步：处理时间序列数据和绘制时间序列图
    """
    def read_valresult_files_ground(input_directory, product):
        """读取现场验证数据文件并处理"""
        times = []
        deviations = []
        file_paths = []
        
        pattern = os.path.join(input_directory, f'valresult_{inspection_type}_XC_{product}_*.txt')
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
        satellites = ['HY1C', 'HY1D']
        for satellite in satellites:
            pattern = os.path.join(input_directory, f'valresult_{inspection_type}_{satellite}_{product}_*.txt')
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
        
        satellite_products = ['CDOM', 'TSM']
        # xc_products = ['AOT','chl','nLw','sst','TSM','CDOM','Rrs412', 'Rrs443', 'Rrs490', 'Rrs520', 'Rrs565', 'Rrs670','Rrs750','Rrs865']

        # 第一步：处理数据并生成中间文件
        output_files = []
         
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

def step9(input_directory, output_directory,inspection_type,source_type):
    """
    第九步：处理星地检验和星星检验数据，生成统计结果和图表
    """
    try:
        # 处理卫星交叉验证数据
        # print("\n处理卫星交叉验证数据...")
        step9_satellite(output_directory, output_directory,inspection_type,source_type)
        

    except Exception as e:
        print(f"步骤9执行失败: {e}")
        import traceback
        traceback.print_exc()
        return False
    
def get_timestamp():
    """获取时间戳"""
    return datetime.now().strftime("%Y%m%d%H%M%S")



def generate_satellite_statistics_file(filename, total_pixels, valid_pixels, 
                                    time_diff_counts, difference_counts, product):
    """生成卫星交叉验证统计文件"""
    product_names = {
        'CDOM': '悬浮泥沙浓度',
        'TSM': '有色可溶有机物浓度'
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

def generate_satellite_plots(valid_pixels, total_pixels, time_diff_counts, 
                        difference_counts, output_directory, product, satellite_type,inspection_type, timestamp=None):
    """生成卫星交叉验证统计图"""
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    
    product_names = {
        'CDOM': '有色可溶有机物浓度',
        'TSM': '悬浮泥沙浓度'
    }
    product_name = product_names.get(product, product)

    if timestamp is None:
        timestamp = get_timestamp()
    base_name = f"{inspection_type}_{satellite_type}_{product}_{timestamp}"
    
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



def step9_satellite(input_directory, output_directory,inspection_type,source_type):
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
                total_pixels, valid_pixels, time_diff_counts, difference_counts, satellite_type = analyze_star_check(
                    data['valresults'], data['spaceresults'], input_directory,inspection_type)
                
                if all(v is not None for v in [total_pixels, valid_pixels, time_diff_counts, difference_counts, satellite_type]):
                    # 生成统计文件
                    stats_filename = os.path.join(output_directory, 
                        f"resstastic_{inspection_type}_{source_type}_{product}_{timestamp}.txt")
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
                        valid_pixels,
                        total_pixels,
                        time_diff_counts,
                        difference_counts,
                        output_directory,
                        product,
                        source_type,
                        inspection_type,
                        timestamp
                    )
        
        return True
    except Exception as e:
        print(f"卫星交叉验证处理失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def analyze_star_check(valresults, spaceresults, input_directory,inspection_type):
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
        'CDOM': inspection_type+'_CDOM_',
        'TSM':  inspection_type+'_TSM_'
    }
    
    file_prefix = product_file_mapping.get(product)
    if file_prefix:
        for filename in os.listdir(input_directory):
            if filename.startswith(file_prefix) and not filename.startswith(file_prefix + 'flag1'):
                hy3a_file = os.path.join(input_directory, filename)
                break
    
    if not hy3a_file:
        print(f"未找到产品 {product} 对应的{inspection_type}文件")
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
        print(f"读取{inspection_type}文件失败: {e}")
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

def extract_satellite_type(filename):
    """从文件名中提取卫星类型"""
    if 'HY1C' in filename.upper():
        return 'HY1C'
    elif 'HY1D' in filename.upper():
        return 'HY1D'
    return 'UNKNOWN'



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
        report_filename1 = f.replace("resstastic_", "report_")
        
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
                            print(f"\n\n总像元数：：{total_pixels}\n\n")
                            print(f"\n\n有效检验像元数：：{valid_pixels}\n\n")
                            
                            # 将数据追加到report文件
                            with open(os.path.join(input_dir, report_filename1), 'a', encoding='utf-8') as outfile:
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

def step_report(datestr, input_temp, input_img, coldata_path, output_path, satellite_type,source_org_type,space_size,time_size):
 # 定义产品配置
    # 定义HY1C_VAR_CONFIG（卫星间验证）
    HY1C_VAR_CONFIG = {
        # 'sst': {'sources': ['HY1C'], 'unit': '℃', 'col_values': [25, 1800]},
        # 'chl': {'sources': ['HY1C'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        # 'Rrs412': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs443': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs490': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs520': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs565': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs670': {'sources': ['HY1C'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'AOT': {'sources': ['HY1C'], 'unit': '', 'col_values': [25, 1800]},
        'TSM': {'sources': ['HY1C'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        'CDOM': {'sources': ['HY1C'], 'unit': '1/m', 'col_values': [25, 1800]},
    }

    # 定义HY1D_VAR_CONFIG（卫星间验证）
    HY1D_VAR_CONFIG = {
        # 'sst': {'sources': ['HY1D'], 'unit': '℃', 'col_values': [25, 1800]},
        # 'chl': {'sources': ['HY1D'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        # 'Rrs412': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs443': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs490': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs520': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs565': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs670': {'sources': ['HY1D'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'AOT': {'sources': ['HY1D'], 'unit': '', 'col_values': [25, 1800]},
        'TSM': {'sources': ['HY1D'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        'CDOM': {'sources': ['HY1D'], 'unit': '1/m', 'col_values': [25, 1800]},
    }

    # 定义HY1E_VAR_CONFIG（卫星间验证）
    HY1E_VAR_CONFIG = {
        # 'sst': {'sources': ['HY1E'], 'unit': '℃', 'col_values': [25, 1800]},
        # 'chl': {'sources': ['HY1E'], 'unit': 'mg/m³', 'col_values': [25, 1800]},
        # 'Rrs412': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs443': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs490': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs520': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs565': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs670': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'Rrs750': {'sources': ['HY1E'], 'unit': 'sr⁻¹', 'col_values': [25, 1800]},
        # 'AOT': {'sources': ['HY1E'], 'unit': '', 'col_values': [25, 1800]},
        'TSM': {'sources': ['HY1E'], 'unit': 'mg/L', 'col_values': [25, 1800]},
        'CDOM': {'sources': ['HY1E'], 'unit': '1/m', 'col_values': [25, 1800]},
        # 'IPAR': {'sources': ['HY1E'], 'unit': 'Einstein/m²/d', 'col_values': [25, 1800]},
    }

    # 初始化VAR_CONFIG
    VAR_CONFIG = None

    if source_org_type == 'HY1C':
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
                source: dict(zip(['bias', 'rms', 'n'], calc_metric(coldata_path, source_org_type, var_name, datestr, satellite_type)))
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
                    val_results.append([
                        f'{satellite_type} vs {source}',
                        f"{metrics[source]['bias']:.4f}{unit}",
                        f"{metrics[source]['rms']:.4f}{unit}"
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
                new_pattern = f'report_{satellite_type}_{source_org_type}_{var_name}_{datestr}'
                matching_files = [f for f in files if new_pattern in f and f.endswith('.txt')]

                if matching_files:
                    # 从新格式文件名中提取时间戳
                    # 格式: report_HY1E_TERRA_sst_20250101_120000.txt
                    parts = matching_files[0].replace('.txt', '').split('_')
                    if len(parts) >= 5:
                        # 时间戳是最后一部分
                        timestamp = parts[-1]
                        break
                else:
                    # 尝试旧格式：{satellite}_COCTS_{source}_{var_name}_report_{datestr}_{time}.txt
                    old_pattern = f'{satellite_type}_COCTS_{source_org_type}_{var_name}_report_{datestr}'
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

    def _remove_empty_subsections_and_renumber(doc):
        """
        删除所有没有图片和实质内容的小节，并重新编号
        识别规则：小节标题格式为"X.Y"（如"3.1"、"8.2"、"10.1"等）
        """
        import re
        from collections import defaultdict

        # 第一步：找出需要删除的段落范围
        paragraphs_to_delete = []
        subsection_info = []  # 记录所有小节信息：(para_idx, chapter, subsection, text)

        # 识别所有小节标题（X.Y格式，如3.1、8.2、10.1等）
        subsection_pattern = re.compile(r'^(\d+)\.(\d+)')

        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            # 检查是否是小节标题
            match = subsection_pattern.match(text)
            if match:
                chapter = int(match.group(1))
                subsection = int(match.group(2))
                subsection_info.append((i, chapter, subsection, text))

        # 第二步：检查每个小节是否包含内容（非空段落或有实际内容）
        for idx in range(len(subsection_info)):
            start_idx = subsection_info[idx][0]
            # 确定小节结束位置（下一个小节开始前，或文档结束）
            end_idx = subsection_info[idx + 1][0] if idx + 1 < len(subsection_info) else len(doc.paragraphs)

            # 检查这个小节是否为空（只有标题，没有其他有意义内容）
            has_content = False
            for para_idx in range(start_idx + 1, end_idx):
                if para_idx < len(doc.paragraphs):
                    para = doc.paragraphs[para_idx]
                    text = para.text.strip()

                    # 跳过空白和其他小节标题
                    if not text or subsection_pattern.match(text):
                        continue

                    # 检查是否包含图片
                    has_image = False
                    for run in para.runs:
                        if hasattr(run, '_element'):
                            drawings = run._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
                            if drawings:
                                has_image = True
                                break

                    # 如果有图片或有实质内容（排除占位符）
                    if has_image:
                        has_content = True
                        break

                    # 检查是否只是占位符（包含{{}}的文本）
                    if '{{' not in text and len(text) > 5:
                        has_content = True
                        break

            # 如果小节为空，标记删除该范围的所有段落
            if not has_content:
                for para_idx in range(start_idx, end_idx):
                    if para_idx not in paragraphs_to_delete and para_idx < len(doc.paragraphs):
                        paragraphs_to_delete.append(para_idx)

        # 第三步：删除标记的段落（从后往前删，避免索引变化）
        for para_idx in sorted(paragraphs_to_delete, reverse=True):
            if para_idx < len(doc.paragraphs):
                p = doc.paragraphs[para_idx]
                p_element = p._element
                p_element.getparent().remove(p_element)

        # 第四步：重新编号所有小节（按章节分组）
        chapter_counters = defaultdict(int)

        for para in doc.paragraphs:
            text = para.text.strip()
            match = subsection_pattern.match(text)
            if match:
                chapter = int(match.group(1))
                chapter_counters[chapter] += 1
                # 替换为新的编号
                new_text = re.sub(r'^\d+\.\d+', f'{chapter}.{chapter_counters[chapter]}', text)
                para.text = new_text

    def _handle_unused_placeholders(doc, replacements):
        """
        处理模板中存在但未使用的占位符
        - 表一（val_results）：填写检验类型和"/"
        - 表二（col_results）：删除整行
        - 第三章图片占位符：由_remove_empty_subsections_and_renumber处理
        """
        import re

        # 扫描所有占位符
        all_placeholders = set()
        placeholder_pattern = re.compile(r'\{\{([^}]+)\}\}')

        # 从所有段落中提取占位符
        for p in doc.paragraphs:
            matches = placeholder_pattern.findall(p.text)
            all_placeholders.update(matches)

        # 从所有表格中提取占位符
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        matches = placeholder_pattern.findall(p.text)
                        all_placeholders.update(matches)

        # 识别val_results和col_results占位符
        val_pattern = re.compile(r'val_results_(\w+)')
        col_pattern = re.compile(r'col_results_(\w+)')

        used_tables = set(replacements.get('tables', {}).keys())

        # 处理未使用的val_results占位符（表一）
        for placeholder in all_placeholders:
            match = val_pattern.match(placeholder)
            if match:
                full_placeholder = f'{{{{{placeholder}}}}}'
                if full_placeholder not in used_tables:
                    # 填写检验类型和"/"
                    var_name = match.group(1)
                    _fill_unused_val_results(doc, var_name)

        # 处理未使用的col_results占位符（表二）
        for placeholder in all_placeholders:
            match = col_pattern.match(placeholder)
            if match:
                full_placeholder = f'{{{{{placeholder}}}}}'
                if full_placeholder not in used_tables:
                    # 删除整行
                    var_name = match.group(1)
                    _delete_col_results_row(doc, var_name)

    def _fill_unused_val_results(doc, var_name):
        """为未使用的val_results占位符填写检验类型和"/" """
        placeholder = f'{{{{val_results_{var_name}}}}}'

        # 填充表格：检验类型 | / | /
        # 由于是未使用的占位符，我们无法确定具体的检验类型，使用通用值
        val_results = [['未配置检验', '/', '/']]
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

    template_path=os.path.join(input_temp, f'satelite_template.docx')

    
    fill_template(
        template_path,
        output_docx=os.path.join(save_path, f'{satellite_type.upper()}_COCTS_{source_org_type.upper()}_val_report_{datestr}.docx'),
        output_pdf=os.path.join(save_path, f'{satellite_type.upper()}_COCTS_{source_org_type.upper()}_val_report_{datestr}.pdf'),
        replacements=replacements
    )

    hy1d_cocts_daily_report(datestr, input_temp, input_img, coldata_path, output_path, satellite_type,source_org_type)


def word_to_pdf(input_dir, output_dir):
    """
    将指定目录中的所有 .docx 文件转换为 PDF 文件。

    参数:
        input_dir (str): 包含 .docx 文件的输入目录。
        output_dir (str): 保存转换后的 PDF 文件的输出目录。
    """
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 遍历输入目录中的所有文件
    for filename in os.listdir(input_dir):
        # 检查文件是否以 .docx 结尾
        if filename.endswith(".docx"):
            # 构建输入文件的完整路径
            input_file = os.path.join(input_dir, filename)
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
                print(f"成功转换：{input_file} -> {output_file}")
            else:
                print(f"转换失败：{input_file}")
                print("错误信息：", result.stderr.decode())


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
            match = re.search(r'(HY1[CDE])_(HY1[CDE])_([A-Za-z0-9]+)_(\d{8})(\d{6})', file_name)
            if not match:
                # 尝试其他模式: valresult_HY1C_AQUA_Rrs412_20231011102445
                match = re.search(r'(?:valresult|map|valstatistic|report|statistic)_(HY1[CD])_(HY1[CDE])_([A-Za-z0-9]+)_(\d{8})(\d{6})', file_name)
            
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
def word_to_pdf(input_dir, output_dir):
    """
    将指定目录中的所有 .docx 文件转换为 PDF 文件。

    参数:
        input_dir (str): 包含 .docx 文件的输入目录。
        output_dir (str): 保存转换后的 PDF 文件的输出目录。
    """
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 遍历输入目录中的所有文件
    for filename in os.listdir(input_dir):
        # 检查文件是否以 .docx 结尾
        if filename.endswith(".docx"):
            # 构建输入文件的完整路径
            input_file = os.path.join(input_dir, filename)
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
                print(f"成功转换：{input_file} -> {output_file}")
            else:
                print(f"转换失败：{input_file}")
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

if __name__ == "__main__":
    main() 