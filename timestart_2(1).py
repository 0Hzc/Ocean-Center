import os
import re
from datetime import datetime, timedelta
from collections import defaultdict
import configparser
import subprocess
import sys
import json
from task_queue import TaskQueue

# 从命令行参数获取 start_time_str 和 end_time_str
if len(sys.argv) != 3:
    print("Usage: python xx.py yyyymmdd yyyymmdd")
    sys.exit(1)  # 如果参数数量不正确，退出程序\

start_time_str = sys.argv[1]  # 第一个参数作为 start_time_str
end_time_str = sys.argv[2]    # 第二个参数作为 end_time_str

# 创建task文件和独立输出目录
def create_task_file(start_time_str, end_time_str):
    """在程序启动时立即创建task文件"""
    task_id = f"task_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    task_dir = os.path.join("tasks", task_id)
    os.makedirs(task_dir, exist_ok=True)

    task_info = {
        'task_id': task_id,
        'start_time': start_time_str,
        'end_time': end_time_str,
        'created_at': datetime.now().isoformat(),
        'status': 'running',
        'output_dir': task_dir
    }

    task_file = os.path.join(task_dir, 'task_info.json')
    with open(task_file, 'w', encoding='utf-8') as f:
        json.dump(task_info, f, indent=2, ensure_ascii=False)

    print(f"\n[Task] 任务已创建:")
    print(f"  任务ID: {task_id}")
    print(f"  输出目录: {task_dir}")
    print(f"  任务文件: {task_file}")
    print(f"  时间范围: {start_time_str} -> {end_time_str}\n")

    return task_dir, task_file

# 更新task文件状态
def update_task_status(task_file, status, warnings=None, error=None):
    """更新task文件的状态信息"""
    try:
        with open(task_file, 'r', encoding='utf-8') as f:
            task_info = json.load(f)

        task_info['status'] = status
        task_info['completed_at'] = datetime.now().isoformat()

        if warnings:
            task_info['warnings'] = warnings
            task_info['warning_count'] = len(warnings)

        if error:
            task_info['error'] = str(error)

        with open(task_file, 'w', encoding='utf-8') as f:
            json.dump(task_info, f, indent=2, ensure_ascii=False)

        print(f"\n[Task] 任务状态已更新: {status}")
    except Exception as e:
        print(f"\n[Task] 更新任务状态失败: {e}")

# 立即创建task文件
task_output_dir, task_info_file = create_task_file(start_time_str, end_time_str)

# 初始化 run_warning 字段
run_warning = []

# 读取配置文件
config_file = "config.ini"
config = configparser.ConfigParser()
config.read(config_file)

# 使用 get 方法并提供默认值
source_type = config.get('VALIDATION', 'source_type', fallback="NONE")
satelite_type = config.get('SATELLITE', 'type', fallback="NONE")
# 定义文件路径
# input_path = "input"
# 定义文件路径
try:
    input_path = config.get('PATH', 'input_dir', fallback="input")
except configparser.NoSectionError:
    run_warning.append(f"配置文件 {config_file} 中缺少 [PATH] 部分，使用默认路径 'input'")
    input_path = "input"
input_path = input_path.replace('\\', '/')#兼容linux
sat_path = os.path.join(input_path, "01_sat")
reference_path = os.path.join(input_path, "02_reference")

# 定义列表
HY_L2A_NAME = []
HY_L2B_NAME = []
HY_L2C_NAME = []
HY_L2T_NAME = []
OC_NAME = []
SST_NAME = []
XC_NAME = []
SAME_TIME_FILES = defaultdict(list)  # 使用defaultdict来存储相同时间的文件

# 定义时间格式
time_format = "%Y%m%d"
xc_time_format = "%Y%m%d%H%M%S"
aot_time_format = "%Y%m%d%H%M"  # 新增AOT文件的时间格式

# 定义时间范围的起始和结束时间
try:
    start_time = datetime.strptime(start_time_str, time_format)
    end_time = datetime.strptime(end_time_str, time_format)
    end_time += timedelta(days=1)  # 将截止日期加一天
except ValueError as e:
    run_warning.append(f"时间格式错误：{e}")
    start_time = None
    end_time = None

# 将截止日期加一天
end_time += timedelta(days=1)

# 定义正则表达式匹配时间信息
time_pattern = re.compile(r"(\d{8}T\d{6})")
xc_time_pattern = re.compile(r"(\d{14})")
aot_time_pattern = re.compile(r"(\d{10})")


# 扫描文件并筛选符合时间段的文件
def scan_files(directory, pattern, xc_pattern, aot_pattern, list_name_l2a=None, list_name_l2b=None, list_name_l2c=None, list_name_l2t=None, list_name_oc=None, list_name_sst=None, list_name_xc=None, same_time_files=None):
    

    # 如果某些参数是 None，则初始化为空列表
    list_name_l2a = list_name_l2a if list_name_l2a is not None else []
    list_name_l2b = list_name_l2b if list_name_l2b is not None else []
    list_name_l2c = list_name_l2c if list_name_l2c is not None else []
    list_name_l2t = list_name_l2t if list_name_l2t is not None else []
    list_name_oc = list_name_oc if list_name_oc is not None else []
    list_name_sst = list_name_sst if list_name_sst is not None else []
    list_name_xc = list_name_xc if list_name_xc is not None else []
    same_time_files = same_time_files if same_time_files is not None else defaultdict(list)
    


    # print(f"Scanning file: {time_pattern}\n")
    # print(f"Scanning file: {xc_pattern}\n")
    # print(f"Scanning file: {aot_pattern}\n")

    
    for root, dirs, files in os.walk(directory):
        for file in files:
            match = pattern.search(file)
            if match:
                file_time_str = match.group(1)
                try:
                    file_time = datetime.strptime(file_time_str, "%Y%m%dT%H%M%S")
                except ValueError as e:
                    run_warning.append(f"Error parsing time for file {file}: {e}")
                    continue
            else:
                # 如果普通文件时间格式不匹配，尝试匹配xc文件的时间格式
                match = xc_pattern.search(file)
                if match:
                    file_time_str = match.group(1)
                    try:
                        file_time = datetime.strptime(file_time_str, xc_time_format)
                    except ValueError as e:
                        run_warning.append(f"Error parsing time for file {file}: {e}")
                        continue
                else:
                    # 尝试匹配AOT文件的时间格式
                    match = aot_pattern.search(file)
                    if match:
                        file_time_str = match.group(1)
                        try:
                            file_time = datetime.strptime(file_time_str, aot_time_format)
                        except ValueError as e:
                            run_warning.append(f"Error parsing time for file {file}: {e}")
                            continue
                    else:
                        # 如果两种时间格式都不匹配，跳过该文件
                        continue
            if start_time <= file_time <= end_time:
                if any(keyword in file for keyword in ["AOPRes", "WQP", "AOT", "CTD"]):
                    list_name_xc.append(file)
                elif "L2A" in file:
                    list_name_l2a.append(file)
                elif "L2B" in file:
                    list_name_l2b.append(file)
                elif "L2C" in file:
                    list_name_l2c.append(file)
                elif "L2T" in file:
                    list_name_l2t.append(file)
                elif "OC" in file:
                    list_name_oc.append(file)
                elif "SST" in file:
                    list_name_sst.append(file)
                # 将文件及其时间添加到相同时间文件列表
                file_date = file_time.date()  # 只保留年月日
                same_time_files[file_date].append(file)


# 扫描 01_sat 目录下的文件
try:
    scan_files(sat_path, time_pattern, xc_time_pattern, aot_time_pattern, HY_L2A_NAME, HY_L2B_NAME, HY_L2C_NAME,HY_L2T_NAME,None, None, XC_NAME, SAME_TIME_FILES)
except Exception as e:
    run_warning.append(f"Error scanning files in {sat_path}: {e}")

# 扫描 02_reference 目录下的文件
try:
    scan_files(reference_path, time_pattern, xc_time_pattern, aot_time_pattern, None, None, None, None, OC_NAME, SST_NAME, XC_NAME, SAME_TIME_FILES)
except Exception as e:
    run_warning.append(f"Error scanning files in {reference_path}: {e}")

# # 打印结果
# print("HY_L2A_NAME:", HY_L2A_NAME)
# print("HY_L2B_NAME:", HY_L2B_NAME)
# print("HY_L2B_NAME:", HY_L2C_NAME)
# print("HY_L2B_NAME:", HY_L2T_NAME)
# print("OC_NAME:", OC_NAME)
# print("SST_NAME:", SST_NAME)
# print("XC_NAME:", XC_NAME)

# 打印相同时间的文件
print("Files with the same time:")
for file_time, files in SAME_TIME_FILES.items():
    if len(files) > 1:  # 只打印具有多个文件的组
        print(f"Time: {file_time}")
        for file in files:
            print(f"  {file}")


# 处理每组相同时间的文件
def process_files(file_group):
    print(f"Processing files for time: {file_group[0]}")
    for file in file_group[1]:
        print(f"  Processing file: {file}")
        # 定义分类字典
        categories = {
            "HY1C": {"l2a_file": None, "l2b_file": None},
            "HY1D": {"l2a_file": None, "l2b_file": None},
            "HY1E": {"l2a_file": None, "l2b_file": None,"l2c_file": None,"l2t_file": None},
            "TERRA": {"oc_file": None, "sst_file": None},
            "AQUA": {"oc_file": None, "sst_file": None},
            "SNPP": {"oc_file": None, "sst_file": None},
            "JPSS": {"oc_file": None, "sst_file": None},
            "XC": {"aopres_file": None, "wqp_file": None, "aot_file": None, "ctd_file": None}
        }

        # 定义分类规则
        category_rules = {
            "HY1C": {"prefix": "H1C", "l2a": "L2A", "l2b": "L2B"},
            "HY1D": {"prefix": "H1D", "l2a": "L2A", "l2b": "L2B"},
            "HY1E": {"prefix": "H1E", "l2a": "L2A", "l2b": "L2B", "l2c": "L2C", "l2t": "L2T"},
            "TERRA": {"prefix": "TERRA", "oc": "OC", "sst": "SST"},
            "AQUA": {"prefix": "AQUA", "oc": "OC", "sst": "SST"},
            "SNPP": {"prefix": "SNPP", "oc": "OC", "sst": "SST"},
            "JPSS": {"prefix": "JPSS", "oc": "OC", "sst": "SST"},
            "XC": {"prefix": "XC", "aopres": "AOPRes", "wqp": "WQP", "aot": "AOT", "ctd": "CTD"}
        }

        # # 读取现有的 config.ini 文件
        # config_file = "config.ini"
        # config = configparser.ConfigParser()

        try:
            if os.path.exists(config_file):
                # 显式指定文件编码为 UTF-8
                with open(config_file, 'r', encoding='utf-8') as fp:
                    config.read_file(fp)
            else:
                run_warning.append(f"文件 {config_file} 不存在，将创建新文件。")
        
        except UnicodeDecodeError as e:
            run_warning.append(f"读取文件时发生编码错误：{e}")
            print("尝试以默认编码重新读取文件...")
            config.read(config_file)
        except Exception as e:
            run_warning.append(f"Error reading configuration file {config_file}: {e}")
        
        # 如果文件存在，更新 categories 字典
        for category in categories:
            if category in config:
                for key in categories[category]:
                    if key in config[category]:
                        categories[category][key] = config[category][key]

        # 分类逻辑
        for file in file_group[1]:
            for category, rules in category_rules.items():
                if category == "XC":
                    if rules["aopres"] in file:
                        categories[category]["aopres_file"] = file
                    elif rules["wqp"] in file:
                        categories[category]["wqp_file"] = file
                    elif rules["aot"] in file:
                        categories[category]["aot_file"] = file
                    elif rules["ctd"] in file:
                        categories[category]["ctd_file"] = file
                else:
                    if rules["prefix"] in file:
                        if "l2a" in rules and rules["l2a"] in file:
                            categories[category]["l2a_file"] = file
                        elif "l2b" in rules and rules["l2b"] in file:
                            categories[category]["l2b_file"] = file
                        elif "l2c" in rules and rules["l2c"] in file:
                            categories[category]["l2c_file"] = file
                        elif "l2t" in rules and rules["l2t"] in file:
                            categories[category]["l2t_file"] = file
                        elif "oc" in rules and rules["oc"] in file:
                            categories[category]["oc_file"] = file
                        elif "sst" in rules and rules["sst"] in file:
                            categories[category]["sst_file"] = file

        # 将 None 转换为字符串
        for category, files in categories.items():
            for key, value in files.items():
                if value is None:
                    files[key] = ""

        # 写入.ini文件
        for category, files in categories.items():
            config[category] = files

        try:
            with open(config_file, "w", encoding='utf-8') as configfile:
                config.write(configfile)
            print(f"分类完成，结果已写入 {config_file} 文件")
        except Exception as e:
            run_warning.append(f"Error writing to configuration file {config_file}: {e}")


# 使用任务队列防止并发执行
task_queue = TaskQueue()
print(f"\n[TaskQueue] 尝试获取任务执行权限...")

if task_queue.acquire(timeout=600, wait=True):
    try:
        for file_time, files in SAME_TIME_FILES.items():
            if len(files) > 1:  # 只处理具有多个文件的组
                process_files((file_time, files))
                    # 调用 H1CD.py 脚本
                # 在合适的位置读取 config.ini 文件中的 source_type 值
                if satelite_type in ["HY1C", "HY1D"]:
                    if source_type in ['HY1C', 'HY1D','HY1E']:
                        try:
                            subprocess.run(["python", "satelite.py"], check=True)
                            print("satelite.py 脚本执行完成")
                        except subprocess.CalledProcessError as e:
                            run_warning.append(f"执行 satelite.py 脚本时出错：{e}")
                    else:
                        try:
                            subprocess.run(["python", "H1CD.py"], check=True)
                            print("H1CD.py 脚本执行完成")
                        except subprocess.CalledProcessError as e:
                            run_warning.append(f"执行 H1CD.py 脚本时出错：{e}")
                elif satelite_type in ["HY1E"]:
                    if source_type in ['HY1C', 'HY1D','HY1E']:
                        try:
                            subprocess.run(["python", "satelite.py"], check=True)
                            print("satelite.py 脚本执行完成")
                        except subprocess.CalledProcessError as e:
                            run_warning.append(f"执行 satelite.py 脚本时出错：{e}")
                    else:
                        try:
                            subprocess.run(["python", "docker_setup.py"], check=True)
                            print("docker_setup.py 脚本执行完成")
                        except subprocess.CalledProcessError as e:
                            run_warning.append(f"执行 docker_setup.py 脚本时出错：{e}")
    finally:
        # 确保无论成功还是失败都释放锁
        task_queue.release()

    if run_warning:
        print("\n运行过程中出现以下警告：")
        for warning in run_warning:
            print(f"run_warning:{warning}")
        # 更新任务状态为已完成(带警告)
        update_task_status(task_info_file, 'completed_with_warnings', warnings=run_warning)
    else:
        print("\n运行完成，未发现异常。")
        # 更新任务状态为成功完成
        update_task_status(task_info_file, 'completed')
else:
    print("\n[TaskQueue] 无法获取任务执行权限，退出。")
    # 更新任务状态为失败
    update_task_status(task_info_file, 'failed', error='无法获取任务执行权限')
    sys.exit(1)
