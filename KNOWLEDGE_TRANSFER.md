# HY1C卫星验证系统 - 完整知识传授文档

> 本文档用于帮助新的Claude快速理解整个HY1C卫星验证系统的架构、数据流程和代码结构

---

## 📋 目录

1. [系统概述](#1-系统概述)
2. [目录结构](#2-目录结构)
3. [数据文件格式](#3-数据文件格式)
4. [核心流程](#4-核心流程)
5. [文件命名规则](#5-文件命名规则)
6. [关键代码模块](#6-关键代码模块)
7. [月季年报告生成](#7-月季年报告生成)
8. [常见问题](#8-常见问题)
9. [调试指南](#9-调试指南)

---

## 1. 系统概述

### 1.1 项目背景

HY1C是中国海洋一号C卫星，该系统用于验证HY1C卫星观测数据的准确性。验证方法：
- **交叉验证**: 与其他卫星数据对比（AQUA, TERRA, SNPP, JPSS）
- **现场验证**: 与现场观测数据对比（XC）

### 1.2 验证产品

主要验证以下海洋参数：

| 产品 | 全称 | 单位 | 验证源 |
|------|------|------|--------|
| **sst** | 海表温度 | ℃ | AQUA/TERRA/XC |
| **chl** | 叶绿素浓度 | mg/m³ | AQUA/TERRA/XC |
| **AOT** | 气溶胶光学厚度 | - | AQUA/TERRA/XC |
| **Rrs412** | 412nm遥感反射率 | sr⁻¹ | AQUA/TERRA |
| **Rrs443** | 443nm遥感反射率 | sr⁻¹ | AQUA/TERRA |
| **Rrs490** | 490nm遥感反射率 | sr⁻¹ | AQUA/TERRA |
| **Rrs520** | 520nm遥感反射率 | sr⁻¹ | AQUA/TERRA |
| **Rrs565** | 565nm遥感反射率 | sr⁻¹ | AQUA/TERRA |
| **Rrs670** | 670nm遥感反射率 | sr⁻¹ | AQUA/TERRA |

### 1.3 验证指标

- **bias**: 平均偏差
- **RMS**: 均方根误差
- **n**: 有效匹配点数

---

## 2. 目录结构

```
HY/
├── config.ini                          # 配置文件
├── H1CD.py                             # 核心验证流程
├── generate_report.py                  # 日报告生成（已废弃）
├── generate_monthly_report_v5.py       # 月季年报告生成（最新）
│
├── input/                              # 输入数据
│   ├── 01_sat/                         # 卫星数据
│   │   ├── HY1C_*.txt                  # HY1C原始数据
│   │   ├── HY1C_lat_*.txt              # HY1C纬度数据（很少用）
│   │   └── HY1C_lon_*.txt              # HY1C经度数据（很少用）
│   │
│   ├── 02_reference/                   # 参考数据
│   │   ├── 02_reference_preprocess/    # 预处理后的参考数据
│   │   │   ├── AQUA_*.txt              # AQUA卫星数据
│   │   │   ├── TERRA_*.txt             # TERRA卫星数据
│   │   │   ├── AQUA_Lat_*.txt          # AQUA纬度数据 ⭐
│   │   │   ├── AQUA_Lon_*.txt          # AQUA经度数据 ⭐
│   │   │   ├── TERRA_Lat_*.txt         # TERRA纬度数据 ⭐
│   │   │   └── TERRA_Lon_*.txt         # TERRA经度数据 ⭐
│   │   └── xc/                         # 现场观测数据
│   │       └── *.txt
│   │
│   └── 05_reports/                     # 报告模板
│       ├── new_auqa_terra_template.docx
│       ├── new_snpp_jpss_template.docx
│       └── new_xc_template.docx
│
└── output/                             # 输出数据
    ├── 03_collocation/                 # 时空匹配结果
    │   ├── timeresult_*.txt            # 时间匹配结果
    │   ├── spaceresult_*.txt           # 空间匹配结果（仅XC）
    │   ├── valresult_*.txt             # 验证结果（临时，会被重命名）
    │   └── *_matchup_*.txt             # 重命名后的valresult ⭐
    │
    ├── 04_visualization/               # 可视化数据和图片
    │   ├── map_*.txt                   # 地理分布数据 ⭐⭐⭐
    │   ├── *_PIE_*.jpg                 # 饼图
    │   ├── *_GEO_*.jpg                 # 地理分布图
    │   └── timeseries_*.txt            # 时序数据
    │
    ├── 05_reports/                     # 日报告
    │   ├── report_*.txt                # 统计数据
    │   └── *_val_report_*.pdf          # PDF报告
    │
    └── 06_summary_reports/             # 汇总报告
        ├── monthly/                    # 月报告
        ├── quarterly/                  # 季度报告
        └── yearly/                     # 年报告
```

---

## 3. 数据文件格式

### 3.1 matchup文件格式

**文件名**: `HY1C_COCTS_AQUA_sst_matchup_20251009_104500.txt`

**说明**: 重命名后的valresult文件，包含验证结果

**格式**:
```
/begin header
/HY satellite=HY1C
/Validation source=AQUA
/product=sst
/HY time=20251009104500
/On-site time=20251009103000
/HY file=HY1C_sst_20251009104500.txt
/On-site file=AQUA_sst_20251009103000.txt
/Time difference=0.25h
/Effective pixel count=12345
/Total pixel count=50000
/validation result=2.34%
/fields=number	sst_HY	sst_AQUA	difference
/unites=NA	℃	℃	%
/end header
1	25.3456	25.1234	0.87
2	26.1234	26.0123	0.43
3	24.5678	24.7890	-0.91
...
```

**字段说明**:
- Column 0: 行号（1-based index）
- Column 1: HY1C观测值
- Column 2: 参考观测值（AQUA/TERRA/XC）
- Column 3: 误差百分比

### 3.2 map文件格式 ⭐⭐⭐

**文件名**: `map_HY1C_AQUA_sst_20251009104500.txt`

**位置**: `output/04_visualization/`

**说明**: 用于生成地理分布图的数据文件，包含坐标和误差

**格式**:
```
36.247314	125.436188	35.93
36.23671	125.470238	32.03
36.245987	125.47541	42.35
36.235397	125.509666	46.15
...
```

**字段说明**:
- Column 0: 纬度（latitude）
- Column 1: 经度（longitude）
- Column 2: 误差百分比（error%）

**关键特性**:
- ✅ 已经包含坐标信息，无需再次匹配
- ✅ 可直接用于绘图
- ✅ 每天的验证会生成一个map文件
- ✅ 月报告需要聚合多天的map文件

### 3.3 lat/lon文件格式

**文件名**: 
- `TERRA_Lat_20251009101001.txt`
- `TERRA_Lon_20251009101001.txt`

**位置**: `input/02_reference/02_reference_preprocess/`

**格式**:
```
# Tab-separated values
12.345	12.346	12.347	12.348	...
12.456	12.457	12.458	12.459	...
12.567	12.568	12.569	12.570	...
...
```

**说明**:
- 每行代表一个扫描行
- 每列代表一个像素
- SST等产品: 1维数据（只有一列或每行一个值）
- Rrs等产品: 2维数据（多列）

### 3.4 report文件格式

**文件名**: `HY1C_COCTS_AQUA_sst_report_20251009_104500.txt`

**位置**: `output/05_reports/`

**格式**:
```
/begin header
/HY satellite=HY1C
/Validation source=AQUA
/product=sst
/bias=-0.27
/RMS=1.49
/Effective pixel count=808924
/validation result=1.49
/end header
```

**用途**: 存储统计结果，用于月季年报告聚合

---

## 4. 核心流程

### 4.1 H1CD.py 主流程

```python
def main():
    """
    完整的验证流程
    """
    # 步骤1: 文件预处理
    preprocess_files(input_dir, output_dir)
    
    # 步骤2: 时间匹配
    time_match(satellite_files, reference_files, time_threshold)
    
    # 步骤3: 空间匹配
    space_match(timeresult_files, window_size)
    
    # 步骤4: 数据提取
    extract_data(spaceresult_files)
    
    # 步骤5: 现场数据匹配（仅XC）
    if source_type == 'XC':
        match_field_data(...)
    
    # 步骤6: 生成验证结果
    if source_type == 'XC':
        xc_validation(output_dir, output_dir)
    else:
        satellite_validation(output_dir, output_dir)
    
    # 步骤7: 生成误差地图（非XC）
    if source_type != 'XC':
        step7(satellite_type, output_dir, output_dir)
    
    # 步骤8: 生成折线图
    step8(satellite_type, output_dir, output_dir)
    
    # 步骤9: 生成统计结果和图表
    step9(satellite_type, output_dir, output_dir)
    
    # 步骤10: 生成报告
    if source_type == 'XC':
        make_ground_report_data(output_dir)
    else:
        make_sat_report_data(output_dir)
    
    # 步骤11: 文件整理和重命名
    organize_files(output_dir, final_output_dir)
    rename_output_files(final_output_dir)
```

### 4.2 步骤6: satellite_validation()

**功能**: 生成验证结果文件

**输入**: timeresult文件或spaceresult文件
**输出**: valresult文件（后续被重命名为matchup文件）

**关键代码**:
```python
def satellite_validation(input_path, output_path):
    # 读取HY和参考数据
    X = read_data(HY_file)
    Y = read_data(reference_file)
    
    # 计算统计指标
    bias = np.mean(X - Y)
    RMS = np.sqrt(np.mean((X - Y) ** 2))
    
    # 写入valresult文件
    val_path = os.path.join(output_path, f'valresult_{HY}_{source}_{product}_{timeHY}.txt')
    with open(val_path, 'w') as f:
        f.write('/begin header\n')
        # ... header信息
        f.write('/end header\n')
        
        # 写入数据: 行号\tHY值\t参考值\t误差%
        for i, (x, y) in enumerate(zip(X, Y), 1):
            diff = ((x - y) / y) * 100 if y != 0 else 0
            f.write(f'{i}\t{x}\t{y}\t{diff}\n')
```

### 4.3 步骤7: step7() - 生成地理分布图

**功能**: 从valresult文件生成map文件和地理分布图

**关键流程**:
```python
def step7(satellite_type, input_dir, output_dir):
    # 1. 读取valresult文件
    valresult_files = glob.glob(os.path.join(input_dir, 'valresult*.txt'))
    
    for valresult_file in valresult_files:
        # 2. 解析valresult数据
        valresult_data = read_valresult(valresult_file)
        # 格式: [(row_index, error), ...]
        
        # 3. 查找lat/lon文件
        lat_file = find_file(input_dir, f'{satellite_type}_lat')
        lon_file = find_file(input_dir, f'{satellite_type}_lon')
        
        # 4. 读取lat/lon数据
        lat = read_lat_lon(lat_file)
        lon = read_lat_lon(lon_file)
        
        # 5. 匹配坐标
        matched_data = match_coordinates(valresult_data, lat, lon)
        # 格式: [(latitude, longitude, error), ...]
        
        # 6. 生成map文件
        map_file = os.path.join(output_dir, f'map_{satellite}_{source}_{product}_{date}{time}.txt')
        with open(map_file, 'w') as f:
            for lat, lon, err in matched_data:
                f.write(f'{lat}\t{lon}\t{err}\n')
        
        # 7. 生成地理分布图
        plot_error_map(matched_data, output_file)
```

**match_coordinates逻辑**:
```python
def match_coordinates(valresult_data, lat, lon, product):
    matched_data = []
    
    if lat.ndim == 1 or (lat.ndim == 2 and lat.shape[1] == 1):
        # 一维数据（SST等）
        for row_idx, error in valresult_data:
            if 0 <= row_idx < len(lat):
                matched_lat = lat[row_idx]
                matched_lon = lon[row_idx]
                matched_data.append([matched_lat, matched_lon, error])
    else:
        # 二维数据（Rrs等）
        col_idx = lat.shape[1] // 2  # 使用中间列
        for row_idx, error in valresult_data:
            if 0 <= row_idx < lat.shape[0]:
                matched_lat = lat[row_idx, col_idx]
                matched_lon = lon[row_idx, col_idx]
                matched_data.append([matched_lat, matched_lon, error])
    
    return matched_data
```

### 4.4 文件重命名流程

**rename_output_files()** 中的关键规则:

```python
# valresult文件 → matchup文件
if file_name.startswith("valresult_"):
    new_name = f"{satellite}_COCTS_{source}_{product}_matchup_{date}_{time}.txt"

# map_*.jpg文件 → GEO图
elif file_name.startswith("map_"):
    if file_name.endswith(".jpg"):
        new_name = f"{satellite}_COCTS_{source}_{product}_GEO_{date}_{time}.jpg"
    # 注意: map_*.txt文件不会被重命名！

# valstastic_*.jpg → PIE图
elif file_name.startswith("valstastic_"):
    new_name = f"{satellite}_COCTS_{source}_{product}_PIE_{date}_{time}.jpg"

# report_*.txt → report文件
elif file_name.startswith("report_"):
    new_name = f"{satellite}_COCTS_{source}_{product}_report_{date}_{time}.txt"
```

**关键发现**: 
- ✅ map_*.txt文件**不会被重命名**，保持原名
- ✅ map_*.jpg文件会被重命名为*_GEO_*.jpg
- ✅ valresult_*.txt会被重命名为*_matchup_*.txt

---

## 5. 文件命名规则

### 5.1 原始命名（H1CD.py生成）

```
valresult_{SATELLITE}_{SOURCE}_{PRODUCT}_{YYYYMMDD}_{HHMMSS}.txt
map_{SATELLITE}_{SOURCE}_{PRODUCT}_{YYYYMMDD}{HHMMSS}.txt
map_{SATELLITE}_{SOURCE}_{PRODUCT}_{YYYYMMDD}{HHMMSS}.jpg
valstastic_{SATELLITE}_{SOURCE}_{PRODUCT}_{YYYYMMDD}{HHMMSS}.jpg
report_{SATELLITE}_{SOURCE}_{PRODUCT}_{YYYYMMDD}_{HHMMSS}.txt
```

### 5.2 重命名后（organize_files后）

```
{SATELLITE}_COCTS_{SOURCE}_{PRODUCT}_matchup_{YYYYMMDD}_{HHMMSS}.txt
map_{SATELLITE}_{SOURCE}_{PRODUCT}_{YYYYMMDD}{HHMMSS}.txt  # 不变！
{SATELLITE}_COCTS_{SOURCE}_{PRODUCT}_GEO_{YYYYMMDD}_{HHMMSS}.jpg
{SATELLITE}_COCTS_{SOURCE}_{PRODUCT}_PIE_{YYYYMMDD}_{HHMMSS}.jpg
{SATELLITE}_COCTS_{SOURCE}_{PRODUCT}_report_{YYYYMMDD}_{HHMMSS}.txt
{SATELLITE}_COCTS_{SOURCE}_val_report_{YYYYMMDD}_{HHMMSS}.pdf
```

### 5.3 模式匹配示例

```python
# 查找matchup文件
pattern = f'{satellite}_COCTS_{source}_{product}_matchup_{date}_*.txt'

# 查找map文件
pattern = f'map_{satellite}_{source}_{product}_{date}*.txt'

# 查找GEO图
pattern = f'{satellite}_COCTS_{source}_{product}_GEO_*.jpg'

# 查找PDF报告
pattern = f'{satellite}_COCTS_{source}_val_report_{date}*.pdf'
```

---

## 6. 关键代码模块

### 6.1 H1CD.py 关键函数

| 函数名 | 行号范围 | 功能 |
|--------|----------|------|
| `main()` | 100-200 | 主流程控制 |
| `time_match()` | 500-800 | 时间匹配 |
| `space_match()` | 900-1200 | 空间匹配 |
| `satellite_validation()` | 1830-1990 | 生成valresult |
| `xc_validation()` | 1995-2100 | XC验证 |
| `step7()` | 2110-2500 | 生成地理图 |
| `match_coordinates()` | 2210-2270 | 坐标匹配 |
| `plot_error_map()` | 2298-2440 | 绘制地理图 |
| `step8()` | 2500-2700 | 生成时序图 |
| `step9()` | 2800-3500 | 生成统计图表 |
| `make_sat_report_data()` | 3600-3800 | 生成report文件 |
| `rename_output_files()` | 4650-4750 | 文件重命名 |
| `organize_files()` | 4750-4850 | 文件整理 |

### 6.2 generate_monthly_report_v5.py 结构

```python
class FontManager:
    """字体管理"""
    def _load_font(self):
        # 加载中文字体

class ReportClassifier:
    """报告分类器"""
    @classmethod
    def parse_report_filename(cls, filename):
        # 解析PDF报告文件名
        # 返回: {satellite, source, date, time}
    
    @classmethod
    def get_validation_label(cls, satellite, source):
        # 生成验证标签: "HY1C-AQUA" 或 "HY1C-现场"

class DataAggregator:
    """数据聚合器"""
    def aggregate_validation_data(self, classified_reports):
        # 聚合多个日报告的数据
        # 返回: {validation_label: {products: {...}}}
    
    def _aggregate_products(self, date_reports, satellite, source):
        # 按产品聚合数据
        # 读取report文件，提取bias/RMS/n
    
    def _parse_report_file(self, filepath):
        # 解析report文件
        # 返回: {bias, rms, n}

class ChartGenerator:
    """图表生成器"""
    def generate_pie_chart(self, satellite, source, product, matchup_files, timestamp):
        # 从matchup文件生成饼图
        # 误差区间: 0~5%, 5~10%, 10~15%, 15~20%, ≥20%
    
    def generate_geo_map(self, satellite, source, product, date_strs, timestamp):
        # 从map文件生成地理分布图
        # 1. 查找所有日期的map文件
        # 2. 读取并聚合数据
        # 3. 绘制地理图
    
    def _read_map_file(self, filepath):
        # 读取map文件
        # 返回: latitudes, longitudes, errors
    
    def _plot_error_map(self, latitudes, longitudes, errors, product, output_path):
        # 绘制地理图
        # 支持Basemap或简化版scatter plot

class TemplateFiller:
    """模板填充器"""
    def get_template_path(self, source):
        # 根据验证源选择模板
    
    def generate_replacements(self, ...):
        # 生成替换字典
        # 返回: {text: {...}, tables: {...}, images: {...}}
    
    def fill_template(self, template_path, output_docx, replacements):
        # 填充Word模板

class MonthlyReportGenerator:
    """主生成器"""
    def __init__(self, report_type, year, month=None, quarter=None, ...):
        # 初始化配置和组件
    
    def scan_reports(self):
        # 扫描日报告
        # 返回: classified_reports
    
    def run(self):
        # 主流程
        # 1. 扫描报告
        # 2. 聚合数据
        # 3. 生成图表和报告
```

---

## 7. 月季年报告生成

### 7.1 完整流程

```
步骤1: 扫描PDF报告
  ↓
从文件名解析: satellite, source, date, time
  ↓
分类: {validation_label: {date: [reports]}}
  ↓
步骤2: 聚合数据
  ↓
读取report文件 → 提取bias/RMS/n
  ↓
按产品加权平均: weighted_bias, weighted_rms, total_n
  ↓
步骤3: 生成图表
  ↓
饼图: 从matchup文件读取误差分布
  ↓
GEO图: 从map文件聚合多天坐标和误差
  ↓
步骤4: 填充模板
  ↓
替换占位符: {{date}}, {{val_results_sst}}, {{hy1c_vs_sst_terra_geo}}
  ↓
生成Word文档
```

### 7.2 关键数据聚合逻辑

```python
# 加权平均
total_n = sum(n_list)
weighted_bias = sum(bias * n for bias, n in zip(bias_list, n_list)) / total_n
weighted_rms = sum(rms * n for rms, n in zip(rms_list, n_list)) / total_n

# 示例
# Day 1: bias=0.5, n=1000
# Day 2: bias=0.3, n=2000
# 加权平均: (0.5*1000 + 0.3*2000) / 3000 = 0.367
```

### 7.3 图表生成策略

**饼图**:
```python
# 从matchup文件读取所有误差值
differences = []
for matchup_file in matchup_files:
    # 读取第4列（误差%）
    diffs = parse_matchup_file(matchup_file)
    differences.extend(diffs)

# 统计分布
counts = {
    '0~5%': count(abs(d) < 5 for d in differences),
    '5~10%': count(5 <= abs(d) < 10 for d in differences),
    ...
}

# 绘制饼图
plt.pie(counts.values(), labels=counts.keys())
```

**GEO图**:
```python
# 从map文件聚合所有坐标和误差
all_lats, all_lons, all_errs = [], [], []

for date in dates:
    map_files = glob(f'map_{satellite}_{source}_{product}_{date}*.txt')
    
    for map_file in map_files:
        lats, lons, errs = read_map_file(map_file)
        all_lats.extend(lats)
        all_lons.extend(lons)
        all_errs.extend(errs)

# 绘制地理图
plt.scatter(all_lons, all_lats, c=all_errs, cmap='jet')
```

### 7.4 模板占位符

**文本占位符**:
```
{{date}}          → "2025年10月"
{{satellite_type}} → "HY1C"
{{source_type}}   → "AQUA"
{{unit}}          → "℃"
```

**表格占位符**:
```
{{val_results_sst}}  → 验证结果表格
{{col_results_sst}}  → 匹配结果表格
```

**图片占位符**:
```
{{hy1c_vs_sst_terra_sct}}  → 饼图路径
{{hy1c_vs_sst_terra_geo}}  → GEO图路径
```

**注意**: 占位符中的`terra`是固定的，不随source变化！

---

## 8. 常见问题

### 8.1 为什么没有GEO图？

**可能原因**:
1. ❌ map文件不存在 → 检查`output/04_visualization/map_*.txt`
2. ❌ map文件命名不匹配 → 检查文件名格式
3. ❌ lat/lon文件位置错误 → 应在`input/02_reference/02_reference_preprocess/`
4. ❌ 坐标匹配失败 → 检查lat/lon文件格式和维度

**调试步骤**:
```bash
# 1. 检查map文件
ls -lh output/04_visualization/map_*.txt

# 2. 查看map文件内容
head -5 output/04_visualization/map_HY1C_AQUA_sst_*.txt

# 3. 运行调试模式
python generate_monthly_report_v5.py --month 202510 --debug
```

### 8.2 为什么有些产品没有数据？

**可能原因**:
1. ❌ 该产品没有验证源（如Rrs系列只有AQUA/TERRA）
2. ❌ 数据文件不完整
3. ❌ 时空匹配失败（窗口设置不当）

### 8.3 文件在哪里？

| 寻找目标 | 位置 | 文件名模式 |
|----------|------|------------|
| **matchup文件** | `output/03_collocation/` | `*_matchup_*.txt` |
| **map文件** | `output/04_visualization/` | `map_*.txt` |
| **GEO图** | `output/04_visualization/` | `*_GEO_*.jpg` |
| **PIE图** | `output/04_visualization/` | `*_PIE_*.jpg` |
| **report文件** | `output/05_reports/` | `*_report_*.txt` |
| **PDF报告** | `output/05_reports/` | `*_val_report_*.pdf` |
| **lat/lon文件** | `input/02_reference/02_reference_preprocess/` | `AQUA_Lat_*.txt`, `TERRA_Lon_*.txt` |

### 8.4 为什么lat/lon在reference目录？

**答案**: lat/lon文件对应的是**参考卫星**（AQUA/TERRA）的坐标，不是HY1C的坐标。

- HY1C数据与AQUA数据匹配时，使用**AQUA的lat/lon**
- HY1C数据与TERRA数据匹配时，使用**TERRA的lat/lon**

**原因**: 在step7中，map文件的坐标是从参考卫星的观测位置提取的。

---

## 9. 调试指南

### 9.1 开启调试模式

```bash
python generate_monthly_report_v5.py --month 202510 --debug
```

### 9.2 调试输出解读

**正常输出**:
```
[DEBUG] ========== 处理产品: sst ==========
[DEBUG] 找到matchup: ['HY1C_COCTS_AQUA_sst_matchup_20251009_104500.txt']
[DEBUG] 总计matchup: 2
[INFO] ✓ 生成饼图: HY1C_COCTS_AQUA_sst_PIE_20251212_190000.jpg

[DEBUG] ========== 开始生成GEO图 ==========
[DEBUG] 产品: HY1C_AQUA_sst
[DEBUG] 日期数: 2
[DEBUG] 找到map文件: ['map_HY1C_AQUA_sst_20251009104500.txt']
[DEBUG] 总计map文件: 2
[DEBUG] 读取map文件: map_HY1C_AQUA_sst_20251009104500.txt
[DEBUG] 读取到 15832 个数据点
[DEBUG] 总共聚合 30088 个数据点
[INFO] ✓ 生成地理图: HY1C_COCTS_AQUA_sst_GEO_20251212_190000.jpg
```

**异常输出**:
```
[DEBUG] ========== 开始生成GEO图 ==========
[DEBUG] 未找到map文件
```
→ **问题**: map文件不存在，检查`output/04_visualization/`

```
[DEBUG] 读取到 0 个数据点
```
→ **问题**: map文件为空或格式错误

```
[DEBUG] 总共聚合 0 个数据点
```
→ **问题**: 所有map文件都无法读取

### 9.3 手动验证流程

```bash
# 1. 检查目录结构
tree -L 2 output/

# 2. 统计文件数量
echo "matchup文件数:"
ls output/03_collocation/*_matchup_*.txt 2>/dev/null | wc -l

echo "map文件数:"
ls output/04_visualization/map_*.txt 2>/dev/null | wc -l

echo "GEO图数:"
ls output/04_visualization/*_GEO_*.jpg 2>/dev/null | wc -l

# 3. 查看文件示例
echo "matchup文件示例:"
head -20 output/03_collocation/*_matchup_*.txt | head -20

echo "map文件示例:"
head -10 output/04_visualization/map_*.txt | head -10

# 4. 检查lat/lon文件
echo "lat/lon文件:"
ls input/02_reference/02_reference_preprocess/*_Lat_*.txt
ls input/02_reference/02_reference_preprocess/*_Lon_*.txt
```

### 9.4 常用诊断命令

```bash
# 查找特定日期的文件
find output/ -name "*20251009*" -type f

# 统计每个产品的文件数
for product in sst chl AOT Rrs412 Rrs443 Rrs490 Rrs520 Rrs565 Rrs670; do
    count=$(ls output/04_visualization/map_*_${product}_*.txt 2>/dev/null | wc -l)
    echo "$product: $count"
done

# 检查文件大小（空文件检测）
find output/04_visualization/ -name "map_*.txt" -size 0

# 验证文件格式
awk 'NF!=3 {print FILENAME":"NR": 字段数="NF}' output/04_visualization/map_*.txt | head
```

---

## 10. 快速参考卡

### 10.1 核心数据流

```
原始数据 → H1CD.py → 中间文件 → 月报生成器 → Word报告
   ↓                    ↓              ↓
HY1C.txt           matchup.txt      饼图/GEO图
AQUA.txt           map.txt          Word文档
TERRA.txt          report.txt
```

### 10.2 关键文件速查

| 文件 | 用途 | 位置 | 格式 |
|------|------|------|------|
| **matchup** | 饼图数据 | 03_collocation | 行号\tHY值\t参考值\t误差% |
| **map** | GEO图数据 | 04_visualization | 纬度\t经度\t误差% |
| **report** | 统计数据 | 05_reports | header格式 |
| **lat/lon** | 坐标数据 | 02_reference_preprocess | Tab分隔矩阵 |

### 10.3 验证源配置

| 验证源 | 产品 | 模板 |
|--------|------|------|
| **AQUA** | sst, chl, AOT, Rrs系列 | new_auqa_terra_template.docx |
| **TERRA** | sst, chl, AOT, Rrs系列 | new_auqa_terra_template.docx |
| **SNPP** | sst, chl, AOT | new_snpp_jpss_template.docx |
| **JPSS** | sst, chl, AOT | new_snpp_jpss_template.docx |
| **XC** | sst, chl, AOT | new_xc_template.docx |

### 10.4 命令速查

```bash
# 生成月报告
python generate_monthly_report_v5.py --month 202510 --debug

# 生成季度报告
python generate_monthly_report_v5.py --quarter 2025 1 --debug

# 生成年报告
python generate_monthly_report_v5.py --year 2025 --debug

# 运行H1CD主流程
python H1CD.py

# 检查配置
cat config.ini
```

---

## 11. 代码演进历史

### V1-V3: 尝试从valresult/matchup重建坐标
- ❌ 复杂度高
- ❌ 需要读取lat/lon文件
- ❌ 需要匹配行号到坐标

### V4: 从matchup文件读取行号匹配坐标
- ✅ 逻辑正确
- ❌ 但忽略了map文件已经存在

### V5: 直接从map文件读取坐标和误差 ⭐⭐⭐
- ✅ 最简洁
- ✅ 最高效
- ✅ map文件已包含所有需要的数据

---

## 12. 联系与支持

如有问题，请参考：
1. 本文档第8节"常见问题"
2. 本文档第9节"调试指南"
3. 代码中的注释和docstring

---

**文档版本**: 1.0  
**最后更新**: 2024-12-14  
**适用代码版本**: generate_monthly_report_v5.py  
**作者**: Claude (Session 2024-12-12)

---
