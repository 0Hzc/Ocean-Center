#!/usr/bin/env python3
"""
检查HDF5文件中IPAR数据集的属性，并定位inf值的具体位置
"""
import h5py
import numpy as np
import glob
import os

def check_hdf5_dataset(filepath, dataset_path='Geophysical Data/IPAR', max_inf_samples=20):
    """检查HDF5数据集的属性，并详细定位inf值"""
    print(f"\n检查文件: {os.path.basename(filepath)}")
    print(f"数据集路径: {dataset_path}")
    print("-" * 50)

    try:
        with h5py.File(filepath, 'r') as f:
            if dataset_path not in f:
                print(f"  数据集 '{dataset_path}' 不存在")
                # 列出所有数据集
                print("  可用的数据集:")
                def print_datasets(name, obj):
                    if isinstance(obj, h5py.Dataset):
                        print(f"    - {name}")
                f.visititems(print_datasets)
                return

            ds = f[dataset_path]

            print(f"  数据类型: {ds.dtype}")
            print(f"  数据形状: {ds.shape}")

            # 打印所有属性
            print(f"  属性列表:")
            for attr_name in ds.attrs:
                attr_value = ds.attrs[attr_name]
                print(f"    {attr_name}: {attr_value}")

            # 特别检查scale和offset
            slope = ds.attrs.get('Slope', ds.attrs.get('scale_factor', None))
            intercept = ds.attrs.get('Intercept', ds.attrs.get('add_offset', None))
            fill_value = ds.attrs.get('_FillValue', ds.attrs.get('fill_value', None))
            valid_min = ds.attrs.get('ValidMin', None)
            valid_max = ds.attrs.get('ValidMax', None)

            print(f"\n  关键属性:")
            print(f"    Slope/scale_factor: {slope}")
            print(f"    Intercept/add_offset: {intercept}")
            print(f"    _FillValue: {fill_value}")
            print(f"    ValidMin: {valid_min}")
            print(f"    ValidMax: {valid_max}")

            # 读取数据
            data = ds[:].astype(float)

            # 读取经纬度数据（如果存在）
            lat_data = None
            lon_data = None
            try:
                if 'Navigation Data/Latitude' in f:
                    lat_data = f['Navigation Data/Latitude'][:]
                    lon_data = f['Navigation Data/Longitude'][:]
                    print(f"\n  经纬度数据形状: lat={lat_data.shape}, lon={lon_data.shape}")
            except Exception as e:
                print(f"  无法读取经纬度数据: {e}")

            # 统计各类异常值
            nan_mask = np.isnan(data)
            pos_inf_mask = np.isposinf(data)
            neg_inf_mask = np.isneginf(data)
            fill_mask = (data == fill_value) if fill_value is not None else np.zeros_like(data, dtype=bool)

            nan_count = np.sum(nan_mask)
            pos_inf_count = np.sum(pos_inf_mask)
            neg_inf_count = np.sum(neg_inf_mask)
            fill_count = np.sum(fill_mask)

            print(f"\n  数据统计:")
            print(f"    总像元数: {data.size}")
            print(f"    NaN数量: {nan_count}")
            print(f"    +Inf数量: {pos_inf_count}")
            print(f"    -Inf数量: {neg_inf_count}")
            print(f"    填充值({fill_value})数量: {fill_count}")

            # 有效数据统计
            valid_mask = ~(nan_mask | pos_inf_mask | neg_inf_mask | fill_mask)
            valid_data = data[valid_mask]
            print(f"    有效数据数量: {len(valid_data)}")
            if len(valid_data) > 0:
                print(f"    有效数据范围: [{np.min(valid_data):.6f}, {np.max(valid_data):.6f}]")
                print(f"    有效数据平均值: {np.mean(valid_data):.6f}")

            # ==================== 详细定位inf值 ====================
            print(f"\n" + "="*60)
            print(f"  【详细定位inf值】")
            print("="*60)

            # 找出所有inf的位置
            inf_mask = pos_inf_mask | neg_inf_mask
            inf_indices = np.where(inf_mask)

            if len(inf_indices[0]) > 0:
                print(f"\n  共发现 {len(inf_indices[0])} 个inf值")
                print(f"  显示前 {min(max_inf_samples, len(inf_indices[0]))} 个inf值的详细位置:\n")

                print(f"  {'序号':<6} {'行号':<8} {'列号':<8} {'值':<15} {'经度':<12} {'纬度':<12}")
                print(f"  {'-'*6} {'-'*8} {'-'*8} {'-'*15} {'-'*12} {'-'*12}")

                for i in range(min(max_inf_samples, len(inf_indices[0]))):
                    row = inf_indices[0][i]
                    col = inf_indices[1][i]
                    val = data[row, col]

                    # 获取对应的经纬度
                    if lat_data is not None and lon_data is not None:
                        try:
                            lat = lat_data[row, col] if lat_data.ndim == 2 else lat_data[row]
                            lon = lon_data[row, col] if lon_data.ndim == 2 else lon_data[row]
                            print(f"  {i+1:<6} {row:<8} {col:<8} {str(val):<15} {lon:<12.4f} {lat:<12.4f}")
                        except:
                            print(f"  {i+1:<6} {row:<8} {col:<8} {str(val):<15} {'N/A':<12} {'N/A':<12}")
                    else:
                        print(f"  {i+1:<6} {row:<8} {col:<8} {str(val):<15} {'N/A':<12} {'N/A':<12}")

                # 统计inf值的行分布
                print(f"\n  【inf值的行分布统计】")
                unique_rows, row_counts = np.unique(inf_indices[0], return_counts=True)
                print(f"  inf值分布在 {len(unique_rows)} 行中")

                # 找出inf最多的几行
                top_rows_idx = np.argsort(row_counts)[-10:][::-1]
                print(f"\n  inf数量最多的10行:")
                print(f"  {'行号':<10} {'inf数量':<10} {'占该行比例':<15}")
                for idx in top_rows_idx:
                    row_num = unique_rows[idx]
                    count = row_counts[idx]
                    ratio = count / data.shape[1] * 100
                    print(f"  {row_num:<10} {count:<10} {ratio:.2f}%")

                # 统计inf值的列分布
                print(f"\n  【inf值的列分布统计】")
                unique_cols, col_counts = np.unique(inf_indices[1], return_counts=True)
                print(f"  inf值分布在 {len(unique_cols)} 列中")

                # 找出inf最多的几列
                top_cols_idx = np.argsort(col_counts)[-10:][::-1]
                print(f"\n  inf数量最多的10列:")
                print(f"  {'列号':<10} {'inf数量':<10} {'占该列比例':<15}")
                for idx in top_cols_idx:
                    col_num = unique_cols[idx]
                    count = col_counts[idx]
                    ratio = count / data.shape[0] * 100
                    print(f"  {col_num:<10} {count:<10} {ratio:.2f}%")

            else:
                print(f"\n  未发现inf值")

            # ==================== 检查超出ValidMin/ValidMax范围的值 ====================
            if valid_min is not None and valid_max is not None:
                print(f"\n" + "="*60)
                print(f"  【超出有效范围的值】ValidMin={valid_min}, ValidMax={valid_max}")
                print("="*60)

                # 排除已经是inf/nan/fill_value的数据
                check_mask = ~(nan_mask | pos_inf_mask | neg_inf_mask | fill_mask)
                check_data = data.copy()

                out_of_range_low = (check_data < valid_min) & check_mask
                out_of_range_high = (check_data > valid_max) & check_mask

                low_count = np.sum(out_of_range_low)
                high_count = np.sum(out_of_range_high)

                print(f"\n  低于ValidMin的值数量: {low_count}")
                print(f"  高于ValidMax的值数量: {high_count}")

                if low_count > 0:
                    low_indices = np.where(out_of_range_low)
                    print(f"\n  低于ValidMin的值样例(前10个):")
                    for i in range(min(10, low_count)):
                        row, col = low_indices[0][i], low_indices[1][i]
                        val = data[row, col]
                        print(f"    行{row}, 列{col}: {val}")

                if high_count > 0:
                    high_indices = np.where(out_of_range_high)
                    print(f"\n  高于ValidMax的值样例(前10个):")
                    for i in range(min(10, high_count)):
                        row, col = high_indices[0][i], high_indices[1][i]
                        val = data[row, col]
                        print(f"    行{row}, 列{col}: {val}")

    except Exception as e:
        print(f"  错误: {e}")
        import traceback
        traceback.print_exc()

def main():
    # 搜索L2C文件（包含IPAR）
    search_paths = [
        r"/mnt/d/work/海洋中心/input/*L2C*.h5",
        r"/mnt/d/work/海洋中心/input/*L2C*.HDF",
        r"/mnt/d/work/海洋中心/input/**/*L2C*.h5",
        r"D:/work/海洋中心/input/*L2C*.h5",
    ]

    files = []
    for pattern in search_paths:
        files.extend(glob.glob(pattern, recursive=True))

    if not files:
        print("未找到L2C文件，请手动指定文件路径")
        print("用法: python check_hdf5_ipar.py <hdf5_file_path>")
        return

    print("="*60)
    print(" HDF5 IPAR数据集属性检查")
    print("="*60)

    for filepath in files[:2]:  # 最多检查2个文件
        check_hdf5_dataset(filepath, 'Geophysical Data/IPAR')

        # 也检查SST数据集作为对比
        print("\n" + "="*60)
        print(" 对比: SST数据集")
        sst_file = filepath.replace('L2C', 'L2T')
        if os.path.exists(sst_file):
            check_hdf5_dataset(sst_file, 'Geophysical Data/SST')
        else:
            print(f"  SST文件不存在: {sst_file}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        # 使用命令行参数指定的文件
        check_hdf5_dataset(sys.argv[1], 'Geophysical Data/IPAR')
    else:
        main()
