#!/usr/bin/env python3
"""
检查HDF5文件中IPAR数据集的属性
"""
import h5py
import numpy as np
import glob
import os

def check_hdf5_dataset(filepath, dataset_path='Geophysical Data/IPAR'):
    """检查HDF5数据集的属性"""
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

            print(f"\n  关键属性:")
            print(f"    Slope/scale_factor: {slope}")
            print(f"    Intercept/add_offset: {intercept}")
            print(f"    _FillValue: {fill_value}")

            # 读取数据样本
            data = ds[:]
            print(f"\n  原始数据统计:")
            print(f"    最小值: {np.nanmin(data)}")
            print(f"    最大值: {np.nanmax(data)}")
            print(f"    NaN数量: {np.sum(np.isnan(data.astype(float)))}")
            print(f"    样本值(前10个非零): {data.flatten()[data.flatten() != 0][:10]}")

            # 如果有slope和intercept，计算转换后的值
            if slope is not None and intercept is not None:
                # 获取标量值
                if hasattr(slope, '__len__'):
                    slope = slope[0]
                if hasattr(intercept, '__len__'):
                    intercept = intercept[0]

                print(f"\n  应用转换公式: value = data * {slope} + {intercept}")

                # 转换数据
                converted = data.astype(float) * float(slope) + float(intercept)

                # 处理fill_value
                if fill_value is not None:
                    if hasattr(fill_value, '__len__'):
                        fill_value = fill_value[0]
                    converted[data == fill_value] = np.nan

                valid_data = converted[~np.isnan(converted)]
                if len(valid_data) > 0:
                    print(f"  转换后数据统计:")
                    print(f"    最小值: {np.nanmin(valid_data)}")
                    print(f"    最大值: {np.nanmax(valid_data)}")
                    print(f"    平均值: {np.nanmean(valid_data)}")
                    print(f"    样本值(前10个): {valid_data[:10]}")

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
        check_hdf5_dataset(filepath.replace('L2C', 'L2T'), 'Geophysical Data/SST')

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        # 使用命令行参数指定的文件
        check_hdf5_dataset(sys.argv[1], 'Geophysical Data/IPAR')
    else:
        main()
