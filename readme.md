#月季年生成命令

1. 生成月报告（2023年10月）
python generate_report.py --month 202310

2. 生成季报告（2023年第4季度）
python generate_report.py --quarter 2023 4

3. 生成年报告（2023年）
python generate_report.py --year 2023

4. 指定输出目录
python generate_report.py --month 202310 --output /path/to/output

5. 启用调试模式
python generate_report.py --month 202310 --debug