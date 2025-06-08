import pandas as pd
import glob
import os

# --- 配置 ---
# 获取当前脚本所在的文件夹路径
# 所有 .xlsx 文件都应存放在这个文件夹里
directory = os.path.dirname(os.path.abspath(__file__))

# 定义合并后输出的文件名
output_filename = '合并统计结果.xlsx'

# --- 主程序 ---
try:
    # 查找目录下所有的 .xlsx 文件
    # glob.glob 会返回一个包含所有匹配文件路径的列表
    excel_files = glob.glob(os.path.join(directory, '*.xlsx'))

    # 如果找到了文件
    if not excel_files:
        print("错误：在该目录下没有找到任何 .xlsx 文件。")
        input("请按回车键退出。")
    else:
        print(f"找到了 {len(excel_files)} 个 Excel 文件，正在处理...")

        # 创建一个空的 DataFrame 用于存放所有表格的数据
        all_data = pd.DataFrame()

        # 循环读取每一个 Excel 文件
        for file in excel_files:
            # 读取时，我们假定第一列是国家，第二列是数量，并且没有表头
            # 如果您的 Excel 有表头（像图中那样），pandas 会自动识别
            # 为确保统一，我们直接指定列名
            df = pd.read_excel(file)
            all_data = pd.concat([df], ignore_index=True)

        # 检查合并后的数据是否为空
        if all_data.empty:
            print("错误：读取到的所有 Excel 文件都是空的。")
            input("请按回车键退出。")
        else:
            # 将第一列和第二列重命名，以方便后续处理
            # 这里我们假设第一列是'国家'，第二列是'数量'
            # 您可以根据实际情况修改列名
            all_data.columns = ['国家', '数量']

            # 按“国家”列进行分组，并对“数量”列求和
            # .reset_index() 会将分组结果重新转换为 DataFrame
            merged_data = all_data.groupby('国家')['数量'].sum().reset_index()

            # 为了美观，可以按“数量”降序排序
            merged_data = merged_data.sort_values(by='数量', ascending=False)

            # 将最终结果保存到一个新的 Excel 文件中
            # index=False 表示在输出的 Excel 中不包含行索引号
            output_path = os.path.join(directory, output_filename)
            merged_data.to_excel(output_path, index=False)

            print("-" * 30)
            print(f"任务完成！\n合并后的数据已成功保存到文件：{output_filename}")
            input("请按回车键退出。")

except Exception as e:
    print(f"处理过程中发生错误: {e}")
    input("请按回车键退出。")