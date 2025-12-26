import pandas as pd
import os

def process_gene_data(input_excel, output_dir):
    """
    完整处理基因数据的流程：
    1. 筛选原始Excel文件中的基因ID
    2. 将筛选结果合并为单个工作表
    3. 分割合并后的工作表为A1-A8的条件工作表
    
    参数:
    input_excel -- 原始Excel文件路径
    output_dir -- 输出目录路径
    """

    os.makedirs(output_dir, exist_ok=True)
    

    print(f"步骤1: 筛选基因ID - 正在处理 {input_excel}")
    filtered_excel = os.path.join(output_dir, "filtered_gene_ids.xlsx")
    

    xls = pd.ExcelFile(input_excel)
    with pd.ExcelWriter(filtered_excel, engine='xlsxwriter') as writer:
        for sheet_name in xls.sheet_names:
            data = pd.read_excel(xls, sheet_name=sheet_name)
            
 
            condition1 = data[(data.iloc[:, 3] < 0.05) & (abs(data.iloc[:, 7]) > 0.858)].iloc[:, 0].tolist()
            condition2 = data[(data.iloc[:, 9] < 0.05) & (abs(data.iloc[:, 13]) > 0.858)].iloc[:, 0].tolist()
            

            df_condition1 = pd.DataFrame({'第四列小于0.05且第八列绝对值大于0.858的基因ID': condition1})
            df_condition2 = pd.DataFrame({'第十列小于0.05且第十四列绝对值大于0.085的基因ID': condition2})
            

            df_condition1.to_excel(writer, sheet_name=f'{sheet_name}_条件1', index=False)
            df_condition2.to_excel(writer, sheet_name=f'{sheet_name}_条件2', index=False)
    
    print(f"筛选完成，结果保存到: {filtered_excel}")
    

    print("步骤2: 合并工作表")
    combined_excel = os.path.join(output_dir, "combined_gene_ids.xlsx")
    
    excel_data = pd.ExcelFile(filtered_excel)
    result_data = []
    
    for sheet_name in excel_data.sheet_names:
        df = pd.read_excel(excel_data, sheet_name=sheet_name)
        if not df.empty:
            genes = df.iloc[:, 0].dropna().astype(str).tolist()
            genes_str = ' '.join(genes)
            result_data.append({'Sheet Name': sheet_name, 'Genes': genes_str})
    
    result_df = pd.DataFrame(result_data)
    result_df.to_excel(combined_excel, index=False)
    print(f"合并完成，结果保存到: {combined_excel}")
    

    print("步骤3: 分割工作表")
    final_output = os.path.join(output_dir, "final_output.xlsx")
    
    df = pd.read_excel(combined_excel)
    conditions = ['条件1', '条件2']
    
    with pd.ExcelWriter(final_output, engine='openpyxl', mode='w') as writer:
        for i in range(1, 9):
            sheet_name = f'A{i}'
            for condition in conditions:
                filtered_df = df[df['Sheet Name'].str.contains(sheet_name) & 
                                df['Sheet Name'].str.contains(condition)]
                
                if not filtered_df.empty:
                    output_sheet_name = f'{sheet_name}_{condition}'
                    filtered_df.to_excel(writer, sheet_name=output_sheet_name, index=False)
    
    print(f"分割完成，最终结果保存到: {final_output}")
    print("所有处理步骤已完成!")


if __name__ == '__main__':
    # 输入文件路径
    input_file = 'data.xlsx'  # 替换为您的文件路径
    

    output_directory = 'results'  # 替换为您想要的输出目录
    

    process_gene_data(input_file, output_directory)
