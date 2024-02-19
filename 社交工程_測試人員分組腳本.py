import pandas as pd
import csv, random

def random_grouping_no_repeat(input_list, num_groups):
    # 確認List中的項目數大於等於分組數
    assert len(input_list) >= num_groups, "List中的項目數應大於等於分組數"
    
    # 初始化一個空的分組列表
    grouped_list = [[] for _ in range(num_groups)]
    
    # 將每個項目隨機分配到一個組中，確保不重複
    for item in input_list:
        group_index = random.randint(0, num_groups - 1)
        
        while item in grouped_list[group_index]:
            group_index = random.randint(0, num_groups - 1)
        
        grouped_list[group_index].append(item)
    
    return grouped_list

# 指定Excel文件的路徑
excel_file_path = "113年社交工程測試人員清單.xlsx"

# 使用Pandas的read_excel函數讀取Excel文件
data = pd.read_excel(excel_file_path)

# 選擇 data 中的 Email 欄位
emails = data['Email']

email_list = emails.tolist()

# 隨機打亂 email_list 順序
email_list = random.shuffle( email_list )

# 總共要分成 40 組
num_groups = 40

result = random_grouping_no_repeat(email_list, num_groups)

# 打印結果
print(result)

output_directory = "./"

# 將每組分別寫入文件
for i, group in enumerate(result, start=1):
    # 生成文件名，例如 group1.txt, group2.txt, ...
    file_name = f'group{i}.csv'
    
    # 完整的文件路徑
    file_path = output_directory + file_name
    
    # 開啟文件並寫入結果
    with open(file_path, 'w') as file:
        writer = csv.writer( file )
        writer.writerow( ["First Name", "Last Name", "Email", "Position"] )
        for mail in group:
            writer.writerow( [ "", "", mail, "" ] )

    print(f"組 {i} 的結果已儲存至 {file_path}")
