# coding=utf8
from base.tb_base import get_files_list, append_data_to_worksheet
from base.systemlogger import Logger
import pandas as pd
import os
import shutil
import datetime

didict = {
    "marks": "",                           # 入仓号
    "shipper": "",                         # data.xlsx 第4列 index=3
    "ctns": "",                            # data.xlsx 第6列 index=5
    "lie1": "",                            # 中英文品名DESCRIPTIONS OF GOODS
    "kgs": "",                             # data.xlsx 第7列 index=6
    "cm3": "",                             # data.xlsx 第8列 index=7
    "shipment_id": "",                     # 原单号 货件编号SHIPMENT ID shipment_id
    "amazon_reference_id": "",             # 原单号 亚马逊内部编码AMAZON REFERENCE ID amazon_reference_id
    "fba_warehouse_address": "",           # 原单号  fba_warehouse_address
    "direct_ltl_or_ups": "",               # 基于本表ups_ltl_number, 有单号填UPS, 没单号填DIRECT LTL
    "ups_ltl_number": "",                  # data.xlsx 第11列 index=10
    "rate": "",                            # data.xlsx 第12列 index=11
    "create_date": "",                     # data.xlsx 第13列 index=12
    "appt_deadline": "",                   # 放空
    "remark": "",                          # 放空
}


dixlsx = "delivery_instruction.xlsx"
if os.path.exists(dixlsx):
    os.makedirs("bak", exist_ok=True)
    shutil.move(f"{dixlsx}", f"bak/{dixlsx}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
shutil.copy2(f"mb/{dixlsx}", ".")

'''
# 读取 Excel 文件, 显示第一行为序号
df = pd.read_excel('data.xlsx', header=None)
#print(df)
# 筛选第一列中包含指定字符串的行
result = df[df.iloc[:, 0].str.contains('OB-23030360')]
# 输出结果
print(result)
'''

df = pd.read_excel('data.xlsx', header=None)
#print(df)

txt_lists = get_files_list("obs", "txt")
count_piao = 0
for i, txt_file in enumerate(txt_lists):
    i = i + 1
    ob_piao = 0
    Logger.ins().std_logger().info(f"开始解析第{i}个ob： {txt_file}...")
    with open(txt_file, "r", encoding="utf-8") as file:
        for line in file:
            if line.strip() != "":
                count_piao = count_piao + 1
                ob_piao = ob_piao + 1
                line_dict = eval(line.strip().replace("nan", "''"))  # 转移Nan为空
                didict["marks"] = line_dict["warehouse_number"]
                # 筛选第一列中包含指定字符串的行
                data_line = df[df.iloc[:, 0].str.contains(didict["marks"])]
                print(didict)
                # 输出结果
                
                print(data_line[5].item())
                didict["shipper"] = data_line[3].item()
                didict["ctns"] = str(data_line[5].item())
                didict["lie1"] = str(line_dict["description_of_goods"])
                didict["kgs"] = str(data_line[6].item())
                didict["cm3"] = str(data_line[7].item())
                didict["shipment_id"] = str(line_dict["shipment_id"])
                didict["amazon_reference_id"] =  str(line_dict["amazon_reference_id"])
                didict["fba_warehouse_address"] = str(line_dict["fba_warehouse_address"])
                didict["ups_ltl_number"] = str(data_line[10].item())
                if data_line[10].item() == "" or str(data_line[10].item()).lower() == "nan":
                    didict["direct_ltl_or_ups"] = "DIRECT LTL"
                else:
                    didict["direct_ltl_or_ups"] = "UPS"                    
                didict["rate"] = str(data_line[11].item())
                didict["create_date"] = str(data_line[12].item())
                didict["appt_deadline"] = ""
                didict["remark"] = ""  
                print(didict)

                append_data_to_worksheet(
                    dixlsx,
                    "DELIVERY INSTRUCTION",
                    [
                        [
                            list(didict.values())[0],
                            list(didict.values())[1],
                            list(didict.values())[2],
                            list(didict.values())[3],
                            list(didict.values())[4],
                            list(didict.values())[5],
                            list(didict.values())[6],
                            list(didict.values())[7],
                            list(didict.values())[8],
                            list(didict.values())[9],
                            list(didict.values())[10],
                            list(didict.values())[11],
                            list(didict.values())[12],
                            list(didict.values())[13],
                            list(didict.values())[14],
                        ]
                    ],
                )
                