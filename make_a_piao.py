# coding=utf8
from base.tb_base import get_files_list, append_data_to_worksheet
from base.systemlogger import Logger
import os
import shutil
import datetime

adict = {
    "marks": "",                      # 入仓号
    "no": "",                         # 箱数CTN
    "ctns": "CTNS",                   # CTNS(固定)
    "lie1": "",                       # 中英文品名DESCRIPTIONS OF GOODS
    "no2": "",                        # 总数量PCS
    "pcs": "PCS",                     # PCS(固定)
    "nw_kgs": "",                     # 净重KGS
    "kgs1": "KGS",                    # KGS(固定)
    "gw_kgs": "",                     # 毛重KGS
    "kgs2": "KGS",                    # KGS(固定)
    "usd": "USD",                     # USD(固定)
    "unit_price": "",                 # 清关申报单价（USD）
    "amount": "",                     # 总数量PCS * 清关申报单价（USD）
    "hts": "",                        # 清关编码H.S CODE
    "details": "",                    # 装饰n|树脂|无型号|无享惠|无品牌
}

dixlsx = "A票.xlsx"
if os.path.exists(dixlsx):
    os.makedirs("bak", exist_ok=True)
    shutil.move(f"{dixlsx}", f"bak/{dixlsx}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
shutil.copy2(f"mb/{dixlsx}", ".")

txt_lists = get_files_list("obs", "txt")
count_piao = 0
for i, txt_file in enumerate(txt_lists):
    i = i + 1
    ob_piao = 0
    with open(txt_file, "r", encoding="utf-8") as file:
        for line in file:
            if line.strip() != "":
                count_piao = count_piao + 1
                ob_piao = ob_piao + 1
                line_dict = eval(line.strip().replace("nan", "''"))  # 转移Nan为空
                
                adict["marks"] = str(line_dict["warehouse_number"])
                adict["no"] = str(line_dict["number_of_boxes_ctn"])
                adict["ctns"] = "CTNS"
                adict["lie1"] = str(line_dict["description_of_goods"])
                adict["no2"] = str(line_dict["total_pcs"])
                adict["pcs"] = "PCS"
                adict["nw_kgs"] = str(line_dict["net_weight_kgs"])
                adict["kgs1"] = "KGS"
                adict["gw_kgs"] = str(line_dict["gross_weight_kgs"])
                adict["kgs2"] = "KGS"
                adict["usd"] = "USD"
                adict["unit_price"] = str(line_dict["customs_declaration_unit_price_usd"])
                try:
                    adict["amount"] = int(line_dict["total_pcs"]) * int(line_dict["customs_declaration_unit_price_usd"].replace("美金", "").replace("美元", "").replace("$", ""))
                except:
                    adict["amount"] = "0"
                adict["hts"] = str(line_dict["hs_code"])

                """
                Last cloumn
                手机触控用  中英文用途（必填）USE FOR  "穿搭 Wear" 直接copy
                铝制： Synthetic leather + rubber/合成革＋橡胶 直接copy 
                无型号：如果无和没填就写无型号，如果有型号结尾就copy，如果没有型号结果”copy+型号“
                无品牌：
                不享惠
                品牌类型
                """
                detail_1 = line_dict["use_in_chinese_and_english_required"]
                detail_2 = line_dict["material_required"]
                detail_3 = "无型号" if line_dict["model_number_required"] in ["无", ""] else line_dict["model_number_required"]
                detail_4 = "无享惠" if line_dict["export_preferential_treatment_situation"] in ["无", ""] else line_dict["export_preferential_treatment_situation"]
                detail_5 = "无品牌" if line_dict["brand_required"] in ["无", ""] else line_dict["brand_required"]
                adict["details"] = f"{detail_1}|{detail_2}|{detail_3}|{detail_4}|{detail_5}"
                value_list = list(adict.values())
                print(value_list)

                append_data_to_worksheet(
                    "A票.xlsx",
                    "总表",
                    [
                        [
                            list(adict.values())[0],
                            list(adict.values())[1],
                            list(adict.values())[2],
                            list(adict.values())[3],
                            list(adict.values())[4],
                            list(adict.values())[5],
                            list(adict.values())[6],
                            list(adict.values())[7],
                            list(adict.values())[8],
                            list(adict.values())[9],
                            list(adict.values())[10],
                            list(adict.values())[11],
                            list(adict.values())[12],
                            list(adict.values())[13],
                            list(adict.values())[14],
                        ]
                    ],
                )
                Logger.ins().std_logger().info(f"[INFO] {adict}")
        Logger.ins().std_logger().info(f"[INFO] 开始读取第{i}个OB:{txt_file}, 包含{ob_piao}票")
Logger.ins().std_logger().info(f"[INFO] 一共有{i}个OB单号, 共{count_piao}票数据")
