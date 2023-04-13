import pandas as pd
from base.tb_base import get_files_list
from base.systemlogger import Logger
import re

d0 = {
    "warehouse_number": "",                         # 入仓号  
    "description_of_goods": "",                     # 中英文品名DESCRIPTIONS OF GOODS
    "total_pcs": "",                                # 总数量PCS
    "actual_declared_unit": "",                     # 实际申报单位
    "number_of_boxes_ctn": "",                      # 箱数CTN
    "net_weight_kgs": "",                           # 净重KGS
    "gross_weight_kgs": "",                         # 毛重KGS
    "measurement_cbm": "",                          # 体积Measurement (CBM)
    "shipment_id": "",                              # 货件编号SHIPMENT ID
    "amazon_reference_id": "",                      # 亚马逊内部编码AMAZON REFERENCE ID	
    "customs_declaration_unit_price_usd": "",       # 清关申报单价（USD）
    "customs_declaration_total_amount_usd": "",     # 清关申报总金额（USD）
    "chinese_customs_code": "",                     # 报关编码Chinese customs code
    "hs_code": "",                                  # 清关编码H.S CODE
    "material_required": "",                        # 中英文材质（必填） MATERIAL
    "actual_size_of_goods": "",                     # 实际货物的规格（有规格的要写）长宽高，货物上有LOGO也要体现，包括一句话
    "model_number_required": "",                    # 型号（必填）
    "brand_required": "",                           # 品牌（必填
    "use_in_chinese_and_english_required": "",      # 中英文用途（必填）USE FOR
    "brand_type": "",                               # 品牌类型（无品牌，境内自主品牌，境外收购品牌，境内品牌（贴牌生产），境外品牌（其他）
    "export_preferential_treatment_situation": "",  # 出口享惠情况（不享受，不确定，享受）
    "product_picture_required": "",                 # 图片（必填）PRODUCT PICTURE
    "fba_warehouse_address": "",                    # FBA仓库地址FBA address
    "other_declaration_elements": "",               # 其他申报要素项（自行审查）
    "payment_tax_refund": "",                       # 买单/退税	
    "remarks": ""                                   # 备注		
}

list_xlsx = get_files_list()
for j, excel in enumerate(list_xlsx):
    j = j + 1
    Logger.ins().std_logger().info(f"开始解析第{j}个ob： {excel}...")
    file = excel.replace("xlsx", "txt")
    with open(file, "w", encoding="utf-8") as f:
        f.write("")
    # 读取Excel文件
    sheet_names = pd.ExcelFile(excel).sheet_names
    if len(sheet_names) == 1:
        df = pd.read_excel(excel)
    else:
        pattern = re.compile(r"报关.*资料")
        sheet_names = list(filter(pattern.search, sheet_names))
        df = pd.concat([pd.read_excel(excel, sheet_name=sheet_name) for sheet_name in sheet_names])
    count = 0
    # 遍历每一行
    for index, row in df.iterrows():
        #print(row)

        if int(str(row).count("NaN")) < 15 and row[0] != "入仓号": 
            Logger.ins().std_logger().info(f"开始解析第{index}行：有{len(row)}列, 内容为{row}...")
            row_lenth = len(row)
            d1=d0
            if str(row[0]).lower() == "nan" or "ob-" not in str(row[0]).lower(): #如果入仓号为空或者没有ob则取excel名字的OB名字
                match = re.search(r'OB-\d+', file)  #如果不符合说明名字不符合标准，需要手工检查
                if match:
                    row[0] =  match.group(0)
            if row_lenth == 26: # 如果是26列属于标准
                for i, value in enumerate(row):
                    key = list(d0.keys())[i]
                    d1[key]= value
                #print(d1)
                with open(file, "a", encoding="utf-8") as f:
                    f.write(str(d1) + "\n")
                    
            if row_lenth == 27:  # 如果是27列则删除第二列: 客户公司名称
                z = 0 # 重新定义一个序列号
                for i, value in enumerate(row):  
                    if i == 2: # 当读取到第二列时候 将z序列减去一个 覆盖之前的赋值达到忽略的目的
                        z = z - 1     
                    key = list(d0.keys())[z]
                    d1[key]= value
                    z = z + 1
                #print(d1)
                with open(file, "a", encoding="utf-8") as f:
                    f.write(str(d1) + "\n")
    Logger.ins().std_logger().info("*****************************************************************************")





