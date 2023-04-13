import re
import os
import shutil
from base.globals import *
from base.systemlogger import Logger

def get_model_id(full_folder_name):
    result=""
    match = re.search(r'140.*[A-Z0-9]', full_folder_name)
    if match:
        result = match.group()
    return result
    
def get_folder_list(folder="data"):
    current_dir = os.getcwd()
    data_dir = os.path.join(current_dir, folder)
    folders = [f for f in os.listdir(data_dir) if os.path.isdir(os.path.join(data_dir, f)) and f not in ["__pycache__", "done"]]
    # Logger.ins().std_logger().info the list of folders
    # Logger.ins().std_logger().info(folders)
    return folders

def rename_folder():
    """
    This function renames all folders in the current directory based on their model ID.
    :param None: This function takes no arguments.
    :return: None
    """
    full_folder_name = get_folder_list()  # Get a list of all folders in the current directory
    for name in full_folder_name:  # Loop through each folder name in the list
        id_folder_name = get_model_id(name)  # Get the new name for the folder based on its model ID
        if id_folder_name == "":
            Logger.ins().std_logger().info(f"not rename {name}")
        else:
            Logger.ins().std_logger().info(f"rename {name} to {id_folder_name} start....")  # Logger.ins().std_logger().info a message indicating that the renaming process has started
            os.rename(name, id_folder_name)  # Rename the folder using the new name

        
        
def backup_folder(folder_path, backup_path, move=False):
    """
    Move or copy a folder to a backup location using shutil.

    Args:
        folder_path (str): The path to the folder to be backed up.
        backup_path (str): The path to the backup location.
        move (bool, optional): Whether to move the folder instead of copying it. Defaults to False.
    """
    if move:
        shutil.move(folder_path, backup_path)
    else:
        shutil.copytree(folder_path, backup_path)
    str = "move" if move else "copy"
    Logger.ins().std_logger().info(f"[INFO] {str} {folder_path} to {backup_path} done!")
        
def choose_text_content(full_folder_name):
    if "140" in full_folder_name:
        if "处理器" in full_folder_name:
            return chuliqi_140
        if "离散量" in full_folder_name:
            return lisanliang_shuchu_140
        if "以太网" in full_folder_name:
            return yitaiwang_140
        if "机架背板" in full_folder_name:
            return jijia_beiban_140
        if "光纤头端适配器" in full_folder_name:
            return guangxin_shipeiqi_140
        return common_140


def set_foreground_window(pattern = ".*Google Chrome.*"):
    import win32gui  # 导入 Windows API 模块
    import win32con  # 导入 Windows API 模块
    import re  # 导入正则表达式模块
    """
    Need administrator permission
    """
    prog = re.compile(pattern)  # 编译正则表达式
    hwnd_list = []  # 窗口句柄列表
    win32gui.EnumWindows(lambda hwnd, param: param.append(hwnd) if prog.match(win32gui.GetWindowText(hwnd)) else None, hwnd_list)  # 枚举所有窗口，将匹配的窗口句柄添加到列表中
    if hwnd_list:
        window = hwnd_list[0]  # 取第一个匹配的窗口句柄
        win32gui.SetForegroundWindow(window)  # 将窗口置于前台
        win32gui.ShowWindow(window, win32con.SW_RESTORE)  # 还原窗口
        win32gui.ShowWindow(window, win32con.SW_MAXIMIZE)  # 最大化窗口
    else:
        Logger.ins().std_logger().info("没有找到匹配的窗口")  # 如果没有找到匹配的窗口，则输出提示信息
    

def get_files_list(path='obs', keys='xlsx'):
    """
    列出指定目录下所有指定类型的文件

    :param path: 目录路径，默认为'obs'
    :param keys: 文件类型，默认为'xlsx'
    return list
    """
    Logger.ins().std_logger().info(f"[INFO] 获取目录{path}下的含有关键字{keys}的所有文件.")
    files_list = []
    pattern = re.compile(fr".*{keys}.*")
    for root, dirs, files in os.walk(path):
        for file in files:
            if pattern.match(file):
                # 输出文件路径
                # Logger.ins().std_logger().info(os.path.join(root, file))
                files_list.append(os.path.join(root, file))
    return files_list

def read_excel_column(file_name, column_num):
    import xlrd
    # 打开excel文件
    workbook = xlrd.open_workbook(file_name)
    # 根据sheet索引获取sheet内容
    sheet = workbook.sheet_by_index(0)
    # 获取行数
    rows = sheet.nrows
    # 获取列数，尽管没用到
    cols = sheet.ncols
    all_content = []
    cols = sheet.col_values(column_num) # 获取第一列内容
    for i in range(rows):
        all_content.append(cols[i])
    return all_content




def append_data_to_worksheet(file_path, sheet_name, data):
    import openpyxl
    # open workbook
    workbook = openpyxl.load_workbook(file_path)
    # get worksheet object
    worksheet = workbook[sheet_name]
    # append data to worksheet
    for row in data:
        worksheet.append(row)
    # save workbook
    workbook.save(file_path)