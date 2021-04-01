import fnmatch
import os
from pathlib import Path

import pandas


def slice_table():
    while True:
        file_path = input('请输入要拆分的文件路径：')
        # 判断是否输入
        if not file_path:
            continue
        # 判断文件格式是否正确
        if not file_path.endswith('.xlsx'):
            print('文件需要以 .xlsx 结尾.')
            continue
        # 判断文件是否存在
        if not os.path.exists(file_path):
            print(f'文件 {file_path} 不存在.')
            continue
        break

    while True:
        sheet_name = input('请输入要拆分sheet的序号 (默认为 0)：')
        # 判断是否输入
        if not sheet_name:
            sheet_name = 0
        try:
            wb = pandas.read_excel(file_path, sheet_name=int(sheet_name), engine='openpyxl')
            wb.style.to_excel('./2.xlsx', index=False, encoding='openpyxl')
            break
        except:
            print(f'sheet: {sheet_name} 不存在')

    prompt_str = ''
    for _index, _column in enumerate(wb.columns):
        prompt_str += f'[{_index}]{_column}'
        if _index < len(wb.columns):
            prompt_str += ", "

    while True:
        column_index = input(f'请输入要拆分列的序号 (默认为 0): {prompt_str}')
        if not column_index:
            column_index = 0
        try:
            rets = list(wb.groupby(wb.columns[int(column_index)]))
            break
        except:
            print(f'column: {column_index} 不存在')

    while True:
        save_path = input(f'请输入文件保存路径 (默认文件夹: {os.getcwd()}):')
        try:
            if not save_path:
                save_path = os.getcwd()
            if not os.path.exists(save_path):
                os.makedirs(save_path)
            for book in rets:
                book[1].style.data.to_excel(Path(save_path, f'{str(book[0])}.xlsx'), index=False, engine='openpyxl')
            print(f'文件已保存至 {os.path.realpath(save_path)} 文件夹')
            break
        except:
            print(f'文件路径 {save_path} 错误')


def merge():
    file_list = None
    while True:
        file_path = input('请输入要合并的文件夹路径：')
        if not os.path.exists(file_path):
            print(f'文件路径 {file_path} 不存在.')
            continue
        file_list = fnmatch.filter(os.listdir(file_path), '*.xlsx')
        if len(file_list) == 0:
            print(f'未找到需要合并的文件.')
            continue
        break

    while True:
        save_path = input(f'请输入合并文件保存路径 (默认为当前目录 {os.getcwd()}): ')
        try:
            if not save_path:
                save_path = os.getcwd()
            if not os.path.exists(save_path):
                os.makedirs(save_path)
            break
        except:
            print(f'输入路径不合法 {save_path}')

    try:
        excel_file_list = []
        for _file in file_list:
            f = pandas.read_excel(Path(file_path, _file))
            excel_file_list.append(f)
        rets = pandas.concat(excel_file_list)
        rets.to_excel(Path(save_path, "合并.xlsx"), encoding='utf-8', index=False)
        print(f'文件已保存至 {os.path.realpath(save_path)} 文件夹')
    except Exception as e:
        print(e)
        print('文件合并失败.')


if __name__ == '__main__':
    while True:
        operation = input('请选择要执行的操作 (目前仅支持xlsx): [0]拆分 (默认), [1]合并')
        if not operation:
            operation = '0'
        if operation == '0' or operation == '1':
            break
        print(f'不支持的操作 {operation}')

    if operation == '0':
        slice_table()
    else:
        merge()

    input('按Enter键结束程序.')
