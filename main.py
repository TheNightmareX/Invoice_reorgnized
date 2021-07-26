from database import DataBase
from statement import Statement
import sys
from sys import exit
from os import system as cmd
from os.path import exists
from data import base_dirs
import stdout
from typing import Callable


if __name__ == '__main__':
    database = DataBase(base_dirs.src + '/data/database.xlsx')
    if len(sys.argv) > 1: # 带参启动
        path = sys.argv[1]
        if exists(path): # 参数为存在的路径
            if path.endswith('.txt'):
                stdout.print_info('模式:导入编码数据')
                database.update_from_txt(path)
                database.flush()
            elif path.endswith('.xls') or path.endswith('.xlsx'):
                stdout.print_info('模式:创建XML单据')
                statement = Statement(database, path)
                stdout.print_table(statement.details.data)
                statement.details.lock()
                statement.details.merge()
                statement.convert_to_xml()
            else:
                cmd('color 4')
                stdout.print_err('不支持的文件类型')
        else:
            pass
    else: # 无参启动
        cmd('color 4')
        stdout.print_err('缺少参数')
stdout.pause()
exit()