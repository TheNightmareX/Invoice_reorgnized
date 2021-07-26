from ._database import DataBase as _DataBase
from os import system as _cmd
from os import makedirs as _makedirs
from os.path import abspath as _abspath
from os.path import dirname as _dirname
import stdout as _stdout
from sys import exit as _exit
from data import base_dirs


class DataBase(_DataBase):
    """DataBase的用户接口
    """
    def __init__(self, path:str):
        """检查路径
        """
        path = _abspath(path)
        try:
            super().__init__(path)
        except FileNotFoundError as err:
            _cmd('color 4')
            _stdout.print_err(f'数据库丢失 (找不到 {path} )')
            _makedirs(_dirname(path), exist_ok=True)
            with open(base_dirs.res + '/data/database.frame', 'rb') as src:
                with open(path, 'wb') as tgt:
                    tgt.writelines(src.readlines())
            _stdout.print_info(f'已在此位置新建数据库模板')
            _stdout.pause()
            _exit()

    def load(self):
        """加载数据库
        """
        _stdout.print_info('加载数据...', end='')
        try:
            with _stdout.RotatingStick():
                super().load()
        except Exception as err:
            print()
            _cmd('color 4')
            print(err)
            _stdout.print_err('加载数据库失败')
            _stdout.pause()
            _exit()
        else:
            print()
            _stdout.print_ok('加载完成')

    def update_from_txt(self, path:str):
        """从开票软件导出的txt文件加载数据并用这些数据更新
        """
        _stdout.print_info('导入...', end='')
        try:
            with _stdout.RotatingStick():
                super().update_from_txt(path)
        except Exception as err:
            print()
            _cmd('color 4')
            print(err)
            _stdout.print_err('导入编码数据失败')
            _stdout.pause()
            _exit()
        else:
            print()
            _stdout.print_ok('导入成功')