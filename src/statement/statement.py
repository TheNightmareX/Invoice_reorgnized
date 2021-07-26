from ._statement import Statement as _Statement
from database import DataBase as _DataBase
import stdout as _stdout
from os import system as _cmd



class Statement(_Statement):
    """Statement的用户接口
    """
    class Details(_Statement.Details):
        def lock(self):
            """锁定indexes列出的商品行
            """
            while True:
                user_input = _stdout.ask_input('锁定(以“,”分割):')
                if user_input == '':return # 不锁定
                indexes = user_input.split(',')
                try:
                    super().lock(indexes)
                except ValueError as err:
                    _stdout.print_err(err)
                    continue
                else:
                    break
                
        def merge(self):
            """合并相同商品的商品行 不会合并锁定的商品行
            """
            while True:
                user_input = _stdout.limited_input('合并?[y/n]:', ['y', 'n'], True)
                if user_input == 'y':
                    super().merge()
                    break
                else:
                    break

    def __init__(self, database:_DataBase, path:str):
        """读取数据
        """
        _stdout.print_info('解析数据...', end='')
        try:
            with _stdout.RotatingStick():
                super().__init__(database, path)
        except Exception as err:
            print()
            _cmd('color 4')
            print(err)
            _stdout.print_err('解析失败')
            _stdout.pause()
            _exit()
        else:
            print()
            _stdout.print_ok('解析完成')

    def convert_to_xml(self):
        """将self.details.data分组 用分组后的数据创建符合开票软件接口规范的XML单据
        """
        _stdout.print_info('转换为XML单据...', end='')
        try:
            with _stdout.RotatingStick():
                groups = super().convert_to_xml()
        except Exception as err:
            print()
            _cmd('color 4')
            print(err)
            _stdout.print_err('转换失败')
            _stdout.pause()
            _exit()
        else:
            print()
            for idx, group in enumerate(groups):
                print(f"第{idx + 1}张:{group['add_up']}元")
                _stdout.print_table(group['details'])
            _stdout.print_ok(f'转换成功 单据已存放到{self.path}.xml')
        