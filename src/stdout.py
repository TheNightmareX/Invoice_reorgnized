from os import system as _cmd
from threading import Thread as _Thread
from time import sleep as _sleep 
from time import strftime as _strftime
from easy_table import EasyTable as _EasyTable
from typing import List as _List
from typing import Dict as _Dict


def print_ok(info:str, *args, **kwargs):
    print(f"✓ >> {info}", *args, **kwargs)

def print_err(info:str, *args, **kwargs):
    print(f"! >> {info}", *args, **kwargs)

def print_info(info:str, *args, **kwargs):
    print(f">>>> {info}", *args, **kwargs)

def ask_input(info:str, int_only=False, num_only=False):
    user_input = input(f"? >> {info}")
    return user_input

def pause(info:str=''):
    """显示信息后暂停 用户按下任意键继续
    info 要显示的信息 
    """
    print(info)
    _cmd('pause > nul')

class ConsoleAnimationBase:
    """控制台动画的基类
    继承这个类必须定义 _main_thread 函数
    调用 start 或 __enter__ 方法时会启动 _main_thread 线程
    _main_thread 函数:
        被调用后设置 self.running = True
        self._stop == True 时停止动画并清理资源 并设置 self.running = False
    """
    def __init__(self):
        self._stop = False
        self.running = False
        
    def start(self):
        """启动动画
        """
        _Thread(target=self._main_thread).start()

    def stop(self):
        """停止动画
        """
        self._stop = True
        while self.running == True:
            _sleep(0.1)
        self._stop = False

    def __enter__(self):
        self.start()

    def __exit__(self, exc_type, exc_value, traceback):
        self.stop()

def print_table(data:_List[_Dict]):
    """简易封装_EasyTable
    """
    table = _EasyTable()
    table.setCorners('+', '+', '+', '+')
    table.setOuterStructure('|', '-'),
    table.setInnerStructure('|', '-', '+')
    table.setData(data)
    table.displayTable()

def date_input(info:str, spliter:str='-') -> str:
    """只接受日期输入
    """
    def is_date(s:str, spliter:str='-') -> bool:
        """判断 string 是否为日期
        """
        if s.count(spliter) == 2:
            l:_List[str] = s.split(spliter)
            if (
                len(l[0]) == 4 and l[0].isnumeric() and
                len(l[1]) == 2 and l[1].isnumeric() and
                len(l[2]) == 2 and l[2].isnumeric()
            ):
                return True
        return False

    while True:
        user_input:str = ask_input(info)
        if is_date(user_input):return user_input
        print_err(f"“{user_input}”不是日期")

def limited_input_index(info:str, choices:_List[str]) -> str:
    """列出 choices 里的所有值并附上序号 只接受范围内的序号 返回序号以及对应的条目的值
    """
    def show_choices(choices:_List[str]) -> None:
        """列出 choices 里的所有值并附上序号
        """
        print_table([{'序号':idx, '选项':value} for idx, value in enumerate(choices)])

    show_choices(choices)
    while True:
        user_input = ask_input(info)
        if user_input.isnumeric() and int(user_input) in range(len(choices)):return (user_input, choices[int(user_input)])
        print_err(f"“{user_input}”超出范围或不是序号")


def limited_input(info:str, choices:_List[str], strict=False) -> str:
    """只接受 choices 里的值
    当 strict == False 且用户输入的值不与 choices 中的任何一个值匹配时, 找出 choices 中相似的条目, 调用 limited_input_index 显示
    当 strict == True 询问直到用户输入choices 中的值
    """
    while True:
        user_input = ask_input(info)
        if user_input in choices:return user_input
        if not strict:
            new_choices = [v for v in choices if user_input in v]
            if len(new_choices) > 0:
                return limited_input_index(f'找到含有“{user_input}”的这些条目,请输入你选择的条目的序号:', new_choices)[1]
            else:
                return limited_input_index(f'没有找到含有“{user_input}”的条目,已列出全部条目,请输入你选择的条目的序号:', choices)[1]
        else:
            print_err(f"{user_input}不在{choices}中")

def number_input(info:str, int_only:bool=False) -> float:
    """只接受 数值 
    若 int_only == True 则只接受整数
    """
    while True:
        user_input = ask_input(info)
        if int_only:
            if user_input.isnumeric():return float(user_input)
        else:
            if user_input.replace('.', '').isnumeric() and user_input[0] != '.' and user_input[-1] != '.': return float(user_input)
        print_err(f"“{user_input}不是{'整数' if int_only else '数字'}")

class RotatingStick(ConsoleAnimationBase):
    """打印一根旋转的小棒

    支持with语句
    """
    def _main_thread(self):
        self.running = True
        TO_PRINT = ['|', '/', '-', '\\']
        index = 0
        print(' ', end='')
        while not self._stop:
            print(f"\b{TO_PRINT[index]}", end='', flush=True)
            index += 1
            if index == 4:index = 0
            _sleep(0.1)
        print('\b \b', end='')
        self.running = False


if __name__ == '__main__':
    pass