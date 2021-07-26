from tkinter import Tk as _Tk
from tkinter.filedialog import askopenfilename as _askopenfilename
from tkinter.filedialog import asksaveasfilename as _asksaveasfilename
from tkinter.filedialog import askdirectory as _askdirectory

_Tk().withdraw()

def check_path(path:str):
    """检查路径
    若路径为空则引发 ValueError('操作取消')
    path: 要检查的路径
    """
    if path == '':raise ValueError('操作取消')

def open_file(*args, **kwargs):
    """弹出打开单个文件的对话框
    若路径为空或取消则引发 ValueError('操作取消')
    所有参数传递给 tkinter._askopenfilename
    """
    path = _askopenfilename(*args, **kwargs)
    check_path(path)
    return path

def save_file(*args, **kwargs):
    """弹出保存单个文件的对话框
    若路径为空或取消则引发 ValueError('操作取消')
    所有参数传递给 tkinter._asksaveasfilename
    """
    path = _asksaveasfilename(*args, **kwargs)
    check_path(path)
    return path

def open_folder(*args, **kwargs):
    """弹出打开单个文件夹的对话框
    若路径为空或取消则引发 ValueError('操作取消')
    所有参数传递给 tkinter._askdirectory
    """
    path = _askdirectory(*args, **kwargs)
    check_path(path)
    return path