import sys as _sys
import os
from os.path import dirname as _dirname
from os.path import isabs as _isabs
from os.path import abspath as _abspath


cwd = _abspath(os.getcwd())
src = _dirname(_abspath(_sys.argv[0])) # 入口所在的目录 双击启动时sys.argv[0]为绝对路径 命令行启动时为相对路径
res = getattr(_sys, '_MEIPASS', src) # 资源文件所在目录 用PyInstaller打包后启动时资源文件将会在%TEMP%/_MEIxxxx 否则与src相同