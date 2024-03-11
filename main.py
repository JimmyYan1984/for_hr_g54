# -*- coding: utf-8 -*-
import os
import sys

from user_function.for_hr import HrTool

if __name__ == '__main__':
    hr_tool = HrTool()
    # 获取可执行文件所在的路径
    exe_path = os.path.dirname(sys.argv[0])
    path = os.path.abspath(exe_path)
    hr_tool.check_g54(path)
    # 获取编译成可执行文件后的路径


# find all images without alternate text
# and give them a red border
