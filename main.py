# -*- coding: utf-8 -*-
import os
from user_function.for_hr import HrTool

if __name__ == '__main__':
    hr_tool = HrTool()
    # 获取当前目录的绝对路径
    path = os.path.abspath(os.path.dirname(__file__))
    hr_tool.check_g54(path)

# find all images without alternate text
# and give them a red border
