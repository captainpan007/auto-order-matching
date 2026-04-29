# -*- coding: utf-8 -*-
"""
对账系统配置文件
部署到倩茹电脑时，只需修改这里的路径
"""

import os

# 倩茹放单据的根目录（每个批次文件夹都在这里面）
BASE_DATA_DIR = r"F:\claude开发项目\atutoordermatching"   # 开发环境
# BASE_DATA_DIR = r"D:\对账文件"   # 部署到倩茹电脑时改用此路径

# 程序所在目录
BASE_APP_DIR = os.path.dirname(os.path.abspath(__file__))

# 对账结果输出目录
OUTPUT_DIR = os.path.join(BASE_APP_DIR, "对账结果")
