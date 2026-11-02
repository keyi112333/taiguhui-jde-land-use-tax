#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件描述：初始化日志配置，设置logger对象，指定输出日志文件等
作者：
"""
import logging
from datetime import datetime
import os, sys
from config.config import PYHON_LOG_FOLDER_PATH
import socket


# current_path = os.getcwd()
# logFilePath = os.path.join(current_path,"log", datetime.now().strftime("%Y-%m-%d")+'_logger.log')

# 获取今天的日期，用于日志文件命名
today = datetime.now().strftime("%Y%m%d")

# 定义输出日志的目录，没有自动创建
log_directory = os.path.expanduser(PYHON_LOG_FOLDER_PATH)

# 获取机器名
str_machine_name = socket.gethostname()

if not os.path.exists(log_directory):
    os.makedirs(log_directory)

# 定义输出日志文件的路径，没有自动创建
logFilePath = os.path.join(log_directory, f"{today}_{str_machine_name}.log")
if not os.path.exists(logFilePath):
    with open(logFilePath, "w") as f:
        pass


logger = logging.getLogger("logger")
logger.setLevel(logging.DEBUG)
logger.propagate = False

if not logger.handlers:
    # 创建一个handler，用于输出到控制台
    console_handler = logging.StreamHandler(stream=sys.stderr)
    # 创建一个handler，用于输出到文件
    file_handler = logging.FileHandler(logFilePath, encoding="utf-8")
    # 设置日志格式
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    # 设置handler的格式
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)

    # 将控制台handler添加到logger中
    logger.addHandler(console_handler)

    # 将文件handler添加到logger中
    logger.addHandler(file_handler)

# 使用logger对象记录日志
# logger.debug('This is a debug message')
# logger.info('This is an info message')
# logger.warning('This is a warning message')
# logger.error('This is an error message')
# logger.critical('This is a critical message')
