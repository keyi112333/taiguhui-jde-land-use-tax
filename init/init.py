#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件描述：初始化流程，读取配置文件、初始化环境等
作者：
"""
from log.loggingConfig import  logger
import psutil
import win32api
import win32con
from config import config


def set_resolution(in_dicConfig):

    # 检查两个key是否存在且值不为空
    width_value = in_dicConfig.get('Resolution_Width')
    height_value = in_dicConfig.get('Resolution_Height')

    if width_value and height_value and width_value != '' and height_value != '':
        try:
            width = int(width_value)
            height = int(height_value)
          
            logger.info(f"使用配置中的分辨率: {width}x{height}")
        except (ValueError, TypeError):
            logger.info("配置中的分辨率值无效，使用默认值1920x1080")
            width, height = 1920, 1080


        devmode = win32api.EnumDisplaySettings(None, win32con.ENUM_CURRENT_SETTINGS)
        devmode.PelsWidth = width
        devmode.PelsHeight = height
        devmode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT

        result = win32api.ChangeDisplaySettings(devmode, win32con.CDS_UPDATEREGISTRY)

        if result == win32con.DISP_CHANGE_SUCCESSFUL:
            logger.info(f"分辨率已更改为: {width}x{height}")
            return True
        else:
            logger.error("分辨率更改失败")
            return False

    else:
        logger.info("width或height配置不存在或为空，不设置")
        return True




def initEnv(list_process:list=[]):
    """
     描述：关闭对应进程名的进程
    :param in_dicConfig，配置信息字典
    :param list_process，需要关闭的进程名列表,如: ["iexplore","chrome","firefox"]

    :return
    """
    logger.info("开始初始化环境")
    # 如果没有传入进程名列表，则从配置文件中获取
    if not list_process:
        str_list_process = str(config.get("基础配置.初始化进程列表", "")).replace("，", ",")
        list_process = str_list_process.split(",")

    list_process = [name.strip() for name in list_process if name.strip()]
        
    #循环关闭对应的进程
    for process_name in list_process:
        for process in psutil.process_iter():
            try:
                # 先获取进程名称，避免后续操作中进程不存在
                proc_name = process.name().lower()
                if process_name.lower() in proc_name:
                    # 检查进程是否仍在运行
                    if process.is_running():
                        process.terminate()
                        logger.info(f"已关闭进程： {process_name} (PID: {process.pid})")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                logger.error(f"关闭进程应用异常：{e}")
    logger.info("完成初始化环境")





