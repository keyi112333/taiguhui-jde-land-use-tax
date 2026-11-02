#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基础任务抽象类
"""
import time
from abc import ABC, abstractmethod

from log.loggingConfig import logger
from utils.helpers import format_time


class TaskBase(ABC):
    """基础任务抽象类

    所有任务（主任务、子任务）都继承此类
    """

    def __init__(self, name: str, context=None):
        self.name = name
        self.context = context  # 任务上下文，用于共享状态
        self.status = "pending"
        self.result = None
        self.start_time = None
        self.end_time = None
        self.elapsed_time = None

    @abstractmethod
    def execute(self) -> bool:
        """执行任务（子类实现具体逻辑）"""
        pass

    def run(self) -> bool:
        """运行任务（带日志和状态管理）"""
        logger.info(f"开始任务: {self.name}")
        self.status = "running"
        self.start_time = time.time()
        try:
            self.result = self.execute()
            self.status = "completed"
            self.end_time = time.time()
            self.elapsed_time = self.end_time - self.start_time
            logger.info(f"完成任务: {self.name}, 耗时: {format_time(self.elapsed_time)}")
            return True
        except Exception as e:
            self.status = "failed"
            self.end_time = time.time()
            self.elapsed_time = self.end_time - self.start_time
            logger.error(f"任务失败: {self.name} - {e}, 耗时: {format_time(self.elapsed_time)}")
            return False

    def get_status(self) -> str:
        """获取任务状态"""
        return self.status