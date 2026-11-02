#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
任务上下文
"""
from typing import Any, Dict


class TaskContext:
    """任务上下文，用于在任务间共享状态

    支持存储：
    - 任意数据（通过 set/get 方法）
    - 多个系统的浏览器实例（browsers字典）
    - 多个客户端应用实例（clients字典）
    - 配置实例
    """

    def __init__(self):
        self._data: Dict[str, Any] = {}
        self.browsers: Dict[str, Any] = {}  # 按系统名称保存浏览器管理器
        self.clients: Dict[str, Any] = {}  # 按应用名称保存客户端实例
        self.config = None  # 配置实例

    def set(self, key: str, value: Any) -> None:
        """存储数据"""
        self._data[key] = value

    def get(self, key: str, default: Any = None) -> Any:
        """获取数据"""
        return self._data.get(key, default)

    def set_browser(self, system_name: str, browser: Any) -> None:
        """存储浏览器实例"""
        self.browsers[system_name] = browser

    def get_browser(self, system_name: str) -> Any:
        """获取浏览器实例"""
        return self.browsers.get(system_name)

    def set_client(self, app_name: str, client: Any) -> None:
        """存储客户端应用实例"""
        self.clients[app_name] = client

    def get_client(self, app_name: str) -> Any:
        """获取客户端应用实例"""
        return self.clients.get(app_name)

    def clear(self) -> None:
        """清空所有数据"""
        self._data.clear()

    def close_all_browsers(self) -> None:
        """关闭所有浏览器"""
        for system_name, browser in self.browsers.items():
            try:
                if browser and hasattr(browser, 'close'):
                    browser.close()
            except Exception as e:
                pass  # 忽略关闭错误
        self.browsers.clear()

    def close_all_clients(self) -> None:
        """关闭所有客户端应用"""
        for app_name, client in self.clients.items():
            try:
                if client and hasattr(client, 'close'):
                    client.close()
            except Exception as e:
                pass  # 忽略关闭错误
        self.clients.clear()
