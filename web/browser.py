#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用浏览器管理（Playwright封装）
"""
from __future__ import annotations

from log.loggingConfig import logger

try:
    from playwright.sync_api import sync_playwright, Page, Browser
    HAS_PLAYWRIGHT = True
except ImportError:
    HAS_PLAYWRIGHT = False
    logger.warning("需要安装 playwright: pip install playwright && playwright install")


class BrowserManager:
    """通用浏览器管理

    封装Playwright浏览器操作
    """

    def __init__(self, headless: bool = False, browser_type: str = "chromium"):
        self.playwright = None
        self.browser = None
        self.page = None
        self.headless = headless
        self.browser_type = browser_type

    def start(self) -> bool:
        """启动浏览器"""
        if not HAS_PLAYWRIGHT:
            raise ImportError("需要安装 playwright: pip install playwright && playwright install")

        logger.info(f"启动浏览器: {self.browser_type}, headless={self.headless}")
        self.playwright = sync_playwright().start()

        # 选择浏览器类型
        if self.browser_type == "chromium":
            self.browser = self.playwright.chromium.launch(headless=self.headless)
        elif self.browser_type == "firefox":
            self.browser = self.playwright.firefox.launch(headless=self.headless)
        elif self.browser_type == "webkit":
            self.browser = self.playwright.webkit.launch(headless=self.headless)
        else:
            self.browser = self.playwright.chromium.launch(headless=self.headless)

        self.page = self.browser.new_page()
        logger.info("浏览器启动成功")
        return True

    def close(self) -> None:
        """关闭浏览器"""
        try:
            if self.browser:
                self.browser.close()
                logger.info("浏览器已关闭")
        finally:
            self.browser = None
            self.page = None
            if self.playwright:
                self.playwright.stop()
                self.playwright = None

    def get_page(self) -> Page:
        """获取页面对象"""
        return self.page