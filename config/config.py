#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置模块
"""
from pathlib import Path
from typing import Dict, Optional
import shutil
from config.config_loader import ConfigLoader

# 项目python代码的log输出目录，根据实际情况修改
PYHON_LOG_FOLDER_PATH = r"./log_data/"
# UserConfig路径，后续需修改为项目工作目录下的指定文件路径
USER_CONFIG_PATH = r"./data/UserConfig.xlsx"
SYSTEM_CONFIG_PATH = r"./data/Config.xlsx"


class Config:
    """项目配置"""

    def __init__(self, config_path: str = None):
        self._loader: Optional[ConfigLoader] = None
        self._paths: Dict[str, Path] = {}
        self._config_source_path: Optional[Path] = None  # 用户配置文件源路径
        self._system_config_source_path: Optional[Path] = None

        if config_path:
            self.load(config_path)

    def load(self, config_path: str, system_config_path: str = None) -> None:
        """加载配置文件"""
        self._config_source_path = Path(config_path)  # 保存源路径
        self._loader = ConfigLoader(config_path)
        system_path = Path(system_config_path) if system_config_path else Path(config_path).with_name('Config.xlsx')
        self._system_config_source_path = None
        if system_config_path or system_path.exists():
            system_loader = ConfigLoader(system_path)
            # 用户配置对同名配置项有优先权，兼容旧版单文件配置。
            for sheet, values in system_loader.config.items():
                self._loader._config.setdefault(sheet, {})
                self._loader._config[sheet] = {**values, **self._loader._config[sheet]}
            # 多行配置（如登录配置）用户页有数据时优先，否则回落到系统配置。
            for sheet, rows in system_loader._rows.items():
                if not self._loader._rows.get(sheet):
                    self._loader._rows[sheet] = rows
            self._system_config_source_path = system_path

        # 读取基础配置
        base_config = self._loader._config.get('基础配置', {})
        self.project_name = base_config.get('项目名称', 'JDE流程')
        self.base_folder = base_config.get('基础路径', './data')
        if not Path(self.base_folder).is_absolute():
            self.base_folder = str(Path(config_path).resolve().parent.parent / self.base_folder)

        # 读取目录配置
        dir_config = self._loader._config.get('目录配置', {})
        self._dir_names = {
            'config': dir_config.get('配置目录', '01_配置'),
            'input': dir_config.get('输入目录', '02_输入'),
            'output': dir_config.get('输出目录', '03_输出'),
            'log': dir_config.get('日志目录', '04_日志'),
            'screenshot': dir_config.get('截图目录', '05_错误截图'),
            'temp': dir_config.get('临时目录', '06_临时文件'),
        }
        self._init_paths()

    def _init_paths(self) -> None:
        """初始化路径"""
        base = Path(self.base_folder) / self.project_name
        for key, name in self._dir_names.items():
            self._paths[key] = base / name

    def get_path(self, path_type: str, filename: str = "") -> Path:
        """获取路径"""
        path = self._paths.get(path_type, self._paths.get('input'))
        if filename:
            return path / filename
        return path

    def get(self, key: str, default=None):
        """获取配置项（支持 sheet.key 格式）

        Args:
            key: 配置键，格式为 "sheet名.配置项"，如 "基础配置.项目名称"
            default: 默认值

        Returns:
            配置值，不存在则返回默认值
        """
        if self._loader:
            return self._loader.get(key, default)
        return default

    def get_rows(self, sheet_name: str) -> list:
        """获取允许多行的配置页（如 登录配置），返回按首行标题组织的字典列表。"""
        if self._loader:
            return self._loader.get_rows(sheet_name)
        return []

    def sheet_names(self) -> list:
        """获取已加载的配置页名称，用于遍历 WechatRobot1、WechatRobot2 等通知配置。"""
        if self._loader:
            return self._loader.sheet_names
        return []

    def set(self, key: str, value) -> None:
        """覆盖配置项，格式为 "sheet名.配置项"；用于按账号切换本次运行参数。"""
        if self._loader:
            self._loader.set(key, value)

    def ensure_directories(self) -> None:
        """确保目录存在，并复制用户配置文件到配置目录"""
        for path in self._paths.values():
            path.mkdir(parents=True, exist_ok=True)

        # 复制用户配置文件到配置目录
        for source in (self._config_source_path, self._system_config_source_path):
            if not source or not source.exists():
                continue
            config_dir = self._paths.get('config')
            if config_dir:
                dest_path = config_dir / source.name
                if not dest_path.exists():
                    shutil.copy2(source, dest_path)


# 全局配置实例
config = Config()