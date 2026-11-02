#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置模块
"""
from .config import Config, config
from .config_loader import ConfigLoader, ConfigError

__all__ = ['Config', 'config', 'ConfigLoader', 'ConfigError']