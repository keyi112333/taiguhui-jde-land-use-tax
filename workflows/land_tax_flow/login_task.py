"""JDE 登录子任务，登录后保留浏览器供后续文件处理使用。"""
import os

from core.task_base import TaskBase
from utils.credentials import read_windows_credential
from web.JDE.jde_function import jde_login


class JDELoginTask(TaskBase):
    def __init__(self, context):
        super().__init__("JDE登录", context)

    def execute(self):
        config = self.context.config
        username, password = read_windows_credential(
            config.get("JDE配置.jde_credential_name", "")
        )
        browser = self.context.get_browser("JDE")
        if browser is None:
            raise RuntimeError("JDE浏览器尚未初始化")
        jde_login(
            in_page=browser.get_page(),
            jde_url=os.environ.get("JDE_URL", config.get("JDE配置.jde_url", "")),
            jde_environment=os.environ.get(
                "JDE_ENVIRONMENT", config.get("JDE配置.jde_environment", "")
            ),
            jde_role=os.environ.get("JDE_ROLE", config.get("JDE配置.jde_role", "")),
            jde_user=username,
            jde_pass=password,
        )
        self.context.set("jde_logged_in", True)
        return True
