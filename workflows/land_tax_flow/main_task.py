"""按「登录配置」逐行登录，处理共享盘“待处理”目录中的土地使用税文件。"""

import re
import shutil
from datetime import datetime
from pathlib import Path

from config.config_loader import ConfigError
from core.context import TaskContext
from core.task_base import TaskBase
from log.loggingConfig import logger
from utils.credentials import CredentialNotFoundError, read_windows_credential
from web.browser import BrowserManager
from workflows.land_tax_flow.land_tax import LandUseTaxFlow
from workflows.land_tax_flow.login_task import JDELoginTask
from workflows.land_tax_flow.run_report import (
    append_run_result,
    build_notification_message,
    notify_wechat_robots,
    result_file_path,
)

RUN_SHEET = "登录配置"
REQUIRED_COLUMNS = ("window凭据", "凭据账号", "共享盘路径")
PENDING_DIR_KEY = "土地使用税配置.待处理目录"
PROCESSED_DIR_KEY = "土地使用税配置.已处理目录"
SUCCESS_DIR = "成功"
FAILED_DIR = "失败"
FILE_PATTERN = re.compile(
    r"^(?P<company>.+?)-(?P<period>\d{6})-土地使用税\.(?P<ext>xlsx|xls)$",
    re.IGNORECASE,
)


class SkipRun(Exception):
    """当前行凭据未配置，跳过该行。"""


def parse_shared_path(value):
    """从共享盘路径末段目录（如“太古汇 2501”）解析公司简称与公司代码。"""
    path = Path(str(value).strip())
    tokens = path.name.split()
    if len(tokens) != 2 or not tokens[1].isdigit():
        raise ConfigError(
            f"共享盘路径末段需为“公司简称 公司代码”（如 太古汇 2501），当前：{path.name}"
        )
    return path, tokens[0], tokens[1]


def read_headless(value):
    if isinstance(value, str):
        text = value.strip().lower()
        if text not in ("true", "false", "1", "0", "yes", "no"):
            raise ConfigError("JDE配置.headless 必须为 true 或 false")
        return text in ("true", "1", "yes")
    return bool(value)


def find_pending_files(directory: Path):
    """扫描待处理目录，返回 (文件路径, 公司名, 期间) 列表。"""
    files = []
    seen = {}
    for path in sorted(directory.iterdir()):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        match = FILE_PATTERN.match(path.name)
        if match is None:
            continue
        period = match.group("period")
        month = int(period[4:6])
        if not 1 <= month <= 12:
            raise ConfigError(f"文件名中的月份不正确：{path.name}")
        company = match.group("company").strip()
        key = (company.lower(), period)
        if key in seen:
            raise ConfigError(
                f"待处理目录中 {company}-{period} 同时存在多个土地使用税文件："
                f"{seen[key].name}、{path.name}"
            )
        seen[key] = path
        files.append((path, company, period))
    return files


class JDEMainTask(TaskBase):
    def __init__(self, config):
        context = TaskContext()
        context.config = config
        super().__init__("JDE土地使用税流程", context)

    def execute(self) -> bool:
        config = self.context.config
        if not config.get("JDE配置.jde_url", ""):
            raise ConfigError("请配置 JDE配置.jde_url")
        runs = config.get_rows(RUN_SHEET)
        if not runs:
            raise ConfigError(
                f"{RUN_SHEET} 页没有数据行，请填写 "
                f"{'、'.join(REQUIRED_COLUMNS)}；每行会处理该共享盘下的待处理文件"
            )
        failures = []
        for index, run in enumerate(runs, 1):
            label = f"[{index}/{len(runs)}] {run.get('共享盘路径')}"
            logger.info(f"开始执行 {label}")
            try:
                self.run_one(run)
                logger.info(f"{label} 执行完成")
            except SkipRun as exc:
                logger.warning(f"{label} 跳过：{exc}")
            except Exception as exc:
                failures.append(f"{label}：{exc}")
                logger.exception(f"{label} 执行失败：{exc}")
        if failures:
            raise ConfigError("以下账号执行失败：" + "；".join(failures))
        return True

    def prepare(self, run):
        """校验本行配置并把凭据、公司代码写入土地使用税配置。"""
        config = self.context.config
        values = {name: str(run.get(name) or "").strip() for name in REQUIRED_COLUMNS}
        missing = [name for name, value in values.items() if not value]
        if missing:
            raise ConfigError(f"{RUN_SHEET} 缺少：{', '.join(missing)}")

        shared_path, short_name, code = parse_shared_path(values["共享盘路径"])
        if not shared_path.is_dir():
            raise ConfigError(f"共享盘路径不存在：{shared_path}")

        pending = Path(str(config.get(PENDING_DIR_KEY, "待处理") or "待处理").strip())
        if not pending.is_absolute():
            pending = shared_path / pending
        if not pending.is_dir():
            raise ConfigError(f"共享盘缺少待处理目录：{pending}")

        processed = Path(str(config.get(PROCESSED_DIR_KEY, "已处理") or "已处理").strip())
        if not processed.is_absolute():
            processed = shared_path / processed

        try:
            username, _ = read_windows_credential(values["window凭据"])
        except CredentialNotFoundError:
            raise SkipRun(f"Windows 凭据 {values['window凭据']} 不存在，已跳过") from None
        if username.split("\\")[-1].strip().lower() != values["凭据账号"].lower():
            logger.warning(
                f"Windows 凭据 {values['window凭据']} 的用户名 {username} "
                f"与「凭据账号」{values['凭据账号']} 不一致，请确认数据归属"
            )

        company = code.zfill(5)
        overrides = {
            "JDE配置.jde_credential_name": values["window凭据"],
            "土地使用税配置.分录公司": company,
        }
        for key, value in overrides.items():
            config.set(key, value)
        logger.info(
            f"运行参数：凭据={values['window凭据']} 账号={values['凭据账号']} "
            f"公司={company} 共享盘={shared_path} 待处理={pending} 已处理={processed}"
        )
        return {
            "credential": values["window凭据"],
            "company": company,
            "short_name": short_name,
            "shared_path": shared_path,
            "pending": pending,
            "success_dir": processed / SUCCESS_DIR,
            "failed_dir": processed / FAILED_DIR,
        }

    def _record(self, record):
        """把单个文件的运行结果写入年度结果表；写入失败只记日志，不影响主流程。"""
        try:
            append_run_result(self.context.config, record)
        except Exception as exc:
            logger.warning(f"运行结果表写入失败：{exc}")

    def _notify(self, info, records, result_path):
        """按共享盘维度汇总本次结果并发送企业微信通知。"""
        success = sum(1 for record in records if record["运行结果"] == "成功")
        periods = "、".join(sorted({str(record["会计期间"]) for record in records}))
        message = build_notification_message(
            info["short_name"],
            periods,
            len(records),
            success,
            len(records) - success,
            result_path.name,
        )
        notify_wechat_robots(self.context.config, message, result_path)

    def _move(self, path: Path, target_dir: Path, failures):
        """把已处理文件移入结果目录；重名时追加序号，避免覆盖历史文件。"""
        try:
            target_dir.mkdir(parents=True, exist_ok=True)
            target = target_dir / path.name
            index = 1
            while target.exists():
                target = target_dir / f"{path.stem}_{index}{path.suffix}"
                index += 1
            shutil.move(str(path), str(target))
            logger.info(f"已移动文件：{path.name} -> {target}")
        except Exception as exc:
            failures.append(f"{path.name} 移动失败：{exc}")
            logger.exception(f"{path.name} 移动失败：{exc}")

    def _screenshot(self, browser, info, company="", period=""):
        page = browser.get_page() if browser else None
        if page is None:
            return
        suffix = "_".join(filter(None, (company or info.get("short_name"), period)))
        name = f"jde_failure_{suffix}.png" if suffix else "jde_failure.png"
        try:
            page.screenshot(path=str(self.context.config.get_path("screenshot", name)))
        except Exception:
            logger.warning("JDE错误截图保存失败")

    def run_one(self, run):
        info = self.prepare(run)
        files = find_pending_files(info["pending"])
        if not files:
            logger.warning(f"{info['pending']} 中没有土地使用税文件，跳过")
            return

        config = self.context.config
        result_path = result_file_path(config)
        records = []
        browser = BrowserManager(headless=read_headless(config.get("JDE配置.headless", False)))
        self.context.set_browser("JDE", browser)
        failures = []
        try:
            browser.start()
            try:
                if not JDELoginTask(self.context).run():
                    raise RuntimeError("JDE 登录失败")
            except Exception:
                self._screenshot(browser, info)
                raise

            for path, company, period in files:
                logger.info(f"开始处理：{path.name}（公司={company} 期间={period}）")
                self.context.set("land_tax_status", None)
                self.context.set("land_tax_summary", None)
                self.context.set("jde_journal_batch", None)
                started = datetime.now()
                try:
                    LandUseTaxFlow(self.context).run(path, company, period)
                    batch = str(self.context.get("jde_journal_batch") or "")
                    logger.info(f"{path.name} 处理完成，JDE 批号：{batch}")
                    self._move(path, info["success_dir"], failures)
                    outcome, error = "成功", ""
                except Exception as exc:
                    failures.append(f"{path.name}：{exc}")
                    logger.exception(f"{path.name} 处理失败：{exc}")
                    self._screenshot(browser, info, company, period)
                    self._move(path, info["failed_dir"], failures)
                    batch, outcome, error = "", "失败", str(exc)
                records.append({
                    "会计期间": period,
                    "公司": company,
                    "分录批号": batch,
                    "运行结果": outcome,
                    "异常情况": error,
                    "运行开始时间": started,
                    "运行结束时间": datetime.now(),
                })
                self._record(records[-1])
        finally:
            self.context.close_all_browsers()

        if records:
            self._notify(info, records, result_path)
        if failures:
            raise ConfigError("以下土地使用税文件处理失败：" + "；".join(failures))
