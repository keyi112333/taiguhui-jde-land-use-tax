"""JDE 登录与日记帐分录录入。"""

import functools
import re
import time

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from log.loggingConfig import logger

MENU_IFRAME = 'iframe[name="e1menuAppIframe"]'
# 日记帐分录录入前需要选中的分录方案（JDE #gtab0_1 下拉值）。
JOURNAL_SCHEME_OPTION = "8d5afde055f149ea973fb07587dafa28"
# 粘贴前必须存在的列及顺序（TSV 按该顺序整块粘贴）。
JOURNAL_COLUMNS = ("帐号", "借项金额", "贷项金额", "备注")


def retry_on_failure(max_retries=3, delay=5):
    """函数失败后重试，最后一次仍失败则抛出原异常。"""

    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception:
                    if attempt < max_retries - 1:
                        logger.warning(f"{func.__name__} 执行失败，准备重试")
                        logger.info(f"等待 {delay} 秒后进行第 {attempt + 2} 次尝试...")
                        time.sleep(delay)
                    else:
                        logger.info(f"{func.__name__} 最终执行失败，已重试 {max_retries} 次")
                        raise
            return None

        return wrapper

    return decorator


@retry_on_failure(max_retries=3, delay=5)
def jde_login(in_page, jde_url, jde_environment, jde_role, jde_user, jde_pass):
    """登录 JDE；已登录时直接返回。"""
    in_page.goto(jde_url)

    element_exists = in_page.locator("#drop_mainmenu")
    if element_exists.is_visible(timeout=5000):
        logger.info("登录成功")
        return

    element_exists = in_page.get_by_role("textbox", name="Environment")
    if not element_exists.is_visible(timeout=2000):
        in_page.get_by_role("link", name="Details").click()
        in_page.wait_for_timeout(500)

    if jde_environment:
        in_page.get_by_role("textbox", name="Environment").fill(jde_environment)
        in_page.wait_for_timeout(500)
    if jde_role:
        in_page.get_by_role("textbox", name="Role").fill(jde_role)
        in_page.wait_for_timeout(500)

    in_page.get_by_role("textbox", name="User ID").fill(jde_user)
    in_page.wait_for_timeout(500)
    in_page.get_by_role("textbox", name="Password").fill(jde_pass)
    in_page.wait_for_timeout(500)
    in_page.get_by_role("button", name="Sign In").click()
    in_page.wait_for_timeout(3000)

    element_exists = in_page.locator("#drop_mainmenu")
    element_exists.wait_for(state="visible", timeout=20000)
    if element_exists.is_visible():
        logger.info("登录成功")
    else:
        raise Exception("登录失败，未进入JDE界面，请检查！")


def _enter_total_ledger(frame, timeout):
    """先进入所有页面，再进入总账页签（页签在 iframe 内点击）。"""
    select_page = frame.get_by_role("button", name="选择页面")
    select_page.wait_for(state="visible", timeout=timeout)
    select_page.click()
    all_pages_grid = frame.get_by_role("gridcell", name="所有页面")
    all_pages_grid.wait_for(state="visible", timeout=timeout)
    all_pages_grid.click()

    all_pages = frame.get_by_text("所有页面").first
    all_pages.click()
    if all_pages.is_visible():
        select_page.click()

    tab = frame.locator("#tabBar").get_by_text("总帐")
    tab.wait_for(state="visible", timeout=timeout)
    tab.click()


def _click_menu_link(in_page, name, timeout):
    """在主文档和当前 iframe 中查找可见菜单项并点击。"""
    deadline = time.monotonic() + timeout / 1000
    while True:
        for frame in in_page.frames:
            links = frame.get_by_role("link", name=name)
            for index in range(links.count()):
                link = links.nth(index)
                if link.is_visible():
                    logger.info(f"在页面/frame中找到菜单项并点击：{name}")
                    link.click(timeout=timeout)
                    return

        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise PlaywrightTimeoutError(
                f"等待 {timeout} 毫秒后，主文档和 iframe 中均未找到可见菜单项：{name}"
            )
        in_page.wait_for_timeout(min(500, max(1, int(remaining * 1000))))


def _set_clipboard_text(text):
    """把整块 TSV 写入系统剪贴板，供 JDE 网格 Ctrl+V 粘贴。"""
    try:
        import win32clipboard
    except ImportError as exc:
        raise RuntimeError(f"缺少 pywin32，无法写入剪贴板：{exc}") from exc
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


def _focus_journal_grid(frame, in_page, timeout):
    """把焦点移到分录网格首格，确认不了焦点就抛错，避免粘贴到日期栏。"""
    focused_selector = "tr[id] input.JSTextfield:focus"
    for selector in ("tr[id^='G0_'] input.JSTextfield", "#G0_1 input.JSTextfield"):
        cells = frame.locator(selector)
        if cells.count() == 0:
            continue
        cells.first.click(timeout=timeout)
        in_page.wait_for_timeout(300)
        if frame.locator(focused_selector).count() > 0:
            logger.info(f"焦点已移到分录网格：{selector}")
            return
    for _ in range(12):
        in_page.keyboard.press("Tab")
        in_page.wait_for_timeout(200)
        if frame.locator(focused_selector).count() > 0:
            logger.info("已用 Tab 把焦点移入分录网格")
            return
    raise RuntimeError("未能把焦点移到分录网格，已跳过粘贴，避免写入日期栏")


def _normalize_label(text):
    return re.sub(r"\s+", "", str(text or ""))


def _role_header_labels(frame):
    """读取网格 role=columnheader 的表头文字。"""
    labels = []
    try:
        headers = frame.get_by_role("columnheader")
        for index in range(headers.count()):
            item = headers.nth(index)
            try:
                if not item.is_visible():
                    continue
                label = _normalize_label(item.inner_text())
            except Exception:
                continue
            if label:
                labels.append(label)
    except Exception:
        return []
    return labels


def _th_header_labels(frame):
    """标准 table 表头兜底。"""
    labels = []
    try:
        headers = frame.locator("th")
        for index in range(headers.count()):
            item = headers.nth(index)
            try:
                if not item.is_visible():
                    continue
                label = _normalize_label(item.inner_text())
            except Exception:
                continue
            if label:
                labels.append(label)
    except Exception:
        return []
    return labels


def _position_header_labels(frame):
    """按表头文字与网格首行的相对位置读取列顺序。"""
    inputs = frame.locator("#G0_1 input.JSTextfield")
    if inputs.count() == 0:
        inputs = frame.locator("tr[id^='G0_'] input.JSTextfield")
    if inputs.count() == 0:
        return []
    try:
        anchor = inputs.first.bounding_box()
    except Exception:
        return []
    if not anchor:
        return []
    boxes = []
    for name in JOURNAL_COLUMNS:
        locator = frame.get_by_text(name)
        found = []
        for index in range(min(locator.count(), 20)):
            item = locator.nth(index)
            try:
                if not item.is_visible():
                    continue
                box = item.bounding_box()
            except Exception:
                continue
            if not box or box["width"] <= 0 or box["height"] <= 0:
                continue
            if box["y"] + box["height"] <= anchor["y"] + 2:
                found.append(box)
        if not found:
            return []
        found.sort(key=lambda box: box["y"] + box["height"])
        boxes.append(found[-1])
    tops = [box["y"] for box in boxes]
    if max(tops) - min(tops) > 40:
        return []
    return [
        name for name, _ in sorted(zip(JOURNAL_COLUMNS, boxes), key=lambda pair: pair[1]["x"])
    ]


def _journal_header_labels(frame):
    """返回能找到全部目标列的表头顺序；找不到返回空列表。"""
    for labels in (_role_header_labels(frame), _th_header_labels(frame),
                   _position_header_labels(frame)):
        if labels and all(name in labels for name in JOURNAL_COLUMNS):
            return labels
    return []


def _verify_journal_columns(frame, timeout):
    """粘贴前校验网格列顺序，不通过则中止，避免数据写入错误列。"""
    wait_timeout = min(timeout, 5000)
    try:
        frame.locator("#G0_1 input.JSTextfield").first.wait_for(
            state="visible", timeout=wait_timeout
        )
    except Exception:
        try:
            frame.locator("tr[id^='G0_'] input.JSTextfield").first.wait_for(
                state="visible", timeout=wait_timeout
            )
        except Exception:
            pass
    labels = _journal_header_labels(frame)
    if not labels:
        raise RuntimeError("无法读取日记帐分录网格表头，未执行粘贴")
    positions = {}
    for index, label in enumerate(labels):
        positions.setdefault(label, index)
    missing = [name for name in JOURNAL_COLUMNS if name not in positions]
    if missing:
        raise RuntimeError(
            f"分录网格缺少列：{'、'.join(missing)}；当前表头：{' / '.join(labels)}"
        )
    order = [name for name in sorted(positions, key=positions.get)]
    if tuple(order) != JOURNAL_COLUMNS:
        raise RuntimeError(
            f"分录网格列顺序不符：期望 {'、'.join(JOURNAL_COLUMNS)}，"
            f"实际 {'、'.join(order)}；当前表头：{' / '.join(labels)}"
        )
    logger.info(f"分录网格列顺序校验通过：{' / '.join(labels)}")


def _click_cancel(frame, timeout):
    """在当前 frame 点击一次取消，失败返回 False。"""
    try:
        button = frame.get_by_role("button", name="取消 (Ctrl+Alt+L)").first
        if button.count() == 0 or not button.is_visible():
            return False
        button.click(timeout=timeout)
        return True
    except Exception as exc:
        logger.warning(f"点击取消失败：{exc}")
        return False


def _click_cancel_anywhere(in_page, timeout, prefer_new=False):
    """跨页面/frame 查找并点击取消；找不到返回 False。"""
    pages = list(in_page.context.pages)
    if prefer_new:
        pages = [page for page in pages if page != in_page] + [in_page]
    for page in pages:
        try:
            frames = page.frames
        except Exception:
            continue
        for frame in frames:
            try:
                buttons = frame.get_by_role("button", name="取消 (Ctrl+Alt+L)")
                for index in range(min(buttons.count(), 5)):
                    button = buttons.nth(index)
                    if button.is_visible():
                        button.click(timeout=timeout)
                        return True
            except Exception:
                continue
    return False


def _discard_entry(frame, in_page, timeout):
    """取消分录录入：第二次取消可能不存在或出现在新页面，按尽力而为处理。"""
    if not _click_cancel(frame, timeout):
        _click_cancel_anywhere(in_page, timeout)
    in_page.wait_for_timeout(800)
    if _click_cancel_anywhere(in_page, min(timeout, 5000), prefer_new=True):
        logger.info("已点击第二次取消")
    else:
        logger.info("未发现第二次取消按钮，跳过")


def create_journal_entry(in_page, company, description, gl_date, rows,
                         scheme_option=JOURNAL_SCHEME_OPTION,
                         timeout=60000, discard=True):
    """录入日记帐分录并返回 JDE 生成的批号。

    rows 为 (帐号, 借项金额, 贷项金额, 备注) 序列，按 TSV 整块粘贴到第一个单元格，
    行数随数据动态变化；粘贴前先选分录方案并校验网格列顺序。
    discard=True 时按录制脚本取消退出（第二次取消可能不存在或位于新页面）。
    """
    if not rows:
        raise ValueError("分录没有可录入的数据行")
    home = in_page.get_by_title("主页")
    home.wait_for(state="visible", timeout=timeout)
    home.click()
    logger.info("已返回主页，重新选择页面并进入总帐")
    frame = in_page.frame_locator(MENU_IFRAME)
    _enter_total_ledger(frame, timeout)
    logger.info("打开日记帐分录")
    _click_menu_link(in_page, "日记帐分录", timeout)

    frame.get_by_role("button", name="增加 (Ctrl+Alt+A)").click(timeout=timeout)
    frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click(timeout=timeout)

    batch_box = frame.get_by_role("textbox", name="批号")
    batch_box.wait_for(state="visible", timeout=timeout)
    batch_box.click(timeout=timeout)
    batch = batch_box.input_value(timeout=timeout)
    logger.info(f"JDE 批号：{batch}")

    frame.get_by_role("textbox", name="Document Company").fill(
        str(company).strip(), timeout=timeout
    )
    frame.get_by_role("textbox", name="说明").fill(
        str(description).strip(), timeout=timeout
    )
    frame.get_by_role("textbox", name="总帐日期").fill(
        str(gl_date).strip(), timeout=timeout
    )

    lines = []
    for code, debit, credit, remark in rows:
        lines.append("\t".join("" if value is None else _cell_text(value)
                               for value in (code, debit, credit, remark)))
    _set_clipboard_text("\n".join(lines))
    logger.info(f"分录数据已写入剪贴板：{len(lines)} 行")

    scheme = frame.locator("#gtab0_1")
    scheme.wait_for(state="visible", timeout=timeout)
    scheme.select_option(label="RPA", timeout=timeout)
    logger.info("已选择分录方案：RPA")
    in_page.wait_for_timeout(1000)

    try:
        _verify_journal_columns(frame, timeout)
    except Exception:
        _discard_entry(frame, in_page, timeout)
        raise
    _focus_journal_grid(frame, in_page, timeout)
    in_page.keyboard.press("Control+V")
    in_page.wait_for_timeout(1000)

    if discard:
        _discard_entry(frame, in_page, timeout)
        logger.info("已取消分录录入（未保存）")
    return batch


def _cell_text(value):
    """数值去掉多余小数位，避免粘贴时出现 1.146979552E7 之类的科学计数法。"""
    if isinstance(value, float):
        return format(value, "f").rstrip("0").rstrip(".")
    return str(value)
