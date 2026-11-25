"""土地使用税：复制模板、更新会计期间、核对后按需录入日记帐分录。"""

import calendar
import re
import shutil
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from config.config_loader import ConfigError
from data_processing.excel_com import (
    CellWrite,
    SheetData,
    WorkbookData,
    find_sheet,
    read_workbook,
    update_labels,
    write_cells,
)
from log.loggingConfig import logger
from web.JDE.jde_function import create_journal_entry

CALC_SHEET_KEYWORD = "土地使用税计算底稿"
DECLARATION_SHEET_KEYWORD = "土地使用税申报底稿"
PROPERTY_TABLE = "财产税和行为税纳税申报表"
LAND_TABLE = "城镇土地使用税税源明细表"
PERIOD_LABEL = "会计期间："
BATCH_LABEL = "批号："
MONTHLY_TAX_HEADER = "本月应补计税额"
LAND_AREA_HEADER = "土地面积"
JOURNAL_HEADERS = ("帐号", "借项金额", "贷项金额", "备注")
HEADER_ALIASES = {"账号": "帐号"}
DATE_FORMAT = "yyyy-m-d"
OUTPUT_NAME = "土地使用税计税表 {company} {period}.xlsx"
INVALID_FILE_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _text(value):
    """单元格文字；整数浮点去掉 .0，日期按年月日输出。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def _label(value):
    """标题比较用文本：去掉所有空白并转小写。"""
    return re.sub(r"\s+", "", _text(value)).lower()


def _header_label(value):
    label = _label(value)
    return HEADER_ALIASES.get(label, label)


def _number(value, where):
    if value is None or isinstance(value, bool):
        raise ConfigError(f"{where} 不是数值：{value!r}")
    try:
        result = Decimal(str(value))
    except (InvalidOperation, ValueError):
        raise ConfigError(f"{where} 不是数值：{value!r}") from None
    if not result.is_finite():
        raise ConfigError(f"{where} 不是有效数值：{value!r}")
    return result


def _amount(value, where):
    """分录金额：空白及“-”视为 0。"""
    text = _text(value)
    if text in ("", "-", "——", "—"):
        return Decimal(0)
    return _number(value, where)


def _sum(values, where):
    total = Decimal(0)
    for index, value in enumerate(values, 1):
        total += _amount(value, f"{where}第 {index} 行")
    return total


def configured_period(value):
    """会计期间配置值；空白（含只有空格）视为未配置，按自然月返回今天。"""
    text = "" if value is None else str(value).strip()
    return text or date.today()


def month_bounds(period):
    """返回会计期间自然月的第一天和最后一天。"""
    if isinstance(period, (datetime, date)):
        year, month = period.year, period.month
    else:
        text = "" if period is None else str(period).strip()
        match = re.fullmatch(r"(\d{4})[-/年]?(\d{1,2})月?", text)
        if not match:
            raise ConfigError(
                f"会计期间格式不正确，应为 YYYY-MM 或 YYYY/MM 或 YYYYMM 等可识别格式；当前值：{text}"
            )
        year, month = int(match.group(1)), int(match.group(2))
    if not 1 <= month <= 12:
        raise ConfigError(f"会计期间月份不正确，应为 1-12；当前值：{month}")
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def format_jde_date(value, label):
    """JDE 日期统一为 MM/DD/YYYY。"""
    if isinstance(value, (datetime, date)):
        return value.strftime("%m/%d/%Y")
    text = str(value).strip()
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, pattern).strftime("%m/%d/%Y")
        except ValueError:
            continue
    raise ConfigError(
        f"{label} 日期格式不正确，应为 YYYY-MM-DD 或 MM/DD/YYYY 等可识别格式；当前值：{text}"
    )


@dataclass
class HeaderLocation:
    row: int
    columns: Dict[str, int]


@dataclass
class TableLocation:
    title_row: int
    header: HeaderLocation


@dataclass
class CheckResult:
    passed: bool
    monthly_total: Decimal
    debit_total: Decimal
    credit_total: Decimal
    land_total: Decimal
    journal_rows: List[Tuple[str, object, object, str]]
    details: List[dict]


def find_header_row(sheet: SheetData, required: Sequence[str],
                    start_row: Optional[int] = None) -> HeaderLocation:
    """按标题文字定位表头行；标题忽略空白与大小写。"""
    wanted = {_header_label(name) for name in required}
    start = sheet.first_row if start_row is None else start_row
    for row in range(start, sheet.max_row + 1):
        columns: Dict[str, int] = {}
        for col in range(sheet.first_col, sheet.max_col + 1):
            label = _header_label(sheet.cell(row, col))
            if label in wanted and label not in columns:
                columns[label] = col
        if wanted <= set(columns):
            return HeaderLocation(row=row, columns=columns)
    names = "、".join(required)
    raise ConfigError(f"{sheet.name} 未找到表头：{names}")


def find_label_cell(sheet: SheetData, text: str, startswith: bool = False) -> Optional[Tuple[int, int]]:
    """按文字定位单元格，返回绝对行列坐标。"""
    target = _label(text)
    for row in range(sheet.first_row, sheet.max_row + 1):
        for col in range(sheet.first_col, sheet.max_col + 1):
            label = _label(sheet.cell(row, col))
            if (label.startswith(target) if startswith else label == target):
                return row, col
    return None


def find_total_row(sheet: SheetData, header_row: int, label_col: int) -> int:
    """在表头左侧列或全表中定位“合计”行。"""
    for row in range(header_row + 1, sheet.max_row + 1):
        if _label(sheet.cell(row, label_col)) == "合计":
            return row
    for row in range(header_row + 1, sheet.max_row + 1):
        for col in range(sheet.first_col, sheet.max_col + 1):
            if _label(sheet.cell(row, col)) == "合计":
                return row
    raise ConfigError(f"{sheet.name} 未找到“合计”行")


def read_journal_rows(sheet: SheetData) -> Tuple[HeaderLocation, List[Tuple[str, object, object, str]]]:
    """定位“帐号/借项金额/贷项金额/备注”表格，读取除标题外的数据行。"""
    header = find_header_row(sheet, JOURNAL_HEADERS)
    account_col = header.columns[_header_label("帐号")]
    debit_col = header.columns[_header_label("借项金额")]
    credit_col = header.columns[_header_label("贷项金额")]
    remark_col = header.columns[_header_label("备注")]
    rows: List[Tuple[str, object, object, str]] = []
    for row in range(header.row + 1, sheet.max_row + 1):
        account = _text(sheet.cell(row, account_col))
        if not account:
            break
        if _label(account) in ("合计", "总计"):
            break
        rows.append((
            account,
            sheet.cell(row, debit_col),
            sheet.cell(row, credit_col),
            _text(sheet.cell(row, remark_col)),
        ))
    if not rows:
        raise ConfigError(f"{sheet.name} 的分录表没有数据行")
    return header, rows


class LandUseTaxFlow:
    """单行土地使用税：模板复制、会计期间更新、核对与分录。"""

    def __init__(self, context):
        self.context = context
        self.config = context.config

    def get(self, key, default=None):
        return self.config.get(f"土地使用税配置.{key}", default)

    def engine(self):
        return str(self.get("重算引擎", "Excel.Application") or "Excel.Application").strip()

    def tolerance(self):
        raw = self.get("核对容差", "0.01")
        try:
            value = Decimal(str(raw).strip() or "0.01")
            if not value.is_finite() or value < 0:
                raise ValueError
            return value
        except (InvalidOperation, ValueError):
            raise ConfigError(f"土地使用税配置.核对容差 必须为非负金额；当前值：{raw!r}") from None

    def calc_sheet(self):
        """更新会计期间、读取分录与回写批号使用的工作表名。"""
        name = str(self.get("计算底稿工作表", CALC_SHEET_KEYWORD) or "").strip()
        return name or CALC_SHEET_KEYWORD

    def accounting_period(self):
        """会计期间首日：显式配置则用配置值，空白按自然月（当月）。"""
        first, _ = month_bounds(configured_period(self.get("会计期间")))
        return first

    def template_path(self):
        """共享盘模板「土地使用税计税表 {公司简称}.xlsx」。"""
        raw = self.get("模板文件")
        if raw is None or not str(raw).strip():
            raise ConfigError("土地使用税配置.模板文件 未配置")
        template = Path(str(raw).strip())
        if not template.is_absolute():
            template = self.config.get_path("input") / template
        if not template.is_file():
            raise ConfigError(f"土地使用税模板不存在：{template}")
        return template

    def output_path(self, template):
        """本期生成表：土地使用税计税表 {公司简称} {YYYYMM}.xlsx。"""
        company = str(self.get("公司", "") or "").strip()
        if not company or INVALID_FILE_CHARS.search(company):
            raise ConfigError("土地使用税配置.公司 必填且不能包含文件名非法字符")
        period = f"{self.accounting_period():%Y%m}"
        output_dir = Path(str(self.get("输出目录") or self.config.get_path("output")))
        if not output_dir.is_absolute():
            output_dir = self.config.get_path("output") / output_dir
        destination = output_dir / OUTPUT_NAME.format(company=company, period=period)
        if destination.resolve() == template.resolve():
            raise ConfigError(f"生成文件不能与模板重合：{destination}")
        return destination

    def create_output(self):
        """复制共享盘模板生成带期间的文件；同名文件直接覆盖。"""
        template = self.template_path()
        destination = self.output_path(template)
        destination.parent.mkdir(parents=True, exist_ok=True)
        if destination.exists():
            try:
                destination.unlink()
            except OSError as exc:
                raise ConfigError(
                    f"无法覆盖同名生成表 {destination}："
                    "文件可能仍被 Excel 占用，请关闭后重试"
                ) from exc
            logger.info(f"已覆盖同名生成表：{destination}")
        shutil.copy2(template, destination)
        logger.info(f"已复制模板生成本期计税表：{template} -> {destination}")
        return destination

    def update_period(self, path):
        """按文字动态定位「会计期间：」，写入本期期间首日并重算保存。"""
        first = self.accounting_period()
        update_labels(
            path,
            self.calc_sheet(),
            [(PERIOD_LABEL, first, DATE_FORMAT)],
            engine=self.engine(),
        )
        logger.info(f"会计期间已更新为 {first:%Y-%m-%d}：{path}")
        return first

    def _summary_totals(self, calc: SheetData):
        """读取合计行的“本月应补计税额”和“土地面积”。"""
        header = find_header_row(calc, (MONTHLY_TAX_HEADER, LAND_AREA_HEADER))
        label_col = min(header.columns.values())
        total_row = find_total_row(calc, header.row, label_col)
        monthly_total = _number(
            calc.cell(total_row, header.columns[_header_label(MONTHLY_TAX_HEADER)]),
            f"{calc.name}本月应补计税额合计",
        )
        land_total = _number(
            calc.cell(total_row, header.columns[_header_label(LAND_AREA_HEADER)]),
            f"{calc.name}土地面积合计",
        )
        return monthly_total, land_total

    def land_area_total(self, calc: SheetData) -> Decimal:
        """只读取土地面积合计，供更新申报底稿使用。"""
        _, land_total = self._summary_totals(calc)
        return land_total

    def check(self, calc: SheetData) -> CheckResult:
        """检查月度底稿：本月应补计税额合计=借项合计，借项合计=贷项合计。"""
        monthly_total, land_total = self._summary_totals(calc)

        _, journal_rows = read_journal_rows(calc)
        debit_total = _sum((row[1] for row in journal_rows), "借项金额")
        credit_total = _sum((row[2] for row in journal_rows), "贷项金额")
        tolerance = self.tolerance()

        details = []
        comparisons = (
            ("本月应补计税额合计=借项金额合计", monthly_total, debit_total),
            ("借项金额合计=贷项金额合计", abs(debit_total), abs(credit_total)),
        )
        passed = True
        for name, left, right in comparisons:
            difference = left - right
            ok = abs(difference) <= tolerance
            passed = passed and ok
            details.append({
                "核对项": name,
                "左值": str(left),
                "右值": str(right),
                "差额": str(difference),
                "容差": str(tolerance),
                "通过": ok,
            })
            logger.info(
                f"[{'通过' if ok else '不通过'}] {name}：左={left} 右={right} "
                f"差额={difference} 容差={tolerance}"
            )
        return CheckResult(
            passed=passed,
            monthly_total=monthly_total,
            debit_total=debit_total,
            credit_total=credit_total,
            land_total=land_total,
            journal_rows=journal_rows,
            details=details,
        )

    def declaration_year(self, sheet: SheetData) -> int:
        """读取“所属期：”右侧单元格中的年份。"""
        found = find_label_cell(sheet, "所属期", startswith=True)
        if found is None:
            raise ConfigError(f"{sheet.name} 未找到“所属期：”单元格")
        row, col = found
        for offset in range(1, 4):
            value = sheet.cell(row, col + offset)
            year = self._year(value)
            if year is not None:
                return year
        raise ConfigError(f"{sheet.name} “所属期：”右侧未找到年份")

    @staticmethod
    def _year(value):
        if isinstance(value, (datetime, date)):
            return value.year
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            if 1900 <= value <= 2100:
                return int(value)
        match = re.search(r"(19|20)\d{2}", _text(value))
        return int(match.group(0)) if match else None

    def _find_table(self, sheet: SheetData, title: str, headers: Sequence[str]) -> TableLocation:
        found = find_label_cell(sheet, title)
        if found is None:
            raise ConfigError(f"{sheet.name} 未找到“{title}”")
        title_row, _ = found
        header = find_header_row(sheet, headers, start_row=title_row + 1)
        return TableLocation(title_row=title_row, header=header)

    def _find_serial_one(self, sheet: SheetData, title: str, table: TableLocation) -> int:
        serial_col = table.header.columns[_header_label("序号")]
        for row in range(table.header.row + 1, sheet.max_row + 1):
            value = _text(sheet.cell(row, serial_col))
            if value in ("1", "1.0"):
                return row
            if _label(value) in ("合计", "总计"):
                break
        raise ConfigError(f"{sheet.name} “{title}”未找到序号 1 行")

    def plan_update(self, declaration: SheetData, year: int, land_total: Decimal):
        """生成年度申报底稿的写入清单。"""
        property_table = self._find_table(
            declaration, PROPERTY_TABLE,
            ("序号", "计税依据", "税款所属期起", "税款所属期止"),
        )
        land_table = self._find_table(
            declaration, LAND_TABLE,
            ("序号", "占用土地面积", "有效起期", "有效止期"),
        )
        property_row = self._find_serial_one(declaration, PROPERTY_TABLE, property_table)
        land_row = self._find_serial_one(declaration, LAND_TABLE, land_table)

        start = date(year, 1, 1)
        end = date(year, 12, 31)
        writes = [
            CellWrite(
                declaration.name, property_row,
                property_table.header.columns[_header_label("计税依据")], land_total,
            ),
            CellWrite(
                declaration.name, property_row,
                property_table.header.columns[_header_label("税款所属期起")], start, DATE_FORMAT,
            ),
            CellWrite(
                declaration.name, property_row,
                property_table.header.columns[_header_label("税款所属期止")], end, DATE_FORMAT,
            ),
            CellWrite(
                declaration.name, land_row,
                land_table.header.columns[_header_label("占用土地面积")], land_total,
            ),
            CellWrite(
                declaration.name, land_row,
                land_table.header.columns[_header_label("有效起期")], start, DATE_FORMAT,
            ),
            CellWrite(
                declaration.name, land_row,
                land_table.header.columns[_header_label("有效止期")], end, DATE_FORMAT,
            ),
        ]
        logger.info(
            f"申报底稿更新：{declaration.name} "
            f"计税依据/占用土地面积={land_total}，"
            f"税款所属期/有效起止={start:%Y-%m-%d}~{end:%Y-%m-%d}"
        )
        return writes

    def needs_upload(self, result: CheckResult) -> bool:
        """借贷金额全为 0 视为本期无需计提，不上传分录。"""
        tolerance = self.tolerance()
        return abs(result.debit_total) > tolerance or abs(result.credit_total) > tolerance

    def write_journal(self, rows, period: str) -> str:
        """把计算底稿分录整块录入 JDE 并返回批号。"""
        company = str(self.get("分录公司", "") or "").strip()
        if not company:
            raise ConfigError("土地使用税配置.分录公司 未配置")
        browser = self.context.get_browser("JDE")
        page = browser.get_page() if browser else None
        if page is None:
            raise RuntimeError("JDE 浏览器未打开；请先登录 JDE")
        description = next((row[3] for row in rows if str(row[3]).strip()), "")
        if not description:
            description = f"计提{period}土地使用税"
        logger.info(f"JDE 分录公司：{company}；分录行数：{len(rows)}")
        batch = create_journal_entry(
            page,
            company=company,
            description=description,
            gl_date=format_jde_date(date.today(), "总帐日期"),
            rows=rows,
        )
        self.context.set("jde_journal_batch", batch)
        self.context.set("jde_journal_rows", len(rows))
        return batch

    def write_batch_number(self, path, batch: str):
        """按文字动态定位「批号：」，把 JDE 批号写入其右侧单元格。"""
        if not batch:
            return
        update_labels(
            path,
            self.calc_sheet(),
            [(BATCH_LABEL, batch, "@")],
            engine=self.engine(),
        )
        logger.info(f"JDE 批号 {batch} 已回写：{path}")

    def run(self):
        """执行单行：复制模板 -> 更新会计期间 -> 更新申报底稿 -> 核对 -> 按需录入分录 -> 回写批号。"""
        self.context.set("land_tax_status", "running")
        self.context.set("land_tax_note", None)
        try:
            path = self.create_output()
            self.update_period(path)

            book: WorkbookData = read_workbook(path, engine=self.engine())
            calc = find_sheet(book, self.calc_sheet())
            land_total = self.land_area_total(calc)

            declaration = find_sheet(book, DECLARATION_SHEET_KEYWORD)
            year = self.declaration_year(declaration)
            writes = self.plan_update(declaration, year, land_total)
            write_cells(path, writes, engine=self.engine())

            # 申报底稿写入并重算保存后，再读取生成表做核对。
            book = read_workbook(path, engine=self.engine())
            calc = find_sheet(book, self.calc_sheet())
            result = self.check(calc)
            if not result.passed:
                failed = "；".join(
                    f"{item['核对项']}（左={item['左值']} 右={item['右值']}）"
                    for item in result.details if not item["通过"]
                )
                raise ConfigError(f"土地使用税文件检查未通过：{failed}")

            period = f"{self.accounting_period():%Y%m}"
            upload = self.needs_upload(result)
            self.context.set("land_tax_summary", {
                "文件": str(path),
                "公司": str(self.get("公司", "") or "").strip(),
                "分录公司": str(self.get("分录公司", "") or "").strip(),
                "期间": period,
                "本月应补计税额": str(result.monthly_total),
                "土地面积": str(result.land_total),
                "借项金额合计": str(result.debit_total),
                "贷项金额合计": str(result.credit_total),
                "分录行数": len(result.journal_rows),
                "是否需要上传": "是" if upload else "否",
            })

            if not upload:
                note = "分录借贷金额均为 0，无需上传"
                logger.info(f"{path.name} {note}")
                self.context.set("land_tax_note", note)
                self.context.set("land_tax_status", "completed")
                return result

            batch = self.write_journal(result.journal_rows, period)
            self.write_batch_number(path, batch)
            self.context.set("land_tax_status", "completed")
            return result
        except Exception:
            self.context.set("land_tax_status", "failed")
            raise
