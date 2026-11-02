"""土地使用税：检查月度底稿、更新年度申报底稿并录入日记帐分录。"""

import re
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Optional, Sequence, Tuple

from config.config_loader import ConfigError
from data_processing.excel_com import (
    CellWrite,
    SheetData,
    WorkbookData,
    find_sheet,
    read_workbook,
    write_cells,
)
from log.loggingConfig import logger
from web.JDE.jde_function import create_journal_entry

CALC_SHEET_KEYWORD = "土地使用税计算底稿"
DECLARATION_SHEET_KEYWORD = "土地使用税申报底稿"
PROPERTY_TABLE = "财产税和行为税纳税申报表"
LAND_TABLE = "城镇土地使用税税源明细表"
MONTHLY_TAX_HEADER = "本月应补计税额"
LAND_AREA_HEADER = "土地面积"
JOURNAL_HEADERS = ("帐号", "借项金额", "贷项金额", "备注")
HEADER_ALIASES = {"账号": "帐号"}
DATE_FORMAT = "yyyy-m-d"


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


def find_header_row(sheet: SheetData, required: Sequence[str], start_row: Optional[int] = None) -> HeaderLocation:
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
    """单个土地使用税文件的检查、更新与分录。"""

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

    def check(self, calc: SheetData) -> CheckResult:
        """检查月度底稿：本月应补计税额合计=借项合计，借项合计=贷项合计。"""
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
    def _year(value) -> Optional[int]:
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

    def write_journal(self, rows, period: str) -> str:
        """把月度底稿分录表整块录入 JDE 并返回批号。"""
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

    def write_batch_number(self, path, calc: SheetData, batch: str):
        """把 JDE 批号回写到“批号：”右侧第一个单元格。"""
        if not batch:
            return
        found = find_label_cell(calc, "批号", startswith=True)
        if found is None:
            logger.warning(f"{calc.name} 未找到“批号：”单元格，批号未回写")
            return
        row, col = found
        write_cells(
            path,
            [CellWrite(calc.name, row, col + 1, str(batch), "@")],
            engine=self.engine(),
        )

    def run(self, path, file_company: str, period: str):
        """执行单个文件：检查 -> 更新申报底稿 -> 录入分录 -> 回写批号。"""
        self.context.set("land_tax_status", "running")
        try:
            book: WorkbookData = read_workbook(path, engine=self.engine())
            calc = find_sheet(book, CALC_SHEET_KEYWORD)
            result = self.check(calc)
            if not result.passed:
                failed = "；".join(
                    f"{item['核对项']}（左={item['左值']} 右={item['右值']}）"
                    for item in result.details if not item["通过"]
                )
                raise ConfigError(f"土地使用税文件检查未通过：{failed}")

            declaration = find_sheet(book, DECLARATION_SHEET_KEYWORD)
            year = self.declaration_year(declaration)
            writes = self.plan_update(declaration, year, result.land_total)
            write_cells(path, writes, engine=self.engine())

            self.context.set("land_tax_summary", {
                "文件": str(path),
                "公司": file_company,
                "分录公司": str(self.get("分录公司", "") or "").strip(),
                "期间": period,
                "本月应补计税额": str(result.monthly_total),
                "土地面积": str(result.land_total),
                "借项金额合计": str(result.debit_total),
                "贷项金额合计": str(result.credit_total),
                "分录行数": len(result.journal_rows),
            })
            batch = self.write_journal(result.journal_rows, period)
            self.write_batch_number(path, calc, batch)
            self.context.set("land_tax_status", "completed")
            return result
        except Exception:
            self.context.set("land_tax_status", "failed")
            raise
