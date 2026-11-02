"""通过 Excel COM 读取重算后的值并写回单元格，兼容 .xls 与 .xlsx。"""
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Dict, Iterable, Optional, Tuple

from config.config_loader import ConfigError
from log.loggingConfig import logger

DEFAULT_ENGINE = "Excel.Application"


@dataclass
class SheetData:
    """一张工作表的已用区域值；first_row/first_col 为绝对坐标起点。"""

    name: str
    first_row: int
    first_col: int
    values: Tuple[Tuple[object, ...], ...] = ()

    @property
    def max_row(self) -> int:
        return self.first_row + len(self.values) - 1 if self.values else self.first_row - 1

    @property
    def max_col(self) -> int:
        width = max((len(row) for row in self.values), default=0)
        return self.first_col + width - 1 if width else self.first_col - 1

    def cell(self, row: int, col: int):
        row_index = row - self.first_row
        col_index = col - self.first_col
        if row_index < 0 or col_index < 0:
            return None
        if row_index >= len(self.values):
            return None
        values = self.values[row_index]
        if col_index >= len(values):
            return None
        return values[col_index]


@dataclass
class WorkbookData:
    """工作簿值快照。"""

    path: Path
    sheets: Dict[str, SheetData] = field(default_factory=dict)


@dataclass
class CellWrite:
    """一次单元格写入。"""

    sheet: str
    row: int
    col: int
    value: object
    number_format: Optional[str] = None


def _normalize_values(values, rows: int, cols: int):
    if rows <= 0 or cols <= 0 or values is None:
        return ()
    if rows == 1 and cols == 1:
        return ((values,),)
    if rows == 1:
        return (tuple(values),)
    if cols == 1:
        return tuple((item,) for item in values)
    return tuple(tuple(row) for row in values)


def _open_excel(engine: str):
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ConfigError(f"缺少 pywin32，无法调用 Excel COM：{exc}") from exc

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx(engine)
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AlertBeforeOverwriting = False
    app.AutomationSecurity = 3
    return pythoncom, app


def read_workbook(path: Path, engine: str = DEFAULT_ENGINE) -> WorkbookData:
    """只读打开工作簿，重算后读取各工作表已用区域，不保存文件。"""
    path = Path(path).resolve()
    if not path.is_file():
        raise ConfigError(f"待处理文件不存在：{path}")
    if path.stat().st_size == 0:
        raise ConfigError(f"待处理文件为空：{path}")

    pythoncom, app = _open_excel(engine)
    book = None
    try:
        try:
            book = app.Workbooks.Open(
                Filename=str(path),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
        except Exception as exc:
            raise ConfigError(
                f"Excel 无法打开待处理文件：{path}（{path.stat().st_size} 字节）；"
                "请确认文件未被其他程序占用"
            ) from exc

        try:
            app.CalculateFullRebuild()
        except Exception as exc:
            logger.warning(f"公式重算失败，将使用文件中的现有值：{exc}")

        data = WorkbookData(path=path)
        for index in range(1, book.Sheets.Count + 1):
            sheet = book.Worksheets(index)
            used = sheet.UsedRange
            rows, cols = int(used.Rows.Count), int(used.Columns.Count)
            values = _normalize_values(used.Value, rows, cols)
            data.sheets[str(sheet.Name)] = SheetData(
                name=str(sheet.Name),
                first_row=int(used.Row),
                first_col=int(used.Column),
                values=values,
            )
        logger.info(f"已读取工作簿：{path}（{len(data.sheets)} 个工作表）")
        return data
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                app.Quit()
            finally:
                pythoncom.CoUninitialize()


def _date_serial(value) -> float:
    moment = value if isinstance(value, datetime) else datetime(
        value.year, value.month, value.day
    )
    return (moment - datetime(1899, 12, 30)).total_seconds() / 86400


def write_cells(path: Path, writes: Iterable[CellWrite], engine: str = DEFAULT_ENGINE) -> int:
    """打开工作簿写入单元格，重算并保存；.xls 与 .xlsx 均保持原格式。"""
    path = Path(path).resolve()
    writes = list(writes)
    if not writes:
        return 0
    if not path.is_file():
        raise ConfigError(f"待更新文件不存在：{path}")

    pythoncom, app = _open_excel(engine)
    book = None
    try:
        try:
            book = app.Workbooks.Open(
                Filename=str(path),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
        except Exception as exc:
            raise ConfigError(f"Excel 无法打开待更新文件：{path}") from exc
        if book.ReadOnly:
            raise ConfigError("待更新文件被占用或为只读，无法保存更新结果")

        for write in writes:
            try:
                sheet = book.Worksheets(write.sheet)
            except Exception as exc:
                raise ConfigError(f"工作簿缺少工作表：{write.sheet}") from exc
            cell = sheet.Cells(write.row, write.col)
            value = write.value
            if write.number_format:
                cell.NumberFormat = write.number_format
            if value is None:
                cell.ClearContents()
            elif isinstance(value, str) and value.startswith("="):
                cell.Formula = value
            elif isinstance(value, (date, datetime)):
                cell.Value2 = _date_serial(value)
                cell.NumberFormat = write.number_format or "yyyy-m-d"
            else:
                if hasattr(value, "to_eng_string"):
                    value = float(value)
                cell.Value = value
            logger.info(
                f"已写入 {write.sheet}!R{write.row}C{write.col} = {write.value!r}"
            )

        try:
            app.CalculateFullRebuild()
        except Exception as exc:
            logger.warning(f"写入后重算失败：{exc}")
        book.Save()
        logger.info(f"已保存更新：{path}（{len(writes)} 个单元格）")
        return len(writes)
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                app.Quit()
            finally:
                pythoncom.CoUninitialize()


def find_sheet(book: WorkbookData, keyword: str) -> SheetData:
    """按工作表名包含的关键字查找；忽略空白差异并忽略大小写。"""
    def plain(text):
        return "".join(str(text or "").split()).lower()

    target = plain(keyword)
    matched = [sheet for name, sheet in book.sheets.items() if target in plain(name)]
    if not matched:
        names = "、".join(book.sheets)
        raise ConfigError(f"工作簿缺少包含“{keyword}”的工作表；当前工作表：{names}")
    if len(matched) > 1:
        names = "、".join(sheet.name for sheet in matched)
        raise ConfigError(f"包含“{keyword}”的工作表不唯一：{names}")
    return matched[0]
