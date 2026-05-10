import sys
import types
from pathlib import Path
from typing import cast

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet


def _install_dify_plugin_stub() -> None:
    if "dify_plugin" in sys.modules:
        return

    dify_plugin_module = types.ModuleType("dify_plugin")

    class Tool:  # pragma: no cover - minimal import stub
        pass

    setattr(dify_plugin_module, "Tool", Tool)

    entities_module = types.ModuleType("dify_plugin.entities")

    class I18nObject:  # pragma: no cover - minimal import stub
        def __init__(self, **kwargs):
            self.values = kwargs

    setattr(entities_module, "I18nObject", I18nObject)

    tool_module = types.ModuleType("dify_plugin.entities.tool")

    class ToolInvokeMessage:  # pragma: no cover - minimal import stub
        pass

    class ToolParameter:  # pragma: no cover - minimal import stub
        class ToolParameterType:
            FILE = "file"

        class ToolParameterForm:
            FORM = "form"

        def __init__(self, *args, **kwargs):
            self.args = args
            self.kwargs = kwargs

    setattr(tool_module, "ToolInvokeMessage", ToolInvokeMessage)
    setattr(tool_module, "ToolParameter", ToolParameter)

    file_package = types.ModuleType("dify_plugin.file")
    file_module = types.ModuleType("dify_plugin.file.file")

    class File:  # pragma: no cover - minimal import stub
        pass

    setattr(file_module, "File", File)
    setattr(file_package, "file", file_module)

    sys.modules["dify_plugin"] = dify_plugin_module
    sys.modules["dify_plugin.entities"] = entities_module
    sys.modules["dify_plugin.entities.tool"] = tool_module
    sys.modules["dify_plugin.file"] = file_package
    sys.modules["dify_plugin.file.file"] = file_module


_install_dify_plugin_stub()

from tools.excel_extractor import ExcelExtractorTool


@pytest.fixture
def excel_tool() -> ExcelExtractorTool:
    return object.__new__(ExcelExtractorTool)


def test_render_row_text_preserves_internal_blanks_and_trims_trailing_blanks(
    excel_tool: ExcelExtractorTool,
) -> None:
    row_text = excel_tool._render_row_text((None, "B", None, "C", None))

    assert row_text == " | B |  | C"


def test_extract_text_xlsx_preserves_blank_cells_between_values(
    excel_tool: ExcelExtractorTool, tmp_path: Path
) -> None:
    workbook = Workbook()
    sheet = cast(Worksheet, workbook.active)
    sheet.title = "Sheet1"
    sheet.append(["A", None, "C"])
    sheet.append(["A", "B", "C"])
    temp_path = tmp_path / "blank-cells.xlsx"
    workbook.save(temp_path)

    extracted_text = excel_tool._extract_text_xlsx(str(temp_path))

    assert "Row 1: A |  | C" in extracted_text
    assert "Row 2: A | B | C" in extracted_text