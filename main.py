# @Author: YiLong Li
# @Date: 2026/11/12 10:55
# @File: main.py
# @Software: PyCharm


"""
项目名称：TableKit 表格数据处理工作台

项目介绍：
将网页模板、Excel 解析库及图标封装为可独立运行的 HTML 工作台。
生成后无需持续运行 Python，可将 HTML 单独复制到其他电脑离线使用。
业务数据由浏览器读取并在本地处理，结果通过网页导出。

主要功能：
分页预览、多条件筛选、复合排序、删除与编辑、新增行、
重复记录处理、多表透视、字段映射与合并汇总、多格式导出及操作撤销。

目录说明：
assets/         网页模板、图标、Excel 解析库及许可证
assets/test/    清洗规则示例.json、预期结果.json
test/           模拟测试数据
help/           中文操作手册及项目说明
output/         生成的独立 HTML；不存在时自动创建

运行方式：
python main.py
python main.py --no-browser
uv run main.py

可选 Excel 清洗：
python main.py --clean-excel "业务.xlsx"
python main.py --clean-excel "业务.xlsx" --rules "自定义规则.json"

构建 HTML 仅使用 Python 标准库。
可选 Excel 清洗模式需要 openpyxl，不会覆盖原始业务文件。
"""


import argparse
import base64
import json
import sys
import webbrowser
from pathlib import Path

BASE = Path(__file__).resolve().parent
ASSETS_DIR = BASE / "assets"
DEFAULT_RULES = ASSETS_DIR / "test" / "清洗规则示例.json"


def main():
    """封装页面资源，生成可独立使用的 HTML。"""
    parser = argparse.ArgumentParser(
        description="生成可独立离线使用的表格数据工作台"
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=BASE / "output",
        help="生成 HTML 的目录",
    )
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="生成后不自动打开浏览器",
    )
    args = parser.parse_args()

    template = ASSETS_DIR / "dashboard_template.html"
    library = ASSETS_DIR / "xlsx.full.min.js"
    logo = ASSETS_DIR / "cml_logo_white.svg"
    license_file = ASSETS_DIR / "SheetJS-LICENSE.txt"

    for file in (template, library, logo, license_file):
        if not file.is_file():
            parser.error(f"缺少必要文件，请完整解压项目：{file}")

    text = template.read_text(encoding="utf-8")
    script = library.read_text(encoding="utf-8").replace(
        "</script", r"<\/script"
    )
    license_text = license_file.read_text(encoding="utf-8")

    text = text.replace(
        "</head>",
        "<!-- SheetJS third-party license:\n"
        + license_text.replace("--", "—")
        + "\n--></head>",
        1,
    )
    text = text.replace(
        '<script src="xlsx.full.min.js"></script>',
        "<script>" + script + "</script>",
    )
    text = text.replace(
        'src="cml_logo_white.svg"',
        'src="data:image/svg+xml;base64,'
        + base64.b64encode(logo.read_bytes()).decode("ascii")
        + '"',
    )

    output = args.output.expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)

    target = output / "表格数据处理工作台.html"
    target.write_text(text, encoding="utf-8")

    print(
        f"\n看板已生成：{target}\n"
        "可以关闭 Python，双击 HTML 独立使用。\n"
        "换电脑只需复制该 HTML；业务数据需在网页中重新选择。\n"
        "建议使用 Chrome 或 Edge 普通窗口。"
    )

    if not args.no_browser:
        webbrowser.open(target.as_uri())


def clean_excel(file_path, rules, logic="OR", output_dir=None):
    """按字段条件清洗全部工作表，另存新文件及逐表报告。"""
    from copy import copy

    from openpyxl import load_workbook
    from openpyxl.formula.translate import Translator

    source = Path(file_path).expanduser().resolve()

    if not source.is_file():
        raise FileNotFoundError(f"找不到源文件：{source}")

    if source.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError("保留格式模式仅支持 XLSX / XLSM")

    if logic not in ("AND", "OR"):
        raise ValueError("logic 必须为 AND 或 OR")

    if not isinstance(rules, list) or not rules:
        raise ValueError("至少设置一个清洗条件")

    for rule in rules:
        if not isinstance(rule, dict):
            raise ValueError("每条清洗规则必须是 JSON 对象")

        field = rule.get("field")
        if (
            not isinstance(field, str)
            or not field.strip()
            or rule.get("op") not in ("contains", "=", "empty")
        ):
            raise ValueError(
                "条件必须包含有效 field 与 op"
                "（contains、=、empty）"
            )

        if rule["op"] != "empty":
            value = rule.get("value")
            if value is None or not str(value).strip():
                raise ValueError(
                    "关键词不能为空，删除空白请使用 empty"
                )

    wb = load_workbook(
        source,
        keep_vba=source.suffix.lower() == ".xlsm",
    )
    cached = None

    try:
        cached = load_workbook(
            source,
            data_only=True,
            read_only=True,
        )

        total = sum(max(0, ws.max_row - 1) for ws in wb.worksheets)
        processed = 0
        reports = []
        uncertain = 0

        def blank(value):
            if value is None:
                return True

            text = str(value).translate(
                dict.fromkeys(
                    map(ord, "\u200b\ufeff\u2060\u200d")
                )
            )
            return text.strip() == ""

        for ws in wb.worksheets:
            columns = {}

            for cell in ws[1]:
                if cell.value is not None:
                    columns.setdefault(
                        str(cell.value).strip(),
                        cell.column,
                    )

            missing = [
                rule["field"]
                for rule in rules
                if rule["field"] not in columns
            ]
            before = max(0, ws.max_row - 1)

            if missing:
                processed += before
                reports.append(
                    {
                        "sheet": ws.title,
                        "before": before,
                        "deleted": 0,
                        "remaining": before,
                        "status": "跳过：缺少 " + "、".join(missing),
                    }
                )
                print(
                    f"{processed}/{total} | "
                    f"{ws.title}：缺少字段，跳过"
                )
                continue

            # 无缓存公式不直接视为空白，避免误删。
            values = cached[ws.title].iter_rows(
                min_row=2,
                max_row=ws.max_row,
                values_only=True,
            )

            target_row = 2
            deleted = 0

            for row_number, cached_values in enumerate(values, 2):
                matches = []

                for rule in rules:
                    column = columns[rule["field"]]
                    source_cell = ws.cell(row_number, column)
                    value = source_cell.value

                    if source_cell.data_type == "f":
                        value = (
                            cached_values[column - 1]
                            if column <= len(cached_values)
                            else None
                        )
                        if value is None:
                            uncertain += 1
                            matches.append(False)
                            continue

                    if rule["op"] == "empty":
                        matched = blank(value)
                    elif value is None:
                        matched = False
                    elif rule["op"] == "contains":
                        matched = (
                            str(rule["value"]).strip()
                            in str(value).strip()
                        )
                    else:
                        matched = (
                            str(value).strip()
                            == str(rule["value"]).strip()
                        )

                    matches.append(matched)

                matched = (
                    all(matches)
                    if logic == "AND"
                    else any(matches)
                )

                if matched:
                    deleted += 1
                else:
                    if target_row != row_number:
                        for column in range(1, ws.max_column + 1):
                            src = ws.cell(row_number, column)
                            dst = ws.cell(target_row, column)
                            value = src.value

                            if src.data_type == "f":
                                try:
                                    value = Translator(
                                        value,
                                        origin=src.coordinate,
                                    ).translate_formula(
                                        dst.coordinate
                                    )
                                except Exception:
                                    # 不支持的公式保留原文本。
                                    pass

                            dst.value = value
                            dst._style = copy(src._style)
                            dst.comment = copy(src.comment)
                            dst._hyperlink = copy(src.hyperlink)

                            if dst.hyperlink:
                                dst.hyperlink.ref = dst.coordinate

                        dimension = copy(
                            ws.row_dimensions[row_number]
                        )
                        dimension.index = target_row
                        ws.row_dimensions[target_row] = dimension

                    target_row += 1

                processed += 1

                if (
                    processed % max(1, total // 100) == 0
                    or processed == total
                ):
                    print(
                        f"\r{processed}/{total} "
                        f"({processed / max(1, total):.1%})"
                        f" | {ws.title}",
                        end="",
                        flush=True,
                    )

            if target_row <= ws.max_row:
                ws.delete_rows(
                    target_row,
                    ws.max_row - target_row + 1,
                )

            reports.append(
                {
                    "sheet": ws.title,
                    "before": before,
                    "deleted": deleted,
                    "remaining": target_row - 2,
                    "status": "完成",
                }
            )

        folder = (
            Path(output_dir).expanduser().resolve()
            if output_dir
            else source.parent
        )
        folder.mkdir(parents=True, exist_ok=True)

        stem = source.stem + "_清洗结果"
        destination = folder / (stem + source.suffix)
        serial = 2

        while (
            destination.exists()
            or destination.with_suffix(".report.json").exists()
        ):
            destination = folder / (
                f"{stem}_{serial}{source.suffix}"
            )
            serial += 1

        # 独占创建，避免覆盖源文件或已有结果。
        with destination.open("xb") as stream:
            wb.save(stream)

        report = {
            "logic": logic,
            "rules": rules,
            "sheets": reports,
            "deleted": sum(item["deleted"] for item in reports),
            "remaining": sum(item["remaining"] for item in reports),
            "uncertain_formula_checks": uncertain,
        }

        report_path = destination.with_suffix(".report.json")
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        print(
            f"\n保存：{destination}\n"
            f"报告：{report_path}\n"
            f"删除：{report['deleted']}；"
            f"保留：{report['remaining']}；"
            f"无缓存公式检查：{uncertain}"
        )

        return destination, report

    finally:
        wb.close()
        if cached is not None:
            cached.close()


def clean_excel_cli():
    """读取清洗规则并执行可选 Excel 清洗模式。"""
    parser = argparse.ArgumentParser(
        description="多工作表 Excel 保留格式清洗"
    )
    parser.add_argument(
        "--clean-excel",
        required=True,
        type=Path,
        help="需要清洗的 XLSX / XLSM 文件",
    )
    parser.add_argument(
        "--rules",
        type=Path,
        default=DEFAULT_RULES,
        help=(
            "JSON 清洗规则，默认使用 "
            "assets/test/清洗规则示例.json"
        ),
    )
    parser.add_argument(
        "--clean-output",
        type=Path,
        help="结果保存目录，默认保存到源文件所在目录",
    )
    args = parser.parse_args()

    rules_path = args.rules.expanduser().resolve()

    if not rules_path.is_file():
        parser.error(f"找不到清洗规则文件：{rules_path}")

    try:
        config = json.loads(
            rules_path.read_text(encoding="utf-8-sig")
        )
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        parser.error(f"清洗规则读取失败：{exc}")

    if not isinstance(config, dict):
        parser.error("清洗规则必须是 JSON 对象")

    if not isinstance(config.get("rules"), list) or not config["rules"]:
        parser.error("清洗规则中的 rules 必须是非空数组")

    if not all(
        isinstance(rule, dict)
        for rule in config["rules"]
    ):
        parser.error("rules 中的每条规则必须是 JSON 对象")

    try:
        clean_excel(
            args.clean_excel,
            config["rules"],
            config.get("logic", "OR"),
            args.clean_output,
        )
    except ImportError:
        parser.error(
            '可选清洗模式需安装：'
            'python -m pip install "openpyxl>=3.1.5,<4"'
        )
    except (OSError, ValueError) as exc:
        parser.error(f"清洗失败：{exc}")


if __name__ == "__main__":
    if any(
        arg == "--clean-excel"
        or arg.startswith("--clean-excel=")
        for arg in sys.argv[1:]
    ):
        clean_excel_cli()
    else:
        main()