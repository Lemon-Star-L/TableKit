# @Author: YiLong Li
# @Date: 2026/11/12 10:55
# @File: main.py
# @Software: PyCharm


"""
项目名称：表格数据处理工作台

项目介绍：
将网页模板、Excel 解析库及 CML 图标打包为可独立运行的 HTML 看板。
生成后即可关闭 Python，将 HTML 单独复制到其他电脑使用，无需联网或后台服务。
数据由浏览器读取用户选择的文件，缓存于浏览器本地数据库，结果支持 CSV、XLSX、TSV、JSON 导出。

主要功能：
大表分页预览、表单切换、多条件筛选、复合排序、指定或排除字段去重、单表及多表透视、多格式导出与会话内操作撤销。

目录说明：
test/          模拟测试数据
assets/         看板模板、CML 图标及离线 Excel 解析库
output/  运行后自动生成，保存可独立使用的 HTML

运行方式：
python main.py 或 uv run main.py
构建 HTML 仅使用标准库；可选 --clean-excel 模式需要 openpyxl。
"""


import argparse
import base64
from pathlib import Path
import webbrowser

BASE = Path(__file__).resolve().parent

def main():
    parser = argparse.ArgumentParser(description='生成可独立离线使用的表格数据看板')
    parser.add_argument('--output', type=Path, default=BASE/'output', help='生成 HTML 的目录')
    parser.add_argument('--no-browser', action='store_true', help='生成后不自动打开浏览器')
    args = parser.parse_args()
    template = BASE/'assets'/'dashboard_template.html'
    library = BASE/'assets'/'xlsx.full.min.js'
    logo = BASE/'assets'/'cml_logo_white.svg'
    for file in (template, library, logo, BASE/'assets'/'SheetJS-LICENSE.txt'):
        if not file.is_file():
            parser.error(f'缺少必要文件，请完整解压项目：{file}')
    text = template.read_text(encoding='utf-8')
    # Escape closing script tags without changing JS string values.
    script = library.read_text(encoding='utf-8').replace('</script', r'<\/script')
    license_text = (BASE/'assets'/'SheetJS-LICENSE.txt').read_text(encoding='utf-8')
    text = text.replace('</head>', '<!-- SheetJS third-party license:\n' + license_text.replace('--', '—') + '\n--></head>', 1)
    text = text.replace('<script src="xlsx.full.min.js"></script>', '<script>' + script + '</script>')
    text = text.replace('src="cml_logo_white.svg"', 'src="data:image/svg+xml;base64,' + base64.b64encode(logo.read_bytes()).decode('ascii') + '"')
    output = args.output.expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    target = output/'表格数据处理工作台.html'
    target.write_text(text, encoding='utf-8')
    print(f'\n看板已生成：{target}\n可以关闭 Python，双击 HTML 独立使用。\n换电脑只需复制该 HTML；业务数据需在网页中重新选择。\n建议使用 Chrome 或 Edge 普通窗口。')
    if not args.no_browser:
        webbrowser.open(target.as_uri())

def clean_excel(file_path, rules, logic='OR', output_dir=None):
    """按字段清洗全部工作表，保留未匹配表；另存新文件及逐表报告。"""
    import json
    from copy import copy
    from openpyxl import load_workbook
    from openpyxl.formula.translate import Translator
    source = Path(file_path).resolve()
    if source.suffix.lower() not in ('.xlsx', '.xlsm'):
        raise ValueError('保留格式模式仅支持 XLSX / XLSM')
    if logic not in ('AND', 'OR') or not rules:
        raise ValueError('至少设置一个条件，logic 必须为 AND 或 OR')
    for rule in rules:
        if not str(rule.get('field', '')).strip() or rule.get('op') not in ('contains', '=', 'empty'):
            raise ValueError('条件必须包含 field 与 op（contains、=、empty）')
        if rule['op'] != 'empty' and not str(rule.get('value', '')).strip():
            raise ValueError('关键词不能为空，删除空白请使用 empty')
    wb = load_workbook(source, keep_vba=source.suffix.lower()=='.xlsm')
    cached = None
    try:
        cached = load_workbook(source, data_only=True, read_only=True)
        total = sum(max(0, w.max_row-1) for w in wb.worksheets)
        processed = 0; reports = []; uncertain = 0
        def blank(v):
            return v is None or str(v).translate(dict.fromkeys(map(ord, '\u200b\ufeff\u2060\u200d'))).strip()==''
        for ws in wb.worksheets:
            columns = {}
            for c in ws[1]:
                if c.value is not None: columns.setdefault(str(c.value).strip(), c.column)
            missing = [r['field'] for r in rules if r['field'] not in columns]
            before = max(0, ws.max_row-1)
            if missing:
                processed += before
                reports.append(dict(sheet=ws.title, before=before, deleted=0, remaining=before, status='跳过：缺少 '+ '、'.join(missing)))
                print(f'{processed}/{total} | {ws.title}：缺少字段，跳过')
                continue
            # 无缓存公式不作为空白删除，防止把无法判断的值当作空结果。
            values = cached[ws.title].iter_rows(min_row=2, max_row=ws.max_row, values_only=True)
            target = 2; deleted = 0
            for row, cv in enumerate(values, 2):
                matches = []
                for r in rules:
                    col = columns[r['field']]; c = ws.cell(row,col); v = c.value
                    if c.data_type == 'f':
                        v = cv[col-1] if col <= len(cv) else None
                        if v is None:
                            uncertain += 1; matches.append(False); continue
                    matches.append(blank(v) if r['op']=='empty' else v is not None and (str(r['value']).strip() in str(v).strip() if r['op']=='contains' else str(v).strip()==str(r['value']).strip()))
                matched = all(matches) if logic=='AND' else any(matches)
                if matched: deleted += 1
                else:
                    if target != row:
                        for col in range(1,ws.max_column+1):
                            src,dst = ws.cell(row,col),ws.cell(target,col); value = src.value
                            if src.data_type=='f':
                                try: value=Translator(value,origin=src.coordinate).translate_formula(dst.coordinate)
                                except Exception: pass
                            dst.value=value;dst._style=copy(src._style);dst.comment=copy(src.comment)
                            dst._hyperlink=copy(src.hyperlink)
                            if dst.hyperlink: dst.hyperlink.ref=dst.coordinate
                        dim=copy(ws.row_dimensions[row]);dim.index=target;ws.row_dimensions[target]=dim
                    target+=1
                processed+=1
                if processed % max(1,total//100)==0 or processed==total:
                    print(f'\r{processed}/{total} ({processed/max(1,total):.1%}) | {ws.title}',end='',flush=True)
            if target<=ws.max_row: ws.delete_rows(target,ws.max_row-target+1)
            reports.append(dict(sheet=ws.title,before=before,deleted=deleted,remaining=target-2,status='完成'))
        folder=Path(output_dir).resolve() if output_dir else source.parent;folder.mkdir(parents=True,exist_ok=True)
        stem=source.stem+'_清洗结果';dest=folder/(stem+source.suffix);i=2
        while dest.exists() or dest.with_suffix('.report.json').exists():
            dest=folder/(stem+'_'+str(i)+source.suffix);i+=1
        # 独占创建，避免覆盖源文件或已有结果。
        with dest.open('xb') as stream: wb.save(stream)
        report=dict(logic=logic,rules=rules,sheets=reports,deleted=sum(r['deleted'] for r in reports),remaining=sum(r['remaining'] for r in reports),uncertain_formula_checks=uncertain)
        dest.with_suffix('.report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print('\n保存：',dest, '\n删除：',report['deleted'],'；保留：',report['remaining'],'；无缓存公式检查：',uncertain)
        return dest,report
    finally:
        wb.close()
        if cached: cached.close()


def clean_excel_cli():
    import json
    parser=argparse.ArgumentParser(description='多工作表 Excel 保留格式清洗')
    parser.add_argument('--clean-excel',required=True,type=Path)
    parser.add_argument('--rules',required=True,type=Path,help='JSON 条件配置文件')
    parser.add_argument('--clean-output',type=Path)
    args=parser.parse_args()
    config=json.loads(args.rules.read_text(encoding='utf-8-sig'))
    try: clean_excel(args.clean_excel,config['rules'],config.get('logic','OR'),args.clean_output)
    except ImportError: parser.error('可选清洗模式需安装：python -m pip install openpyxl')

if __name__ == '__main__':
    import sys
    if '--clean-excel' in sys.argv: clean_excel_cli()
    else: main()