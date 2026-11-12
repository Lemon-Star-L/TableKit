# @Author: YiLong Li
# @Date: 2026/11/12 10:55
# @File: main.py
# @Software: PyCharm


"""
项目名称：表格数据处理工作台

项目介绍：
将网页模板、Excel 解析库及 CML 图标打包为可独立运行的 HTML 看板。
生成后即可关闭 Python，将 HTML 单独复制到其他电脑使用，无需联网或后台服务。
数据由浏览器读取用户选择的文件，缓存于浏览器本地数据库，结果以 CSV 下载。

主要功能：
大表分页预览、表单切换、多条件筛选、复合排序、指定或排除字段去重、单表及多表透视。

目录说明：
test/          模拟测试数据
assets/         看板模板、CML 图标及离线 Excel 解析库
output/  运行后自动生成，保存可独立使用的 HTML

运行方式：
python main.py 或 uv run main.py
只使用 Python 标准库，无需安装第三方 Python 依赖。
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

if __name__ == '__main__':
    main()