"""打包脚本：生成单文件 exe。

用法：
    python build.py

依赖：
    pip install pyinstaller openpyxl
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ENTRY = os.path.join(HERE, "src", "db_view_export.py")
NAME = "DB视图导出"

def main():
    try:
        import PyInstaller.__main__ as pyi
    except ImportError:
        print("缺少 PyInstaller，请先执行： pip install pyinstaller")
        return 1

    if not os.path.isfile(ENTRY):
        print("找不到源码：%s" % ENTRY)
        return 1

    os.chdir(HERE)
    pyi.run([
        "--noconfirm",
        "--clean",
        "--onefile",
        "--console",
        "--name", NAME,
        ENTRY,
    ])
    print("")
    print("构建完成：%s" % os.path.join(HERE, "dist", NAME + ".exe"))
    return 0

if __name__ == "__main__":
    sys.exit(main())
