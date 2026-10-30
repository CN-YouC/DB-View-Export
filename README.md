# DB 视图导出工具

把 SQLite 数据库里的**视图**（可选表）逐个导出成 Excel（`.xlsx`）或 CSV 的小工具。

只做一件事：读数据库 → 写文件。不做分析、不联网、默认不改原库。

## 为什么需要它

- 数据库里的视图在图形工具里打不开、显示错位、列宽乱掉
- 想把视图数据发给别人看，但对方没有装数据库工具
- 想快速核对某个视图有多少行、有哪些列

## 特点

- **拖拽即用**：把 `.db`、`.lnk` 快捷方式，或含 `.db` 的文件夹直接拖进窗口
- **默认只读**：以 `mode=ro&immutable=1` 打开，不改原库、不产生 `-wal` / `-shm` 文件
- **整库或单表**：可以一键导出全部视图，也可以只导指定的某几个
- **格式可选**：`.xlsx` 或 CSV
- **大库优化**：`--optimize` 在临时副本上建索引后再导，原文件一个字节都不动；导出结束自动清理
- **免安装**：提供打包好的 exe，目标机器不需要装 Python

## 下载

到 [Releases](../../releases) 下载 `DB视图导出.exe`，与仓库里的 `.bat` 放在同一目录即可使用。

也可以克隆仓库后从源码运行，或自行打包（见下文）。

## 使用（三步）

1. 双击 `DB视图导出.bat`（或直接双击 exe）
2. 把 `.db` 文件拖进窗口，回车
   - 也可以拖 `.lnk` 快捷方式
   - 也可以拖一个文件夹，自动选里面最新的 `.db`
   - 也可以直接粘贴路径
3. 屏幕列出库里所有对象，输入序号：
   - `0` 或直接回车 = 全部导出（按行数从小到大）
   - `1`、`2`、`3`… = 只导那一个

## 命令行

```bat
DB视图导出.exe --db D:\x.db --yes                  :: 全导出，xlsx
DB视图导出.exe --db D:\x.db --yes --format csv     :: 全导出，csv
DB视图导出.exe --db D:\x.db --objects tables --yes :: 导表而不是视图
DB视图导出.exe --db D:\x.db --optimize --yes       :: 临时副本上优化后导出
DB视图导出.exe --db D:\x.db --optimize-inplace     :: 直接改这个库（需输 YES）
```

`--optimize-inplace` 会真的修改文件（建索引、ANALYZE），执行前必须输入 `YES` 确认，导出结束后工具会删除自己建的索引。**建议只在备份库上使用。**

完整参数与故障排查见 [使用说明.txt](使用说明.txt)。

## 从源码运行

```bash
pip install -r requirements.txt
python src/db_view_export.py --db D:\x.db --yes
```

## 自行打包

```bat
build.bat
```

会安装 PyInstaller 与 openpyxl，并在 `dist/` 下生成单文件 exe。构建产物不进仓库。

## 项目结构

| 路径 | 说明 |
| --- | --- |
| `src/db_view_export.py` | 全部逻辑，单文件 |
| `使用说明.txt` | 面向使用者的完整手册（215 行） |
| `DB视图导出.bat` | 启动入口，纯 ASCII |
| `调试运行.bat` | 排错用，会停住窗口并显示退出码 |
| `build.py` / `build.bat` | 打包脚本 |

## 环境

| 项目 | 要求 |
| --- | --- |
| 运行 exe | Windows 10 / 11（64 位），无需 Python |
| 从源码运行 | Python 3.8+，`openpyxl` |

## 许可

[MIT](LICENSE)
