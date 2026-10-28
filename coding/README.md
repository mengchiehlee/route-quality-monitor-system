# DSP 路区项目质量监控自动化

## 目录结构

```text
项目根目录/
├── coding/
│   ├── update_dsp_monitor.py
│   ├── config.json
│   ├── config.example.json
│   ├── requirements.txt
│   ├── run_daily_update.bat
│   └── README.md
├── input/
│   └── 20261028/
│       ├── 质量.xlsx
│       ├── 中心签出时效.csv
│       ├── 中心签入时效 (剔除).csv
│       └── 中心签入时效 (不剔除).csv
└── output/
    └── UpdateYYYYMMDD.xlsx
```

所有默认路径根据 `update_dsp_monitor.py` 自身位置计算，与启动时的工作目录无关。请保持上述目录层级，可将整个项目放在任意位置。

## 每天怎么用

1. 在 `input` 下建立当天的八位日期文件夹（YYYYMMDD），放入上述四个文件，文件名须完全一致，包括空格和括号。
2. 双击 `coding/run_daily_update.bat`。窗口会显示成功结果路径或具体错误，并保留供查看。
3. 在项目根目录的 `output` 中查看结果。

脚本仅扫描 `input` 的直接子目录，选择名称最大的有效八位日期文件夹；忽略非日期名称、无效日历日期和普通文件。最新文件夹缺少任何必需文件时，立即列出全部缺失文件并退出，不回退到旧日期。

`input` 及其数据源仅供读取，脚本不会修改、覆盖、移动或删除它们。输出统一写入 `output`，不支持命令行覆盖输入/输出位置（旧的 `--source-dir`、`--output-dir`、`--workbook`、`--center-timing`、`--signin-excluded`、`--signin-included`、`--output` 参数已移除）。

输出名中的 YYYYMMDD 为**运行当天的本机日期**，并非输入文件夹日期或业务日期。重复运行自动使用 `UpdateYYYYMMDD_1.xlsx`、`UpdateYYYYMMDD_2.xlsx` 等，不覆盖已有结果，并发运行也使用独占创建防止覆盖。

## 安装和命令行运行

安装 Python 3.10 或更高版本，并启用 Python Launcher 或将 Python 加入 PATH。在项目根目录打开终端：

```powershell
python -m pip install -r coding/requirements.txt
python coding/update_dsp_monitor.py
```

也可从任意目录使用脚本的完整路径运行。默认配置为脚本同目录的 `config.json`；`--config` 可指定其他配置。保留 `--date 2026/10/27` 指定业务日期及 `--allow-sunday` 参数，业务处理规则不变。

## 日期逻辑

脚本默认会补齐所有缺失日期：它会读取 `质量.xlsx` 的 sheet `1` 里有哪些日期，再读取 sheet `2` 里已经有哪些日期，然后把源表有、监控表还没有的日期一次性补上。

例如你一周没有运行，周一回来后只要运行一次，脚本会把这一周缺失的日期区块都插入到同一个输出文件里，最新日期会排在最上面。

如果源表最新日期是周日，脚本会跳过，不插入新行。

## 取数逻辑

蓝色指标从 `质量.xlsx` 的 sheet `1` 取数，直接照抄，不重新计算。空白写 `-`。

`2400POD` 从 `中心签出时效.csv` 取数，源日期为监控日期往前最近一个非周日、非法国节假日的日期。

`ASCAN-DSCAN 72H（剔除周日+节假日版）` 从 `中心签入时效 (剔除).csv` 取 M 列 `72H妥投率`，源日期为监控日期往前 3 个有效工作日。

`ASCAN-DSCAN 72H（不剔除周日+节假日版）` 从 `中心签入时效 (不剔除).csv` 取 M 列 `72H妥投率`，源日期为监控日期往前 3 个自然日。

绿色三列有低件量保护：`2400POD` 对应的 2400 件量小于等于 10 时写 `-`；两个 `ASCAN-DSCAN 72H` 对应的 72H 件量小于等于 10 时写 `-`。

法国节假日维护在 `config.json` 的 `france_holidays`。
