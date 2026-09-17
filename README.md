# SQL 字段核查与数据质量校验工具（sql-field-validator）

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://www.python.org)
[![License](https://img.shields.io/badge/license-MIT-blue)](#license)

> **开源脱敏版。** 原项目为某政务数据局的「SQL 字段核查工具」，用于比对源表与目标表字段结构、校验数据质量。本仓库已移除真实表结构 DDL（`.ddl`）、核查结果 `.xlsx` 与打包产物，仅保留通用校验脚本。

一组 Python 小工具，用于**字段级结构比对与数据质量校验**：

- `SQL字段核查工具.py`：按字段类型 / 长度 / 精度比对源表与目标表，定位结构差异
- `data_check.py`：通用数据质量检查逻辑
- `SQL字段核查工具.spec`：PyInstaller 打包配置（用于生成独立 exe）

## 技术栈

- Python 3.8+，标准库为主
- 可选 `openpyxl` / `pandas` 用于读写表格结果
- `SQL字段核查工具.spec` 配合 PyInstaller 可打包为桌面 exe

## 目录结构

```
sql-field-validator/
├── SQL字段核查工具.py     # 字段结构比对主程序
├── data_check.py          # 数据质量校验
└── SQL字段核查工具.spec   # PyInstaller 打包配置
```

## 用法

```bash
# 字段核查（具体参数见脚本内 argparse / 配置）
python SQL字段核查工具.py --source <源表DDL或连接> --target <目标表>

# 数据质量校验
python data_check.py <输入> -o 核查结果.xlsx
```

## 说明

- 原项目中的真实数据库 DDL（`.sql` / `.ddl`）、核查结果 `.xlsx` 含内部表结构，**未纳入本仓库**。
- 脚本中的表名、字段名、连接信息均已替换为占位符，请在你的环境中按真实 schema 调整。
- 如需复现具体校验，请准备你自己的表结构定义作为输入。

## License


本项目基于 [MIT License](./LICENSE) 开源，可自由使用、修改和分发。欢迎按需二次开发。
