# hero-digital-etl · 英雄电竞数据清洗

把泛微导出的业务单据，按《业财项目_数据映射规则》清洗映射成汉得中台期初导入模板。

当前已实现任务：`ap_opening_payment`，即 **应付期初-对公付款单**。

## 目录结构

```text
hero-digital-etl/
├── run.py                              # 任务入口：python run.py <任务名>
├── requirements.txt                    # Python 依赖
├── .env.example                        # 数据库环境变量模板
├── etl/
│   ├── common.py                       # 公共能力：路径、数据库、映射、归一化、Excel 读写
│   └── tasks/
│       └── ap_opening_payment.py       # 应付期初-对公付款单清洗任务
├── data/
│   ├── source/
│   │   └── ap_opening_payment/         # 当前任务源表
│   │       ├── uf_dgfktz-主表.xlsx
│   │       └── uf_dgfktz_dt1-明细表.xlsx
│   ├── rules/
│   │   └── 业财项目_数据映射规则.xlsx
│   └── templates/
│       └── ap_opening_payment/         # 当前任务导入模板
│           └── 英雄期初对公付款单导入模版.xlsx
└── output/
    └── ap_opening_payment/             # 当前任务产出文件
        ├── 英雄期初对公付款单导入_应付期初_<YYYYMMDD>.xlsx
        └── 未匹配清单_应付期初_<YYYYMMDD>.xlsx
```

约定：每个清洗任务都使用同一个任务名作为目录名，`data/source/`、`data/templates/`、`output/` 下都要建同名文件夹，避免后续任务多了之后找不到对应 Excel。

产出文件统一使用运行当天日期后缀，不再使用固定版本号，例如 `_20260614.xlsx`。

## 快速开始

```bash
pip install -r requirements.txt
copy .env.example .env
python run.py ap_opening_payment
```

也可以查看当前可用任务：

```bash
python run.py --list
```

`.env` 读取真实数据库连接信息，本地使用即可，不要提交到 GitHub。脚本只执行 `SELECT` 查询，不写入数据库。

## 当前任务

任务名：`ap_opening_payment`

入口文件：`etl/tasks/ap_opening_payment.py`

运行方式：

```bash
python run.py ap_opening_payment
```

直接运行任务文件也支持：

```bash
python etl/tasks/ap_opening_payment.py
```

数据口径：

- 流程来源：`对公付款`、`个人劳务付款`
- 申请日期：`>= 2026-01-01`
- 流程状态：`审批完成`
- 作废单据：剔除
- 行粒度：主表和明细表按 `ID` 合并，一行对应一条费用明细

主要映射：

| 模板字段 | 取数来源 |
|---|---|
| 来源系统 | 固定 `FW` |
| 来源单据编号、申请日期、备注、合同号、银行账号、计划付款日期 | 泛微主表 |
| 单据类型 | 固定 `AP01-1` |
| 申请人工号 | 泛微 `vspn_xtyy.hrmresource.WORKCODE` |
| 核算主体编号 | 中台 `hfins_base_account.hfac_accounting_entity.acc_entity_code`，仅按 `acc_entity_name` 建映射 |
| 收款方编码 | 中台 `hfbs_system_vender.vender_code` |
| 费用项目编码、费用项目描述 | 预算科目映射规则 |
| 实际已支付金额 | 支付状态为已支付时取本行报账金额，否则为 `0` |
| 报账币种 | 付款币种转 ISO 码 |
| 报账金额 | 明细表付款金额 |

## 新增任务规范

1. 在 `etl/tasks/` 下新建任务文件，文件名使用清晰英文名，不使用拼音。
2. 在 `data/source/<任务名>/` 放源表。
3. 在 `data/templates/<任务名>/` 放导入模板。
4. 任务运行后输出到 `output/<任务名>/`。
5. 在 `run.py` 的 `TASKS` 字典登记任务名。

公共逻辑尽量复用 `etl/common.py`，例如数据库连接、工号映射、供应商映射、核算主体映射、科目映射、币种转换、模板写入和未匹配清单输出。

## 代码规范

- 代码注释和说明性 docstring 统一写中文，方便业务同学核对口径。
- 文件名、目录名、函数名、变量名、常量名必须使用英文，不使用中文，也不使用拼音。
- 业务字段名、Excel 表头、sheet 名和文件展示名可以保留来源系统或模板里的中文。
- 新任务命名要能直接表达业务含义，例如 `ap_opening_payment`，不要使用 `qichu`、`yingfu` 这类拼音。

## 排查输出

任务会额外生成 `output/ap_opening_payment/未匹配清单_应付期初_<YYYYMMDD>.xlsx`，用于核对以下问题：

- 未匹配工号
- 未匹配供应商
- 未匹配核算主体
- 未匹配费用科目
- 分组合并待核对数据
