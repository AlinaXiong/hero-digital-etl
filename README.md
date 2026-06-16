# hero-digital-etl · 英雄电竞数据清洗

把泛微导出的业务单据，按《业财项目_数据映射规则》清洗映射成汉得中台的期初导入模板。

## 可执行任务

查看当前所有任务：`python run.py --list`

| 任务名 | 业务含义 | 数据源                                                                                           | 产出模板 |
| --- | --- |-----------------------------------------------------------------------------------------------| --- |
| `ap_payment_opening` | 应付期初 - 对公付款单 | 对公付款主表 + 明细                                                                                   | 英雄期初对公付款单导入模版 |
| `ap_payment_opening_db` | 应付期初 - 对公付款单(DB直连版) | 泛微 `uf_dgfktz` + `uf_dgfktz_dt1`                                                              | 英雄期初对公付款单导入模版 |
| `ap_prepayment_opening` | 预付期初 - 供应商预付款单 + 零工预付款单 | 预付款主表 + 明细；零工平台付款头数据 + 实际收款人明细                                                                | 英雄期初预付款单导入模版 |
| `ap_prepayment_opening_db` | 预付期初 - 供应商预付款单 + 零工预付款单(DB直连版) | 泛微 `uf_yfkxx` + `uf_yfkxx_dt1`；`uf_lgptfk` +  `formtable_main_279` + `formtable_main_279_dt4` | 英雄期初预付款单导入模版 |
| `ar_invoice_opening` | 应收期初 - 应收报账单 | 开票记录 + 收款登记                                                                                   | 应收报账单期初数据导入模板 |
| `ar_invoice_opening_db` | 应收期初 - 应收报账单(DB直连版) | 泛微 `uf_xtyykp` + `uf_skdj`                                                                    | 应收报账单期初数据导入模板 |

### ap_payment_opening（应付期初 - 对公付款单）

把泛微「对公付款」单据清洗成中台的期初对公付款单导入数据。

- **源表**：`uf_dgfktz-主表.xlsx`（一行=一张付款申请单）+ `uf_dgfktz_dt1-明细表.xlsx`（一行=一条费用明细，按 ID 关联）
- **行过滤**：流程来源 ∈ {对公付款, 个人劳务付款} 且 申请日期 ≥ 2026-01-01 且 流程状态=审批完成 且 非作废
- **行粒度**：主子按 ID 合并，一行=一条费用明细
- **关键映射**：经办人→工号(泛微)、公司主体→核算主体编码(中台)、供应商→收款方编码(中台)、预算科目→费用项目编码(规则表)、付款币种→ISO；实际已支付金额按支付状态判定
- **产出**：`英雄期初对公付款单导入_应付期初_<YYYYMMDD>.xlsx`（24 列单 tab）

### ap_payment_opening_db（应付期初 - 对公付款单 DB 直连版）

与 `ap_payment_opening` 使用同一套过滤、输出和问题清单口径；源数据不读 Excel，直接从泛微库读取 `uf_dgfktz` / `uf_dgfktz_dt1`，并把人员、部门、枚举、币种、预算科目、公司主体等 ID 转成原 Excel 导出同款展示值。

### ap_prepayment_opening（预付期初 - 供应商预付款单 / 零工预付款单）

把泛微「预付款」和「零工平台付款」单据清洗成中台的期初预付款导入数据，一次写入模板的两个 tab。

- **供应商预付款源表**：`uf_yfkxx预付.xlsx`（一行=一张预付款单）+ `uf_yfkxx_dt1.xlsx`（一行=一条预付预算明细，按 ID 关联）
- **供应商预付款过滤**：申请日期 ≥ 2026-01-01 且 流程状态=审批完成 且 非作废
- **供应商预付款粒度**：主子按 ID 合并，一行=一条费用明细
- **供应商预付款关键映射**：填单人→工号、开票单位→核算主体编码、付款对象→收款方编码、预算科目→费用项目编码、付款币种→ISO；付款性质(押金/质保金)→保证金标志
- **供应商预付款金额拆分**：预付款金额取明细金额；已付未核 = 主表剩余冲销/退款金额按明细占比分摊；已到票核销 = 预付款金额 − 已付未核。这样处理一张主单多条预算明细时，费用行合计仍等于主表金额。
- **零工预付款源表**：`零工平台付款_收款人明细_2026.xlsx`，包含「付款头数据」+「实际收款人明细」两个 sheet，按建模付款 ID 关联
- **零工预付款过滤**：付款头数据流程状态=2(审批完成) 且 申请日期 ≥ 2026-01-01 且 非作废
- **零工预付款关键映射**：经办人工号直接取源表；公司主体ID→泛微 `uf_gstt.gsmc`→中台核算主体编码；收款方文本→灵工平台收款方编码；实际收款方→中台供应商编码，若中台暂未建档则保留实际收款方原值
- **产出**：`英雄期初预付款单导入_预付期初_<YYYYMMDD>.xlsx`，写入「期初供应商预付款单&期初投资付款单导入」和「期初灵工预付款单导入」两个 tab

### ap_prepayment_opening_db（预付期初 - 供应商预付款单 / 零工预付款单 DB 直连版）

与 `ap_prepayment_opening` 使用同一套过滤、输出和问题清单口径；源数据不读 Excel，供应商预付直接从泛微 `uf_yfkxx` / `uf_yfkxx_dt1` 读取，零工预付从建模头表 `uf_lgptfk` 关联原流程主表 `formtable_main_279` 和收款人明细表 `formtable_main_279_dt4` 读取。人员、公司主体、合同、银行账号、币种、预算科目、供应商等 ID 在任务内批量解析。

### ar_invoice_opening（应收期初 - 应收报账单）

把泛微「开票记录」清洗成中台的应收报账单期初导入数据，并用「收款登记」按单号汇总补核销金额。

- **源表**：`uf_xtyykp开票.xlsx`（一行=一条开票记录）+ `uf_skdj收款登记.xlsx`（按「开票/预收单号」汇总）
- **行过滤**：申请日期 ≥ 2026-01-01 且 开票状态=已开票 且 非作废
- **行粒度**：一行=一条开票记录；收款登记按「开票/预收单号=流程编号」聚合已收款金额后回填核销金额
- **关键映射**：申请人→工号、公司主体→核算主体编码、客户→付款对象编码、业务类型→`HERO.BUSINESS_TYPE` 编码、开票类型默认合同开票、税率→`hfbs_tax_type.description`
- **产出**：`英雄应收报账单期初数据导入_应收期初_<YYYYMMDD>.xlsx`（71 列单 tab）

### ar_invoice_opening_db（应收期初 - 应收报账单 DB 直连版）

与 `ar_invoice_opening` 使用同一套过滤、输出和问题清单口径；源数据不读 Excel，直接从泛微库读取 `uf_xtyykp`，并按「开票/预收单号」关联 `uf_skdj` 汇总已收款金额。申请人、部门、公司主体、客户、合同、币种等 ID 在公共方法里批量解析。

## 当前清洗进度（2026-06-14）

### 应付期初 - 供应商付款

- **执行命令**：`python run.py ap_payment_opening`
- **过滤结果**：满足流程来源/申请日期/审批状态条件 3885 单，剔除作废 8 单，最终保留主表 3877 单
- **输出结果**：生成导入明细 4383 行
- **待业务确认/补充**：
  - 订单编号：暂未映射，0/4383
  - 收款方编码：已匹配 4116/4383，未匹配清单 156 条
  - 费用项目编码：已匹配 4293/4383，未匹配清单 90 条
- **产出文件**：`output/ap_payment_opening/英雄期初对公付款单导入_应付期初_20260614.xlsx`
- **未匹配清单**：`output/ap_payment_opening/未匹配清单_应付期初_20260614.xlsx`

### 预付期初 - 供应商预付款 / 零工预付款

- **执行命令**：`python run.py ap_prepayment_opening`
- **供应商预付款过滤结果**：满足申请日期/审批状态条件 1023 单，剔除作废 12 单，最终保留主表 1011 单
- **供应商预付款输出结果**：生成导入明细 1087 行
- **零工预付款过滤结果**：满足申请日期/审批状态条件 962 单，剔除作废 0 单，最终保留主表 962 单
- **零工预付款输出结果**：生成导入明细 2731 行
- **待业务确认/补充**：
  - 供应商预付款订单编号：暂未映射，0/1087
  - 供应商预付款收款方编码：已匹配 1057/1087，未匹配清单 30 条
  - 供应商预付款费用项目编码：已匹配 1030/1087，未匹配清单 57 条
  - 零工预付款订单编号：暂未映射，0/2731
  - 零工预付款费用项目编码：规则目前无来源，0/2731
- **产出文件**：`output/ap_prepayment_opening/英雄期初预付款单导入_预付期初_20260614.xlsx`
- **未匹配清单**：`output/ap_prepayment_opening/未匹配清单_预付期初_20260614.xlsx`

### 应收期初 - 应收报账单（2026-06-15）

- **执行命令**：`python run.py ar_invoice_opening`
- **过滤结果**：满足申请日期/开票状态条件 823 行，剔除作废 0 行，最终保留开票记录 823 行
- **输出结果**：生成导入明细 823 行
- **待业务确认/补充**：
  - 付款对象：已匹配 818/823，未匹配清单 5 条（源表客户为空）
  - 合同编号：已填 641/823，缺失 182 行（未匹配清单按来源单据去重后 157 条）
  - 税率类型：已匹配 812/823，缺失 11 行（源表税率为空或未命中字典）
  - 里程碑阶段、平台、自审批、自审核、凭证推送、行号、收入分类：规则标注由汉得后续统一赋值，当前保持为空
- **产出文件**：`output/ar_invoice_opening/英雄应收报账单期初数据导入_应收期初_20260615.xlsx`
- **未匹配清单**：`output/ar_invoice_opening/未匹配清单_应收期初_20260615.xlsx`

## 目录结构

```text
hero-digital-etl/
├── run.py                              # 任务入口：python run.py <任务名>
├── requirements.txt                    # Python 依赖
├── .env.example                        # 数据库环境变量模板
├── etl/
│   ├── common.py                       # 公共能力：路径/数据库/各类映射/归一化/Excel读写/过滤统计
│   └── tasks/
│       ├── ap_payment_opening.py       # 应付期初 - 对公付款单
│       ├── ap_payment_opening_db.py    # 应付期初 - 对公付款单(DB直连版)
│       ├── ap_prepayment_opening.py    # 预付期初 - 供应商预付款单
│       ├── ap_prepayment_opening_db.py # 预付期初 - 供应商预付款单/零工预付款单(DB直连版)
│       ├── ar_invoice_opening.py       # 应收期初 - 应收报账单
│       └── ar_invoice_opening_db.py    # 应收期初 - 应收报账单(DB直连版)
├── data/
│   ├── source/<任务名>/                # 各任务源表(文件名保持来源系统原名)
│   ├── rules/业财项目_数据映射规则.xlsx
│   └── templates/<任务名>/             # 各任务导入模板
└── output/<任务名>/                    # 各任务产出(导入文件 + 未匹配清单)
```

约定：每个任务用同一个任务名作为目录名，`data/source/`、`data/templates/`、`output/` 下都建同名文件夹。产出文件统一用运行当天日期后缀，如 `_20260614.xlsx`。

## 快速开始

```bash
pip install -r requirements.txt
copy .env.example .env                  # 填入真实数据库账密
python run.py ap_payment_opening        # 或 ap_prepayment_opening / ar_invoice_opening
```

`.env` 读取真实数据库连接信息，本地使用即可，不要提交到 GitHub。脚本只执行 `SELECT` 查询，不写入数据库。

数据库访问统一走 SQLAlchemy。调试时可打印已代入参数、可直接复制到 MySQL 执行的 SQL：

```bash
SQL_ECHO=1 python run.py ap_payment_opening_db
```

如需 SQLAlchemy 原生日志，可使用 `SQLALCHEMY_ECHO=1`。

## 泛微字段含义查询

泛微表字段含义以 `workflow_bill` / `workflow_billfield` / `htmllabelinfo` 为准。代码里优先用公共方法：

```python
from etl import common as c

field_df = c.read_fw_field_dictionary('uf_dgfktz')
```

返回列使用代码友好的英文名：`field_id`、`field_name`、`label_name`、`field_db_type`、`field_html_type`、`field_type`、`detail_table`、`display_order`。

主表和明细表字段都从同一个建模表名查，明细字段用 `detail_table` 区分。底层 SQL 口径如下：

```sql
SELECT
    f.id AS field_id,
    f.fieldname AS field_name,
    l.labelname AS label_name,
    f.fielddbtype AS field_db_type,
    f.fieldhtmltype AS field_html_type,
    f.type AS field_type,
    f.detailtable AS detail_table,
    f.dsporder AS display_order
FROM workflow_bill b
JOIN workflow_billfield f
    ON f.billid = b.id
LEFT JOIN htmllabelinfo l
    ON l.indexid = f.fieldlabel
   AND l.languageid = 7
WHERE b.tablename = 'uf_dgfktz';
```

## 排查输出（未匹配清单）

每个任务除导入文件外，还会生成 `未匹配清单_<业务>_<YYYYMMDD>.xlsx`，内容由规则表「是否必填=Y」驱动，不写死字段：

- **必输字段未达100%**：所有必输字段中填充率 < 100% 的汇总（字段、缺失数、填充率、备注）。任何必输字段掉到 100% 以下都会自动出现；若规则备注写明“无需填写”且字段整列为空，备注列标注“无需填写”。
- **缺失_<字段>**：对每个【部分缺失】的必输字段，生成一张缺失明细，列出缺该字段的记录（去重的标识列：来源单据编号 + 各描述列），便于定位与补数。
- 规则：字段全部填满则不出现；字段全空（如尚未映射的订单编号）只在汇总里体现，不导出整表明细。

## 新增任务规范

1. 在 `etl/tasks/` 下新建任务文件，文件名用清晰英文名，不用拼音。
2. 在 `data/source/<任务名>/`、`data/templates/<任务名>/` 放源表与模板。
3. 在 `run.py` 的 `TASKS` 字典登记任务名。
4. 行过滤口径(`filter_main`)写在各任务文件内（各任务差异大）；公共能力复用 `etl/common.py`（数据库连接、工号/供应商/核算主体/科目映射、币种转换、归一化、必输字段识别、未匹配清单、模板写入）。

## 代码规范

- 代码注释与 docstring 用中文，方便业务核对口径。
- 文件名、目录名、函数名、变量名、常量名必须用英文，不用中文也不用拼音。
- 业务字段名、Excel 表头、sheet 名、展示用文件名可保留来源系统/模板里的中文。
- `data/source/` 下的源数据文件保持来源系统原名，不重命名，便于溯源对账。
