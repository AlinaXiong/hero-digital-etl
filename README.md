# hero-digital-etl · 英雄电竞数据清洗

把泛微导出的业务单据，按《业财项目_数据映射规则》清洗映射成汉得中台的期初导入模板。

## 可执行任务

查看当前所有任务：`python run.py --list`

| 任务名 | 业务含义 | 数据源 | 产出模板 |
| --- | --- | --- | --- |
| `ap_payment_opening` | 应付期初 - 对公付款单 | 对公付款主表 + 明细 | 英雄期初对公付款单导入模版 |
| `ap_payment_opening_db` | 应付期初 - 对公付款单(DB直连版) | 泛微 `uf_dgfktz` + `uf_dgfktz_dt1` | 英雄期初对公付款单导入模版 |
| `ap_payment_opening_extra_db` | 应付期初 - 对公付款单补充三 tab(DB直连版) | `ap_payment_opening_db` 口径 + 泛微 `uf_plfy` / `uf_plfy_dt1` + `uf_xtyynbsz` / `uf_xtyynbsz_dt10` / `view_costlist_ys` | 英雄期初对公付款单导入模版 |
| `ap_prepayment_opening` | 预付期初 - 供应商预付款单 + 零工预付款单 | 预付款主表 + 明细；零工平台付款头数据 + 实际收款人明细 | 英雄期初预付款单导入模版 |
| `ap_prepayment_opening_db` | 预付期初 - 供应商预付款单 + 零工预付款单(DB直连版) | 泛微 `uf_yfkxx` + `uf_yfkxx_dt1` + `uf_dgfktz_dt2`；`uf_lgptfk` + `formtable_main_279` + `formtable_main_279_dt3` + `formtable_main_279_dt4` | 英雄期初预付款单导入模版 |
| `ar_invoice_opening` | 应收期初 - 应收报账单 | 开票记录 + 收款登记 | 应收报账单期初数据导入模板 |
| `ar_invoice_opening_db` | 应收期初 - 应收报账单(DB直连版) | 泛微 `uf_xtyykp` + `uf_skdj` | 应收报账单期初数据导入模板 |
| `contract_anchor_db` | 合同迁移 - 智书主播流程(DB直连版) | 泛微 `uf_htk` + `uf_zbkp` / `uf_zbkp_dt1` | 智书合同字段-主播流程 |
| `invoice_info_db` | 发票信息(DB直连版) | 泛微 `fnainvoiceledger` + `fnainvoiceledgerdtl` | 发票信息清洗导入表 |
| `all` | 一次跑核心 DB 导入任务 | 依次执行 `ap_payment_opening_extra_db`、`ap_prepayment_opening_db`、`ar_invoice_opening_db`、`invoice_info_db` | 多个模板/清洗表 |

### all（一键执行核心 DB 任务）

按固定顺序串行执行四个任务：

1. `ap_payment_opening_extra_db`
2. `ap_prepayment_opening_db`
3. `ar_invoice_opening_db`
4. `invoice_info_db`

执行命令：

```bash
python run.py all
```

其中任一子任务失败时，进程会直接报错退出，后续任务不会继续跑。

### ap_payment_opening（应付期初 - 对公付款单）

把泛微「对公付款」单据清洗成中台的期初对公付款单导入数据。

- **源表**：`uf_dgfktz-主表.xlsx`（一行=一张付款申请单）+ `uf_dgfktz_dt1-明细表.xlsx`（一行=一条费用明细，按 ID 关联）
- **行过滤**：流程来源 ∈ {对公付款, 个人劳务付款} 且 申请日期 ≥ 2026-01-01 且 流程状态=审批完成 且 非作废
- **行粒度**：主子按 ID 合并，一行=一条费用明细
- **关键映射**：经办人→工号(泛微)、公司主体→核算主体编码(中台)、供应商→收款方编码(中台)、预算科目→费用项目编码(规则表)、付款币种→ISO；实际已支付金额按支付状态判定
- **产出**：`英雄期初对公付款单导入_应付期初_<YYYYMMDD>.xlsx`（24 列单 tab）

### ap_payment_opening_db（应付期初 - 对公付款单 DB 直连版）

与 `ap_payment_opening` 使用同一套过滤、输出和问题清单口径；源数据不读 Excel，直接从泛微库读取 `uf_dgfktz` / `uf_dgfktz_dt1`，并把人员、部门、枚举、币种、预算科目、公司主体等 ID 转成原 Excel 导出同款展示值。

### ap_payment_opening_extra_db（应付期初 - 对公付款单补充三 tab DB 直连版）

一次生成「应付期初」同一个 Excel 的三个 tab：`期初对公付款单导入`、`批量费用流程`、`只转入外部成本`。

- **期初对公付款单导入**：复用 `ap_payment_opening_db` 的读取、过滤、供应商、合同、银行账号和预算科目逻辑。
- **批量费用流程源表**：泛微 `uf_plfy` + `uf_plfy_dt1`；过滤 `d.sfqr=0`、明细未作废、记录日期 ≥ 2026-01-01。
- **只转入外部成本源表**：泛微 `uf_xtyynbsz` + `uf_xtyynbsz_dt10`，并关联 `view_costlist_ys` 取费用单明细；同时处理赛事来源 `ly=5` 和 MCN 来源 `ly=2`。
- **项目/订单字段**：先把泛微项目浏览框 ID 解析成泛微项目编号，再按 0619 项目&订单清洗表映射订单编号/订单名称，并保留 `泛微项目编号`。
- **校验清单**：除必输字段、供应商、银行账号、项目订单映射异常外，`只转入外部成本` 还会输出每个单据转入/转出正负金额是否配平的检查结果。
- **产出**：`英雄期初对公付款单导入_应付期初_补充_<YYYYMMDD>.xlsx`。

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

DB 直连版源数据不读 Excel；供应商预付从泛微 `uf_yfkxx` / `uf_yfkxx_dt1` 读取，并按「预付单 + 预算科目」关联 `uf_dgfktz_dt2` 汇总对公付款冲销金额。已到票核销金额取 `uf_dgfktz_dt2.cxje` 转正后的汇总金额；已付未核 = `uf_yfkxx_dt1.yfje` - 已到票核销金额。同一预付单同一预算科目出现多条预付明细时，按预付金额占比分摊冲销金额并做尾差调整，保证该组核销合计等于冲销表汇总。零工预付从建模头表 `uf_lgptfk` 关联原流程主表 `formtable_main_279`，再取预算项明细表 `formtable_main_279_dt3` 和收款人明细表 `formtable_main_279_dt4`；其中 `dt3` 对应「对公&报销&零工&批量四合一」里零工平台付款的预算科目/费用金额来源。人员、公司主体、合同、银行账号、币种、预算科目、供应商等 ID 在任务内批量解析；订单编号/订单名称统一按 0619 项目&订单清洗表映射，并保留 `泛微项目编号`。

### ar_invoice_opening（应收期初 - 应收报账单）

把泛微「开票记录」清洗成中台的应收报账单期初导入数据，并用「收款登记」按单号汇总补核销金额。

- **源表**：`uf_xtyykp开票.xlsx`（一行=一条开票记录）+ `uf_skdj收款登记.xlsx`（按「开票/预收单号」汇总）
- **行过滤**：申请日期 ≥ 2026-01-01 且 开票状态=已开票 且 非作废
- **行粒度**：一行=一条开票记录；收款登记按「开票/预收单号=流程编号」聚合已收款金额后回填核销金额
- **关键映射**：申请人→工号、公司主体→核算主体编码、客户→付款对象编码、业务类型→`HERO.BUSINESS_TYPE` 编码、开票类型默认合同开票、税率→`hfbs_tax_type.description`；项目/订单按 0619 项目&订单清洗表映射，并保留 `泛微项目编号`
- **产出**：`英雄应收报账单期初数据导入_应收期初_<YYYYMMDD>.xlsx`（71 列单 tab）

### ar_invoice_opening_db（应收期初 - 应收报账单 DB 直连版）

与 `ar_invoice_opening` 使用同一套过滤、输出和问题清单口径；源数据不读 Excel，直接从泛微库读取 `uf_xtyykp`，并按「开票/预收单号」关联 `uf_skdj` 汇总已收款金额。申请人、部门、公司主体、客户、合同、币种、项目等 ID 在公共方法里批量解析；项目/订单按 0619 项目&订单清洗表映射，并保留 `泛微项目编号`。

### invoice_info_db（发票信息 DB 直连版）

按规则表「发票信息」生成发票信息清洗结果。

- **源表**：泛微 `fnainvoiceledger`，并关联 `fnainvoiceledgerdtl` 汇总发票备注。
- **行过滤**：当前只取 2026 年报销/关联数据；保留 `status IN (1, 2)` 的冻结/核销状态发票，不取初始未使用发票。
- **关键映射**：发票归属人→工号、购买方→核算主体编码、泛微发票类型→汉得 `VAT_INVOICE_TYPE`，含税金额转中文大写。
- **产出**：`发票信息清洗_发票信息_2026_<YYYYMMDD>.xlsx`。

### contract_anchor_db（合同迁移 - 智书主播流程 DB 直连版）

按法务映射规则和「智书合同字段-主播流程」模板，把泛微主播合同库清洗成智书主播流程导入数据。

- **源表**：泛微 `uf_htk` 主表，关联主播卡片 `uf_zbkp` 和平台/房间明细 `uf_zbkp_dt1`。
- **行过滤**：合同类型=主播协议，合同签署状态 ∈ {审批中, 审批完成, 已归档}。
- **输出 sheet**：`字段模板`、`对方信息`、`我方信息`、`费用明细`；`选项` sheet 保留模板原样。
- **关键映射**：合同执行人、合同状态/二级类型/所属平台枚举、主播身份证/战队/签约金等从主播卡片补充；对方主体按客户/供应商分别映射到中台编码；我方主体按合同用印范围映射到核算主体编码。
- **默认值**：计价方式=固定总价，合同期限类型=固定期限，是否需要验收=否，打印模式=黑白双面打印，签约形式=纸质签约-不限制我方/对方先签约，盖章份数=3。
- **产出**：`output/contract_anchor_db/智书合同字段_主播流程_合同迁移_<YYYYMMDD>.xlsx`。

## 公共清洗口径

### 0619 项目/订单清洗映射

项目和订单字段统一从 `data/source/other_cleaned_data/业财项目_项目&订单清洗_0619.xlsx` 读取；如文件放在其他位置，可通过环境变量 `PROJECT_ORDER_MAPPING_XLSX` 指定完整路径。

- **使用 sheet**：`全量项目_清洗后` + `全量订单主表_清洗后`。
- **公共方法**：所有提取对应关系的逻辑都放在 `etl/common.py`，任务文件只调用 `c.project_order_mapping_value(...)` 和 `c.collect_order_mapping_issues(...)`。
- **原泛微项目编码拆分**：`原泛微项目编码` 可能一格维护多个编码，按分号、中文分号、逗号、中文逗号、换行拆分。
- **无优先级规则**：不区分“单独一行”和“集合里的一项”的优先级；同一个泛微项目编号映射到多个订单时，不强行填订单字段，统一列到 `订单映射_多候选`。
- **一对一规则**：只有当一个泛微项目编号最终只对应一个订单时，才回填订单编号、订单名称，以及需要时的清洗后项目编号/项目名称。
- **异常清单**：无法映射的项目进入 `订单映射_未匹配`；映射表中出现过但没有可用订单编号的项目会标明出现位置和订单字段值。
- **当前使用任务**：`ap_prepayment_opening_db`、`ap_payment_opening_extra_db`、`ar_invoice_opening`、`ar_invoice_opening_db`。

### 银行账号

供应商银行账号统一按 Hand 供应商主数据校验：

- 源单有银行账号，且该账号在 Hand 中属于当前收款方：使用源账号对应的银行账号。
- 源单未填，或源账号不属于当前收款方：使用 Hand 中该供应商 `是否默认账户=是` 的银行账号。
- 异常会进入 `银行账号_校验异常`，便于检查供应商缺账号、默认账号缺失或源账号归属不一致。

### 泛微费用项目编码

应付/预付相关导入表会在最后保留 `泛微费用项目编码`，用于回看泛微原预算科目层级，格式保持为原路径，例如：

```text
AR日常运营费用/AR4日常运营费用/AR47办公杂费
```

应收期初导入模板不需要该字段，因此应收任务不输出 `泛微费用项目编码`。

## 当前清洗进度（2026-06-20）

### 应付期初 - 对公付款单补充三 tab

- **执行命令**：`python run.py ap_payment_opening_extra_db`
- **输出结果**：同一个 Excel 写入 `期初对公付款单导入`、`批量费用流程`、`只转入外部成本` 三个 tab
- **订单映射结果**：
  - 期初对公付款单导入：订单编号已填 3163/5192
  - 批量费用流程：订单编号已填 23613/37116
  - 只转入外部成本：订单编号已填 326/610
- **补充校验**：`只转入外部成本` 同时处理赛事 `ly=5` 和 MCN `ly=2`，并输出每个单据转入/转出正负金额配平检查
- **产出文件**：`output/ap_payment_opening_extra_db/英雄期初对公付款单导入_应付期初_补充_20260620.xlsx`
- **未匹配清单**：`output/ap_payment_opening_extra_db/未匹配清单_应付期初_补充_20260620.xlsx`

### 预付期初 - 供应商预付款 / 零工预付款 DB

- **执行命令**：`python run.py ap_prepayment_opening_db`
- **供应商预付款订单映射**：订单编号已填 1181/7718
- **零工预付款订单映射**：订单编号已填 1653/13690
- **产出文件**：`output/ap_prepayment_opening_db/英雄期初预付款单导入_预付期初_20260620.xlsx`
- **未匹配清单**：`output/ap_prepayment_opening_db/未匹配清单_预付期初_20260620.xlsx`

### 应收期初 - 应收报账单

- **执行命令**：`python run.py ar_invoice_opening` / `python run.py ar_invoice_opening_db`
- **文件版项目订单清单**：多候选 74 条，未匹配 290 条
- **DB 版项目订单清单**：多候选 84 条，未匹配 307 条
- **文件版产出**：`output/ar_invoice_opening/英雄应收报账单期初数据导入_应收期初_20260620.xlsx`
- **DB 版产出**：`output/ar_invoice_opening_db/英雄应收报账单期初数据导入_应收期初_20260620.xlsx`

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
│       ├── ap_payment_opening_extra_db.py # 应付期初 - 对公付款单补充三 tab(DB直连版)
│       ├── ap_prepayment_opening.py    # 预付期初 - 供应商预付款单
│       ├── ap_prepayment_opening_db.py # 预付期初 - 供应商预付款单/零工预付款单(DB直连版)
│       ├── ar_invoice_opening.py       # 应收期初 - 应收报账单
│       ├── ar_invoice_opening_db.py    # 应收期初 - 应收报账单(DB直连版)
│       └── invoice_info_db.py          # 发票信息(DB直连版)
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
python run.py all                       # 一次跑核心 DB 导入任务
# 或单独执行一个任务:
python run.py ap_payment_opening_extra_db
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
