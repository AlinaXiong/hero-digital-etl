# hero-digital-etl · 英雄电竞数据清洗

把泛微导出的业务单据，按《业财项目_数据映射规则》清洗映射成汉得中台的期初导入模板。

## 可执行任务

查看当前所有任务：`python run.py --list`

| 任务名 | 业务含义 | 数据源 | 产出模板 |
| --- | --- | --- | --- |
| `ap_payment_opening_extra_db` | 应付期初 - 对公付款单(含批量费用 / 只转入外部成本 / MCN 多 tab)(DB直连版) | 泛微 `uf_dgfktz` / `uf_dgfktz_dt1` + `uf_plfy` / `uf_plfy_dt1` + `uf_xtyynbsz` / `uf_xtyynbsz_dt10` / `view_costlist_ys` | 英雄期初对公付款单导入模版 |
| `ap_prepayment_opening_db` | 预付期初 - 供应商预付款单 + 零工预付款单(DB直连版) | 泛微 `uf_yfkxx` + `uf_yfkxx_dt1` + `uf_dgfktz_dt2`；`uf_lgptfk` + `formtable_main_279` + `formtable_main_279_dt3` + `formtable_main_279_dt4` | 英雄期初预付款单导入模版 |
| `ar_invoice_opening_db` | 应收期初 - 应收报账单(DB直连版) | 泛微 `uf_xtyykp` + `uf_skdj` | 应收报账单期初数据导入模板 |
| `contract_general_db` | 合同迁移 - 一般流程 Excel 导出 | 泛微 `uf_htk` + 项目&订单清洗表 | 智书合同字段-一般流程 |
| `contract_general_attachments_db` | 合同迁移 - 一般流程附件下载 | 泛微合同稿件字段 + `workflow_docshareinfo` / `docimagefile` | 一般流程合同附件 + 下载清单 |
| `contract_anchor_db` | 合同迁移 - 主播流程 Excel 导出 | 泛微 `uf_htk` + `uf_zbkp` / `uf_zbkp_dt1` + Hand 主播档案 | 智书合同字段-主播流程 |
| `contract_anchor_attachments_db` | 合同迁移 - 主播流程附件下载 | 主播合同清洗口径 + 泛微合同稿件字段 | 主播流程合同附件 + 下载清单 |
| `contract_anti_bribery_db` | 合同迁移 - 反商业贿赂协议 Excel 导出 | 泛微一般合同/赛事合同 + 反贿赂模板 | 反商业贿赂协议清洗结果 |
| `contract_anti_bribery_attachments_db` | 合同迁移 - 反商业贿赂协议附件下载 | 反贿赂合同清洗口径 + 泛微合同稿件字段 | 反贿赂合同附件 + 下载清单 |
| `export_feishu_employees` | 飞书全量员工信息导出（合同相关辅助任务） | 飞书 CoreHR 员工接口 | 飞书员工信息 Excel |
| `invoice_info_db` | 发票信息(DB直连版) | 泛微 `fnainvoiceledger` + `fnainvoiceledgerdtl` | 发票信息清洗导入表 |
| `all` | 一次跑核心 DB 导入任务 | 依次执行 `ap_payment_opening_extra_db`、`ap_prepayment_opening_db`、`ar_invoice_opening_db`、`invoice_info_db` | 多个模板/清洗表 |
| `contract_all` | 一次跑所有合同任务(不含附件下载) | 依次执行 `contract_general_db`、`contract_anchor_db`、`contract_anti_bribery_db` | 智书合同各模板 |
| `contract_all_with_attachments` | 一次跑所有合同任务(含附件下载) | `contract_all` 三个任务 + `contract_general_attachments_db`、`contract_anchor_attachments_db`、`contract_anti_bribery_attachments_db` | 智书合同各模板 + 合同附件 |

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

### contract_all / contract_all_with_attachments（一键执行合同任务）

合同迁移本体共有 **6 个任务**：一般流程、主播流程、反商业贿赂协议各包含一个 Excel 导出任务和一个附件下载任务。

`contract_all` 串行执行其中 3 个 Excel 导出任务，**不下载附件**：

1. `contract_general_db`（一般流程 Excel）
2. `contract_anchor_db`（主播流程 Excel）
3. `contract_anti_bribery_db`（反商业贿赂协议 Excel）

`contract_all_with_attachments` 在上面三个之后，再依次执行 3 个附件下载任务，因此会完整执行全部 6 个合同迁移任务：

4. `contract_general_attachments_db`
5. `contract_anchor_attachments_db`
6. `contract_anti_bribery_attachments_db`

先出全部导入 Excel 再下载附件：即使附件下载因 cookie 失效中断，导入 Excel 也已生成。附件下载依赖 `.env` 的 `WEAVER_CONTRACT_ATTACHMENT_COOKIE`，为空时只生成下载清单。与 `all` 一致，任一子任务失败即报错退出。

此外还有 1 个合同相关辅助任务 `export_feishu_employees`，用于将飞书 CoreHR 全量员工信息导出为 Excel，方便人工核对申请人、合同执行人及接收人的在职状态。它不计入上述 6 个合同迁移任务，也不会由 `contract_all` 或 `contract_all_with_attachments` 自动执行；合同清洗任务本身仍直接调用飞书接口获取在职状态。

执行命令：

```bash
python run.py contract_all                    # 不含附件下载
python run.py contract_all_with_attachments   # 含附件下载
python run.py export_feishu_employees          # 单独导出飞书全量员工信息
```

### ap_payment_opening_extra_db（应付期初 - 对公付款单补充三 tab DB 直连版）

一次生成「应付期初」同一个 Excel 的三个 tab：`期初对公付款单导入`、`批量费用流程`、`只转入外部成本`。

- **期初对公付款单导入**：对公付款主流程，从泛微 `uf_dgfktz` / `uf_dgfktz_dt1` 读数；读取、过滤、供应商、合同、银行账号、预算科目等逻辑已内联本模块。
- **批量费用流程源表**：泛微 `uf_plfy` + `uf_plfy_dt1`；过滤 `d.sfqr=0`、明细未作废、记录日期 ≥ 2026-01-01。
- **只转入外部成本源表**：泛微 `uf_xtyynbsz` + `uf_xtyynbsz_dt10`，并关联 `view_costlist_ys` 取费用单明细；同时处理赛事来源 `ly=5` 和 MCN 来源 `ly=2`。
- **项目/订单字段**：先把泛微项目浏览框 ID 解析成泛微项目编号，再按 0619 项目&订单清洗表映射订单编号/订单名称，并保留 `泛微项目编号`。
- **校验清单**：除必输字段、供应商、银行账号、项目订单映射异常外，`只转入外部成本` 还会输出每个单据转入/转出正负金额是否配平的检查结果。
- **产出**：`英雄期初对公付款单导入_应付期初_补充_<YYYYMMDD>.xlsx`。

### ap_prepayment_opening_db（预付期初 - 供应商预付款单 / 零工预付款单 DB 直连版）

DB 直连版源数据不读 Excel；供应商预付从泛微 `uf_yfkxx` / `uf_yfkxx_dt1` 读取，并按「预付单 + 预算科目」关联 `uf_dgfktz_dt2` 汇总对公付款冲销金额。已到票核销金额取 `uf_dgfktz_dt2.cxje` 转正后的汇总金额；已付未核 = `uf_yfkxx_dt1.yfje` - 已到票核销金额。同一预付单同一预算科目出现多条预付明细时，按预付金额占比分摊冲销金额并做尾差调整，保证该组核销合计等于冲销表汇总。零工预付从建模头表 `uf_lgptfk` 关联原流程主表 `formtable_main_279`，再取预算项明细表 `formtable_main_279_dt3` 和收款人明细表 `formtable_main_279_dt4`；其中 `dt3` 对应「对公&报销&零工&批量四合一」里零工平台付款的预算科目/费用金额来源。人员、公司主体、合同、银行账号、币种、预算科目、供应商等 ID 在任务内批量解析；订单编号/订单名称统一按 0619 项目&订单清洗表映射，并保留 `泛微项目编号`。

### ar_invoice_opening_db（应收期初 - 应收报账单 DB 直连版）

把泛微「开票记录」清洗成中台的应收报账单期初导入数据，并按「开票/预收单号」关联 `uf_skdj` 汇总已收款金额补核销；源数据不读 Excel，直接从泛微库读取 `uf_xtyykp`。

- **行过滤**：申请日期 ≥ 2026-01-01 且 开票状态=已开票 且 非作废
- **行粒度**：一行=一条开票记录；收款登记按「开票/预收单号=流程编号」聚合已收款金额后回填核销金额
- **关键映射**：申请人→工号、公司主体→核算主体编码、客户→付款对象编码、业务类型→`HERO.BUSINESS_TYPE` 编码、开票类型默认合同开票、税率→`hfbs_tax_type.description`；申请人/部门/公司主体/客户/合同/币种/项目等 ID 在公共方法里批量解析；项目/订单按 0619 项目&订单清洗表映射，并保留 `泛微项目编号`
- **产出**：`英雄应收报账单期初数据导入_应收期初_<YYYYMMDD>.xlsx`（71 列单 tab）

### invoice_info_db（发票信息 DB 直连版）

按规则表「发票信息」生成发票信息清洗结果。

- **源表**：泛微 `fnainvoiceledger`，并关联 `fnainvoiceledgerdtl` 汇总发票备注。
- **行过滤**：当前只取 2026 年报销/关联数据；保留 `status IN (1, 2)` 的冻结/核销状态发票，不取初始未使用发票。
- **关键映射**：发票归属人→工号、购买方→核算主体编码、泛微发票类型→汉得 `VAT_INVOICE_TYPE`，含税金额转中文大写。
- **产出**：`发票信息清洗_发票信息_2026_<YYYYMMDD>.xlsx`。

### contract_anchor_db（合同迁移 - 智书主播流程 DB 直连版）

按法务映射规则和「智书合同字段-主播流程」模板，把泛微主播合同库清洗成智书主播流程导入数据。

- **源表**：泛微 `uf_htk` 主表，关联主播卡片 `uf_zbkp` 和平台/房间明细 `uf_zbkp_dt1`。
- **行过滤**：合同类型=主播协议，合同签署状态 ∈ {审批中, 审批完成, 已归档}；排除 `our_party_code（我方主体编码）` 对应“苏州厚音文化传媒有限公司”的合同。Hand 生产环境查不到主播时，仅保留合同结束年份为 2026–2031 且泛微合同状态为“归档、上传电子档、法务确认、用印、申请人确定签约性质（兼容泛微节点名‘申请人确认签约性质’）”之一的合同，并按《主播替换.xlsx》替换主播 ID。
- **输出 sheet**：`字段模板`、`对方信息`、`我方信息`、`费用明细`；`选项` sheet 保留模板原样。
- **关键映射**：合同执行人、合同状态/二级类型/所属平台枚举、主播身份证/战队/签约金等从主播卡片补充；对方主体按客户/供应商分别映射到中台编码；我方主体按合同用印范围映射到核算主体编码。
- **默认值**：计价方式=固定总价，合同期限类型=固定期限，是否需要验收=否，打印模式=黑白双面打印，签约形式=纸质签约-不限制我方/对方先签约，盖章份数=3。
- **产出**：`output/contract_anchor_db/智书合同字段_主播流程_合同迁移_<YYYYMMDD>.xlsx`。

### contract_general_db（合同迁移 - 智书一般流程 DB 直连版）

按法务映射规则和「智书合同字段-一般流程」模板，把泛微非主播合同库清洗成智书一般流程导入数据。

- **源表**：泛微 `uf_htk` 主表；项目/订单字段按公共项目&订单清洗表映射。
- **行过滤**：排除合同类型=主播协议，保留合同签署状态 ∈ {审批完成, 已归档}。
- **输出 sheet**：`字段模板`、`关联合同`、`相关单据-订单信息`、`采购申请`、`订单信息明细`、`对方信息`、`我方主体列表`、`付款计划`、`收款计划`、`合同附件`、`其他附件`；`选项` 和 `DropdownOptions` 保留模板原样。
- **关键映射**：合同二级分类按《合同数据迁移-二级分类映射规则》由合同编号前缀、项目/标题关键词和主体信息推导；合同执行人取工号；对方/我方主体映射到中台编码；订单字段按项目&订单清洗表一对一回填。
- **合同附件**：`合同附件` sheet 只写附件名称，不执行文件下载；附件下载单独运行 `python run.py contract_general_attachments_db`。
- **产出**：`output/contract_general_db/智书合同字段_一般流程_<YYYYMMDD>.xlsx`。

### contract_general_attachments_db（一般流程合同附件下载 DB 直连版）

单独下载一般流程合同附件，不影响 `contract_general_db` 的 Excel 生成。

- **附件来源**：`uf_htk.htqdg` 作为合同生效稿兜底；结合 `workflow_docshareinfo` 最终节点文档补充合同初稿/合同签署稿。
- **目录结构**：`output/contract_general_db/一般流程合同附件_<YYYYMMDD>/<合同编码>/合同初稿|合同签署稿|合同生效稿/附件文件`。
- **配置**：`.env` 填 `WEAVER_CONTRACT_ATTACHMENT_COOKIE` 后运行；为空时只生成下载清单。
- **产出清单**：`output/contract_general_db/一般流程合同附件下载清单_<YYYYMMDD>.xlsx`。

## 公共清洗口径

### 0619 项目/订单清洗映射

项目和订单字段统一从 `resources/source/other_cleaned_data/业财项目_项目&订单清洗_0619.xlsx` 读取；如文件放在其他位置，可通过环境变量 `PROJECT_ORDER_MAPPING_XLSX` 指定完整路径。

- **使用 sheet**：`全量项目_清洗后` + `全量订单主表_清洗后`。
- **公共方法**：所有提取对应关系的逻辑都放在 `etl/util/common.py`，任务文件只调用 `c.project_order_mapping_value(...)` 和 `c.collect_order_mapping_issues(...)`。
- **原泛微项目编码拆分**：`原泛微项目编码` 可能一格维护多个编码，按分号、中文分号、逗号、中文逗号、换行拆分。
- **无优先级规则**：不区分“单独一行”和“集合里的一项”的优先级；同一个泛微项目编号映射到多个订单时，不强行填订单字段，统一列到 `订单映射_多候选`。
- **一对一规则**：只有当一个泛微项目编号最终只对应一个订单时，才回填订单编号、订单名称，以及需要时的清洗后项目编号/项目名称。
- **异常清单**：无法映射的项目进入 `订单映射_未匹配`；映射表中出现过但没有可用订单编号的项目会标明出现位置和订单字段值。
- **当前使用任务**：`ap_prepayment_opening_db`、`ap_payment_opening_extra_db`、`ar_invoice_opening_db`。

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

- **执行命令**：`python run.py ar_invoice_opening_db`
- **项目订单清单**：多候选 84 条，未匹配 307 条
- **产出**：`output/ar_invoice_opening_db/英雄应收报账单期初数据导入_应收期初_20260620.xlsx`

## 目录结构

```text
hero-digital-etl/
├── run.py                              # 任务入口：python run.py <任务名>
├── requirements.txt                    # Python 依赖
├── .env.example                        # 数据库环境变量模板
├── etl/
│   ├── util/                           # 公共能力
│   │   ├── common.py                   #   路径/数据库/各类映射/归一化/Excel读写/过滤统计
│   │   └── flatten_project_codes.py    #   项目编码拆分小工具
│   ├── contract/                       # 合同迁移
│   │   ├── contract_general_db.py      #   一般流程(DB直连版)
│   │   ├── contract_general_attachments_db.py
│   │   ├── contract_anchor_db.py       #   主播流程(DB直连版)
│   │   ├── contract_anchor_attachments_db.py
│   │   ├── contract_anti_bribery_db.py #   反商业贿赂协议(DB直连版)
│   │   ├── contract_anti_bribery_attachments_db.py
│   │   └── anti_bribery_signers_db.py  #   反商业贿赂协议签署情况补登
│   ├── process/                        # 期初流程
│   │   ├── ap_payment_opening_extra_db.py #   应付期初 - 对公付款单 / 补充三 tab / MCN
│   │   ├── ap_prepayment_opening_db.py #   预付期初 - 供应商/零工预付款单
│   │   └── ar_invoice_opening_db.py    #   应收期初 - 应收报账单
│   ├── invoice/
│   │   └── invoice_info_db.py          #   发票信息(DB直连版)
│   └── lark/
│       ├── feishu.py                   #   飞书(Lark)客户端
│       └── export_feishu_employees.py  #   全量员工信息导出Excel
├── resources/
│   ├── source/<任务名>/                # 各任务源表(文件名保持来源系统原名)
│   ├── rules/业财项目_数据映射规则.xlsx
│   ├── templates/<任务名>/             # 各任务导入模板
│   └── reference/                      # 字段字典等参考资料(ETL 不读)
└── output/<任务名>/                    # 各任务产出(导入文件 + 未匹配清单)
```

约定：每个任务用同一个任务名作为目录名，`resources/source/`、`resources/templates/`、`output/` 下都建同名文件夹。产出文件统一用运行当天日期后缀，如 `_20260614.xlsx`。

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
SQL_ECHO=1 python run.py ap_payment_opening_extra_db
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

1. 在 `etl/<分组>/`（contract / process / invoice / lark）下新建任务文件，文件名用清晰英文名，不用拼音。
2. 在 `resources/source/<任务名>/`、`resources/templates/<任务名>/` 放源表与模板。
3. 在 `run.py` 的 `TASKS` 字典登记任务名。
4. 行过滤口径(`filter_main`)写在各任务文件内（各任务差异大）；公共能力复用 `etl/util/common.py`（数据库连接、工号/供应商/核算主体/科目映射、币种转换、归一化、必输字段识别、未匹配清单、模板写入）。

## 代码规范

- 代码注释与 docstring 用中文，方便业务核对口径。
- 文件名、目录名、函数名、变量名、常量名必须用英文，不用中文也不用拼音。
- 业务字段名、Excel 表头、sheet 名、展示用文件名可保留来源系统/模板里的中文。
- `resources/source/` 下的源数据文件保持来源系统原名，不重命名，便于溯源对账。
