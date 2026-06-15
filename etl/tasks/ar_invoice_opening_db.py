# -*- coding: utf-8 -*-
"""应收期初 —— 应收报账单(DB 直连版)。

处理流程:
1. 校验泛微字段字典,避免 SQL 字段名/含义写错。
2. 用固定 SQL 从泛微开票表 uf_xtyykp 取数,并左关联收款登记 uf_skdj 的汇总金额。
3. 只对必须跨表/跨系统的 ID 做批量解析,例如申请人、部门、公司主体、客户、合同、币种。
4. 按导入模版字段逐列生成输出,字段旁标注取值来源。

跑法:在项目根执行  python run.py ar_invoice_opening_db
"""
import sys
from pathlib import Path

import pandas as pd

if __package__ is None or __package__ == '':
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from etl import common as c

# ============================ 文件 / 模板 ============================
TASK_NAME = 'ar_invoice_opening_db'
TEMPLATE_DIR = c.TPL_DIR / 'ar_invoice_opening'
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
DATE_SUFFIX = c.today_suffix()

TEMPLATE_FILE = TEMPLATE_DIR / '应收报账单期初数据导入模板.xlsx'
OUTPUT_FILE = OUTPUT_DIR / f'英雄应收报账单期初数据导入_应收期初_{DATE_SUFFIX}.xlsx'
EXCEPTION_FILE = OUTPUT_DIR / f'未匹配清单_应收期初_{DATE_SUFFIX}.xlsx'

TEMPLATE_SHEET = '应收报账单期初数据导入'
RULE_SHEET = '应收期初'
RULE_TABLE = '应收报账单期初数据导入'
DOCUMENT_TYPE = 'OPEN10'
MANAGEMENT_COMPANY = 'Hero'
PAYER_TYPE = '客户'
INCOME_ITEM = '项目收款'
BUSINESS_TYPE_LOV = 'HERO.BUSINESS_TYPE'
INVOICE_TYPE_LOV = 'HERO.INVOICE_TYPE'
CONTRACT_INVOICE_MEANING = '合同'
OUTPUT_COLUMNS = [
    '来源单据号',
    '应收报账单类型',
    '核算主体',
    '管理公司',
    '部门',
    '岗位',
    '申请人',
    '申请日期',
    '支付币种',
    '付款对象类型',
    '付款对象',
    '合同编号',
    '里程碑阶段',
    '平台',
    '业务类型编码',
    '开票类型编码',
    '核销金额',
    '头备注',
    '自审批',
    '自审核',
    '凭证推送',
    '凭证日期',
    '行号',
    '收入分类',
    '收入项目',
    '数量',
    '单价',
    '金额',
    '税率类型',
    '税额',
    '行备注',
    *[f'头维度{i}' for i in range(1, 21)],
    '项目',
    '订单',
    *[f'行维度{i}' for i in range(3, 21)],
]

# 问题清单里,目标字段缺失时带出的泛微源字段。
ISSUE_SOURCE_FIELD_MAP = {
    '来源单据号': '流程编号',
    '核算主体': '公司主体',
    '申请人': '申请人ID',
    '支付币种': '开票币种',
    '付款对象': '客户',
    '合同编号': '开票合同ID',
    '业务类型编码': '业务类型',
    '核销金额': '收款登记已收款金额',
    '金额': '开票金额（含税价）',
    '税率类型': '税率',
}

FW_INVOICE_TABLE = 'uf_xtyykp'
FW_RECEIPT_TABLE = 'uf_skdj'


# ============================ 枚举 / 过滤口径 ============================
# 泛微 uf_xtyykp.kpzt: 开票状态。来源于 workflow_selectitem。
INVOICE_STATUS_MEANINGS = {
    0: '已开票',
    1: '已部分开票',
    2: '已红冲/废票',
    3: '开票失败',
}
VOID_FLAG_MEANINGS = {
    0: '是',
    1: '否',
}
# 泛微 uf_xtyykp.ywlx: 业务类型。0/1 来自 workflow_selectitem;2 在历史数据里表示空业务类型。
BUSINESS_TYPE_CODE_MEANINGS = {
    0: '外部公司',
    1: '外部个人',
    2: '',
}
BUSINESS_TYPE_MEANING = {
    '外部公司': '对公开票',
    '外部个人': '个人开票',
    '': '虚拟开票',
}

TAX_PREFERRED_DESCRIPTIONS = {
    0.00: ['0%销项税，中国', '0%税率', '0%'],
    0.01: ['1%税率(价外)', '1%'],
    0.03: ['3%税率(价外)', '3%'],
    0.06: ['6%销项税，中国', '6%税率', '6%'],
    0.09: ['9%税率(价内)', '9%销项税，中国', '9%'],
    0.13: ['13%税率(价外)', '13%销项税，中国', '13%'],
}

DATE_FROM = '2026-01-01'
ISSUED_STATUS_CODE = 0
VOID_CODE = 0


# ============================ 泛微源 SQL ============================
# 只查开票表,并用子查询按「开票/预收单号」汇总收款登记金额。其他字典/维表解析放到后续批量查询中做。
# 字段含义由 EXPECTED_*_FIELDS + common.validate_fw_fields 在运行时校验。
SOURCE_SQL = """
SELECT
    m.id AS `ID`,
    m.lcbh AS `流程编号`,
    m.sqr AS `申请人ID`,
    m.sqrbm AS `申请人部门ID`,
    m.sqrq AS `申请日期`,
    m.kpht AS `开票合同ID`,
    m.gszt AS `公司主体ID`,
    m.kh AS `客户ID`,
    m.kpjehsj AS `开票金额（含税价）`,
    m.sl AS `税率`,
    m.se AS `税额`,
    m.kpbz AS `开票币种ID`,
    m.kptxt AS `开票备注`,
    m.ywlx AS `业务类型ID`,
    m.xmje AS `不含税金额（明细）`,
    m.xmjshj AS `价税合计（明细）`,
    m.semx AS `税额（明细）`,
    m.slmx AS `税率（明细）`,
    COALESCE(r.receipt_amount, 0) AS `收款登记已收款金额`
FROM uf_xtyykp m
LEFT JOIN (
    SELECT
        kpysdh,
        SUM(COALESCE(bfqrjehj, 0)) AS receipt_amount
    FROM uf_skdj
    WHERE kpysdh IS NOT NULL AND TRIM(kpysdh) <> ''
    GROUP BY kpysdh
) r ON r.kpysdh = m.lcbh
WHERE m.sqrq >= %(date_from)s
  AND m.kpzt = %(issued_status_code)s
  AND (m.sfzf IS NULL OR m.sfzf <> %(void_code)s)
ORDER BY m.id
"""


# 仅用于打印过滤前后数量,便于核对口径。
STATS_SQL = """
SELECT
    SUM(CASE
        WHEN m.sqrq >= %(date_from)s
         AND m.kpzt = %(issued_status_code)s
        THEN 1 ELSE 0 END) AS matched_count,
    SUM(CASE
        WHEN m.sqrq >= %(date_from)s
         AND m.kpzt = %(issued_status_code)s
         AND m.sfzf = %(void_code)s
        THEN 1 ELSE 0 END) AS void_count,
    SUM(CASE
        WHEN m.sqrq >= %(date_from)s
         AND m.kpzt = %(issued_status_code)s
         AND (m.sfzf IS NULL OR m.sfzf <> %(void_code)s)
        THEN 1 ELSE 0 END) AS kept_count
FROM uf_xtyykp m
"""


# 运行前校验字段真实含义。主表字段 detail_table 用空字符串。
EXPECTED_INVOICE_FIELDS = {
    '': {
        'lcbh': '流程编号',
        'sqr': '申请人',
        'sqrbm': '申请人部门',
        'sqrq': '申请日期',
        'kpht': '开票合同',
        'gszt': '公司主体',
        'kh': '客户',
        'kpjehsj': '开票金额（含税价）',
        'sl': '税率',
        'se': '税额',
        'kpzt': '开票状态',
        'kpbz': '开票币种',
        'sfzf': '是否作废',
        'kptxt': '开票备注',
        'ywlx': '业务类型',
        'xmje': '不含税金额（明细）',
        'xmjshj': '价税合计（明细）',
        'semx': '税额（明细）',
        'slmx': '税率（明细）',
    },
}
EXPECTED_RECEIPT_FIELDS = {
    '': {
        'kpysdh': '开票/预收单号',
        'bfqrjehj': '已收款金额',
    },
}


# ============================ DB 查询小工具 ============================
def _query_fw(sql):
    """查询泛微库,统一带上本任务过滤参数。"""
    return c.query_db('FW', 'vspn_xtyy', sql, {
        'date_from': DATE_FROM,
        'issued_status_code': ISSUED_STATUS_CODE,
        'void_code': VOID_CODE,
    })


# ============================ 源值解析 ============================
def _text(value):
    if pd.isna(value):
        return ''
    text = str(value).strip()
    return '' if text in ('', 'nan', 'None', 'NaT') else text


def _first_browser_id(value):
    ids = c.parse_browser_ids(value)
    return ids[0] if ids else ''


def resolve_business_type_name(value):
    """uf_xtyykp.ywlx 业务类型:0=外部公司,1=外部个人,2=空业务类型。"""
    code = c.format_code(value)
    return BUSINESS_TYPE_CODE_MEANINGS.get(int(code), '') if code.isdigit() else ''


def resolve_source_values(source_df):
    """基于主 SQL 返回的泛微 ID 字段补充输出需要的展示值。

    保留原始 ID 列,新增展示列:
    - 申请人ID -> 申请人 / 申请人工号
    - 申请人部门ID -> 申请人部门
    - 公司主体ID -> 公司主体
    - 客户ID -> 客户
    - 开票合同ID -> 开票合同
    - 开票币种ID -> 开票币种
    - 业务类型ID -> 业务类型
    """
    df = source_df.copy()
    employee_map = c.build_fw_employee_info_map_for_ids(df['申请人ID'])
    department_map = c.build_fw_department_name_map_for_ids(df['申请人部门ID'])
    company_map = c.build_fw_company_name_map_for_ids(df['公司主体ID'])
    customer_map = c.build_fw_customer_name_map_for_ids(df['客户ID'])
    contract_map = c.build_fw_contract_code_map_for_ids(df['开票合同ID'])
    currency_map = c.build_fw_currency_name_map_for_ids(df['开票币种ID'])

    # [开票表] sqr -> hrmresource / hrmjobtitles
    df['申请人'] = df['申请人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('name', ''))
    df['申请人工号'] = df['申请人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('code', ''))
    # [开票表] sqrbm -> hrmdepartment.DEPARTMENTNAME
    df['申请人部门'] = df['申请人部门ID'].map(lambda value: department_map.get(c.format_code(value), ''))
    # [开票表] gszt -> uf_gstt.gsmc
    df['公司主体'] = df['公司主体ID'].map(lambda value: company_map.get(c.format_code(value), ''))
    # [开票表] kh -> uf_khgys.khmc
    df['客户'] = df['客户ID'].map(lambda value: customer_map.get(_first_browser_id(value), ''))
    # [开票表] kpht -> uf_htsp.htbh
    df['开票合同'] = df['开票合同ID'].map(lambda value: contract_map.get(_first_browser_id(value), ''))
    # [开票表] kpbz -> fnacurrency.CURRENCYNAME
    df['开票币种'] = df['开票币种ID'].map(lambda value: currency_map.get(c.format_code(value), ''))
    # [开票表] ywlx:0=外部公司,1=外部个人,2=空业务类型
    df['业务类型'] = df['业务类型ID'].map(resolve_business_type_name)
    return df


# ============================ 模板输出 ============================
def _lookup_by_name(mapping, value):
    return '' if pd.isna(value) else mapping.get(c.normalize_name(value), '')


def _business_type_code(value, business_type_map):
    meaning = BUSINESS_TYPE_MEANING.get(_text(value), BUSINESS_TYPE_MEANING[''])
    return business_type_map.get(meaning, '')


def _normalize_tax_rate(value):
    if pd.isna(value):
        return None
    text = str(value).strip().replace('%', '')
    if not text or text in ('nan', 'None'):
        return None
    try:
        rate = float(text)
    except ValueError:
        return None
    if rate > 1:
        rate = rate / 100
    return round(rate, 4)


def _tax_description(value, tax_description_map):
    rate = _normalize_tax_rate(value)
    if rate is None:
        return ''
    return tax_description_map.get(rate, '')


def build_output(invoice_df):
    """DB 源数据 -> 应收报账单期初导入模板 71 列。
    能从泛微 DB 原始字段/ID 直接取得的字段直接取;跨系统编码再查中台/值集。
    """
    # 核算主体: [开票表] gszt -> 公司主体名称 -> Hand hfac_accounting_entity.acc_entity_code。
    entity_map = c.build_accounting_entity_map_for_names(invoice_df['公司主体'])
    # 付款对象: [开票表] kh -> 泛微客户名称 -> Hand hfbs_system_customer.customer_code。
    customer_map = c.build_customer_map_for_names(invoice_df['客户'])
    # 业务类型/开票类型: Hand 值集。
    business_type_map = c.build_lov_meaning_map(BUSINESS_TYPE_LOV)
    invoice_type_map = c.build_lov_meaning_map(INVOICE_TYPE_LOV)
    contract_invoice_code = invoice_type_map.get(CONTRACT_INVOICE_MEANING, '')
    # 税率类型: 税率 -> Hand 税率类型描述。
    tax_description_map = c.build_tax_type_description_map(TAX_PREFERRED_DESCRIPTIONS)

    tax_rate_source = invoice_df['税率'].where(
        invoice_df['税率'].notna(), invoice_df['税率（明细）'])
    tax_amount = invoice_df['税额（明细）'].where(
        invoice_df['税额（明细）'].notna(), invoice_df['税额'])

    output_df = pd.DataFrame(index=invoice_df.index)

    # 固定值。
    output_df['应收报账单类型'] = DOCUMENT_TYPE
    output_df['管理公司'] = MANAGEMENT_COMPANY
    output_df['付款对象类型'] = PAYER_TYPE
    output_df['开票类型编码'] = contract_invoice_code
    output_df['收入项目'] = INCOME_ITEM

    # 泛微开票表直取/解析字段。
    output_df['来源单据号'] = invoice_df['流程编号']                  # [开票表] lcbh
    output_df['部门'] = invoice_df['申请人部门']                      # [开票表] sqrbm -> hrmdepartment
    output_df['岗位'] = ''                                           # 不涉及
    output_df['申请人'] = invoice_df['申请人工号']                    # [开票表] sqr -> hrmjobtitles.JOBTITLENAME
    output_df['申请日期'] = invoice_df['申请日期'].map(c.format_date)  # [开票表] sqrq
    output_df['支付币种'] = invoice_df['开票币种'].map(c.to_iso_currency)  # [开票表] kpbz -> ISO
    output_df['合同编号'] = invoice_df['开票合同'].where(
        invoice_df['开票合同'].notna(), '')                          # [开票表] kpht -> uf_htsp.htbh
    output_df['头备注'] = invoice_df['开票备注'].astype(str).where(
        invoice_df['开票备注'].notna(), '').str.slice(0, 150)         # [开票表] kptxt,截前150字符

    # 当前口径没有直接可用的维度字段,按模板留空。
    output_df['里程碑阶段'] = ''
    output_df['平台'] = ''
    output_df['自审批'] = ''
    output_df['自审核'] = ''
    output_df['凭证推送'] = ''
    output_df['凭证日期'] = ''
    output_df['行号'] = ''
    output_df['收入分类'] = ''
    output_df['数量'] = ''
    output_df['单价'] = ''
    output_df['行备注'] = ''

    # 跨系统映射字段。
    output_df['核算主体'] = invoice_df['公司主体'].map(
        lambda value: _lookup_by_name(entity_map, value))             # [开票表] gszt -> Hand 核算主体编码
    output_df['付款对象'] = invoice_df['客户'].map(
        lambda value: _lookup_by_name(customer_map, value))           # [开票表] kh -> Hand 客户编码
    output_df['业务类型编码'] = invoice_df['业务类型'].map(
        lambda value: _business_type_code(value, business_type_map))  # [开票表] ywlx -> HERO.BUSINESS_TYPE

    # 金额和税率。
    output_df['核销金额'] = pd.to_numeric(
        invoice_df['收款登记已收款金额'], errors='coerce').fillna(0).map(c.round_amount)  # [收款登记] bfqrjehj按单号汇总
    output_df['金额'] = pd.to_numeric(
        invoice_df['开票金额（含税价）'], errors='coerce').map(c.round_amount)  # [开票表] kpjehsj
    output_df['税率类型'] = tax_rate_source.map(
        lambda value: _tax_description(value, tax_description_map))   # [开票表] sl/slmx -> Hand税率类型描述
    output_df['税额'] = pd.to_numeric(tax_amount, errors='coerce').map(c.round_amount)  # [开票表] semx/se

    for column in [f'头维度{i}' for i in range(1, 21)]:
        output_df[column] = ''
    output_df['项目'] = ''
    output_df['订单'] = ''
    for column in [f'行维度{i}' for i in range(3, 21)]:
        output_df[column] = ''

    # write_to_template 按 DataFrame 顺序写入模板,这里显式固定列序。
    return output_df[OUTPUT_COLUMNS]


def read_invoice_source():
    """从 DB 直接读取过滤后的开票记录,并补充输出构建需要的展示值。"""
    c.validate_fw_fields(FW_INVOICE_TABLE, EXPECTED_INVOICE_FIELDS)
    c.validate_fw_fields(FW_RECEIPT_TABLE, EXPECTED_RECEIPT_FIELDS)
    stats = _query_fw(STATS_SQL).iloc[0]
    status_name = INVOICE_STATUS_MEANINGS.get(ISSUED_STATUS_CODE, '')
    void_name = VOID_FLAG_MEANINGS.get(VOID_CODE, '')
    print(f"[应收期初-应收报账单-DB] SQL过滤: 申请日期>={DATE_FROM} "
          f"且 开票状态={ISSUED_STATUS_CODE}({status_name}) 且 是否作废≠{VOID_CODE}({void_name})")
    print(f"  满足前两项 {int(stats['matched_count'] or 0)} 行; "
          f"其中剔除作废 {int(stats['void_count'] or 0)} 行; "
          f"最终保留开票记录 {int(stats['kept_count'] or 0)} 行")

    invoice_df = resolve_source_values(_query_fw(SOURCE_SQL))
    print('[应收期初-应收报账单-DB] SQL开票记录行数:', len(invoice_df))
    return invoice_df


def run():
    # 1. SQL 直接查过滤后的开票记录 + 收款登记汇总金额
    invoice_df = read_invoice_source()

    # 2. 构建输出
    output_df = build_output(invoice_df)
    print('[应收期初-应收报账单-DB] 输出明细行数:', len(output_df))

    # 3. 填充率(必输字段以规则表「是否必填」=Y 为准)
    required_cols = c.required_columns(RULE_SHEET, RULE_TABLE)
    c.report_fill(output_df, required_cols)

    # 4. 写模版
    c.write_to_template(output_df, TEMPLATE_FILE, OUTPUT_FILE, TEMPLATE_SHEET)
    print('已写出:', OUTPUT_FILE)

    # 5. 问题清单
    sheets = {'必输字段未达100%': c.fill_summary(output_df, required_cols, RULE_SHEET, RULE_TABLE)}
    sheets.update(c.collect_field_issues(
        output_df, invoice_df, required_cols, ISSUE_SOURCE_FIELD_MAP, doc_col='来源单据号'))
    c.write_exceptions(EXCEPTION_FILE, sheets)
    print('已写出:', EXCEPTION_FILE, '| 各清单条数:', {
        sheet_name: len(sheet_df) for sheet_name, sheet_df in sheets.items()
    })


if __name__ == '__main__':
    run()
