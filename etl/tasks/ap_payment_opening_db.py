# -*- coding: utf-8 -*-
"""应付期初 —— 对公付款单(DB 直连版)。

处理流程:
1. 校验泛微字段字典,避免 SQL 字段名/含义写错。
2. 用固定 SQL 从泛微主表 uf_dgfktz 和明细表 uf_dgfktz_dt1 取数。
3. 只对必须跨表/跨系统的 ID 做批量解析,例如经办人、公司主体、币种、预算科目、供应商编码。
4. 按导入模版字段逐列生成输出,字段旁标注取值来源。

跑法:在项目根执行  python run.py ap_payment_opening_db
"""
import sys
from pathlib import Path

import pandas as pd

if __package__ is None or __package__ == '':
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from etl import common as c

# ============================ 文件 / 模板 ============================
TASK_NAME = 'ap_payment_opening_db'
TEMPLATE_DIR = c.TPL_DIR / 'ap_payment_opening'
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
DATE_SUFFIX = c.today_suffix()

TEMPLATE_FILE = TEMPLATE_DIR / '英雄期初对公付款单导入模版.xlsx'
OUTPUT_FILE = OUTPUT_DIR / f'英雄期初对公付款单导入_应付期初_{DATE_SUFFIX}.xlsx'
EXCEPTION_FILE = OUTPUT_DIR / f'未匹配清单_应付期初_{DATE_SUFFIX}.xlsx'
SUPPLIER_VENDOR_MISSING_FILE = OUTPUT_DIR / f'Hand按ID查不到的供应商_应付期初_{DATE_SUFFIX}.xlsx'

TEMPLATE_SHEET = '期初对公付款单导入'
RULE_SHEET = '应付期初'
RULE_TABLE = '期初对公付款单'
DOCUMENT_TYPE = 'AP01-1'
OUTPUT_COLUMNS = [
    '来源系统',
    '来源单据编号',
    '申请日期',
    '单据类型',
    '申请人工号',
    '申请人姓名',
    '订单编号',
    '订单名称',
    '核算主体编号',
    '核算主体描述',
    '备注',
    '合同号',
    '合同收支计划行',
    '收款方编码',
    '收款方描述',
    '银行账号',
    '计划付款日期',
    '银行转账备注',
    '实际已支付金额',
    '费用项目编码',
    '费用项目描述',
    '主播房间号',
    '报账币种',
    '报账金额（支付币种）',
    '泛微费用项目编码',
]

# 问题清单里,目标字段缺失时带出的泛微源字段。
ISSUE_SOURCE_FIELDS = {
    '申请人工号': '经办人',
    '收款方编码': '供应商ID',
    '核算主体编号': '公司主体',
    '费用项目编码': '预算科目',
    '报账币种': '付款币种',
    '订单编号': '项目编号',
}

FW_TABLE = 'uf_dgfktz'
FW_DETAIL_TABLE = 'uf_dgfktz_dt1'


# ============================ 枚举 / 过滤口径 ============================
# 泛微 uf_dgfktz.lcly: 流程来源。来源于 workflow_selectitem。
FLOW_SOURCE_MEANINGS = {
    0: '团建对公付款',
    1: '签约金对公付款',
    2: '对公付款',
    3: '零工平台付款',
    4: '零工平台到票',
    5: '预付款',
    6: '预付款退款',
    7: '预付款转移',
    8: '费用分摊',
    9: 'HR薪资',
    10: '个人劳务付款',
    11: '团建对公预付',
}
FLOW_STATUS_MEANINGS = {
    0: '未提交',
    1: '审批中',
    2: '审批完成',
}
VOID_FLAG_MEANINGS = {
    0: '是',
    1: '否',
}
PAYMENT_STATUS_MEANINGS = {
    0: '已支付',
}

# 当前任务只取对公付款和个人劳务付款;预付款、零工平台付款等来源走其他任务。
SOURCE_CODES = (2, 10)
DATE_FROM = '2026-01-01'
APPROVED_STATUS_CODE = 2
VOID_CODE = 0


# ============================ 泛微源 SQL ============================
# 只关联主表和明细表。字典/维表解析放到后续批量查询中做,避免源 SQL 变复杂。
# 字段含义由 EXPECTED_FW_FIELDS + common.validate_fw_fields 在运行时校验。
SOURCE_SQL = """
SELECT
    m.id AS `ID`,
    m.lcbh AS `流程编号`,
    m.sqrq AS `申请日期`,
    m.jbr AS `经办人ID`,
    m.xmbh AS `项目编号`,
    m.gszt AS `公司主体ID`,
    m.bz AS `备注`,
    m.xght AS `相关合同ID`,
    m.gys AS `供应商ID`,
    m.gyswb AS `供应商-文本`,
    m.yhzh AS `银行账号`,
    m.yjfkrq AS `预计付款日期`,
    m.fkbz AS `付款币种ID`,
    m.zfzt AS `支付状态ID`,
    d.fkje AS `付款金额`,
    d.yskm AS `预算科目ID`
FROM uf_dgfktz m
JOIN uf_dgfktz_dt1 d ON d.mainid = m.id
WHERE m.lcly IN %(source_codes)s
  AND m.sqrq >= %(date_from)s
  AND m.lcz = %(approved_status_code)s
  AND (m.sfzf IS NULL OR m.sfzf <> %(void_code)s)
ORDER BY m.id, d.id
"""


# 仅用于打印过滤前后数量,便于核对口径。
STATS_SQL = """
SELECT
    SUM(CASE
        WHEN m.lcly IN %(source_codes)s
         AND m.sqrq >= %(date_from)s
         AND m.lcz = %(approved_status_code)s
        THEN 1 ELSE 0 END) AS matched_count,
    SUM(CASE
        WHEN m.lcly IN %(source_codes)s
         AND m.sqrq >= %(date_from)s
         AND m.lcz = %(approved_status_code)s
         AND m.sfzf = %(void_code)s
        THEN 1 ELSE 0 END) AS void_count,
    SUM(CASE
        WHEN m.lcly IN %(source_codes)s
         AND m.sqrq >= %(date_from)s
         AND m.lcz = %(approved_status_code)s
         AND (m.sfzf IS NULL OR m.sfzf <> %(void_code)s)
        THEN 1 ELSE 0 END) AS kept_count
FROM uf_dgfktz m
"""


# 运行前校验字段真实含义。主表字段 detail_table 用空字符串;明细字段用明细表名。
EXPECTED_FW_FIELDS = {
    '': {
        'lcbh': '流程编号',
        'sqrq': '申请日期',
        'jbr': '经办人',
        'xmbh': '项目编号',
        'gszt': '公司主体',
        'bz': '备注',
        'xght': '相关合同',
        'gys': '供应商',
        'gyswb': '供应商-文本',
        'yhzh': '银行账号',
        'yjfkrq': '预计付款日期',
        'fkbz': '付款币种',
        'zfzt': '支付状态',
        'lcly': '流程来源',
        'lcz': '流程状态',
        'sfzf': '是否作废',
    },
    FW_DETAIL_TABLE: {
        'fkje': '付款金额',
        'yskm': '预算科目',
    },
}


# ============================ DB 查询小工具 ============================
def _query_fw(sql):
    """查询泛微库,统一带上本任务过滤参数。"""
    return c.query_db('FW', 'vspn_xtyy', sql, {
        'source_codes': SOURCE_CODES,
        'date_from': DATE_FROM,
        'approved_status_code': APPROVED_STATUS_CODE,
        'void_code': VOID_CODE,
    })


# ============================ 源值解析 ============================
def _lookup_first_browser_value(mapping, value):
    for item_id in c.parse_browser_ids(value):
        mapped = mapping.get(item_id, '')
        if mapped:
            return mapped
    return ''


def resolve_payment_status_name(value):
    """uf_dgfktz.zfzt 支付状态: 0=已支付。"""
    code = c.format_code(value)
    return PAYMENT_STATUS_MEANINGS.get(int(code), '') if code.isdigit() else ''


def resolve_source_values(source_df):
    """基于主 SQL 返回的泛微 ID 字段补充输出需要的展示值。

    保留原始 ID 列,新增展示列:
    - 经办人ID -> 经办人 / 经办人工号
    - 公司主体ID -> 公司主体
    - 相关合同ID -> 合同编号
    - 付款币种ID -> 付款币种
    - 预算科目ID -> 预算科目完整路径
    """
    df = source_df.copy()
    employee_map = c.build_fw_employee_info_map_for_ids(df['经办人ID'])
    company_map = c.build_fw_company_name_map_for_ids(df['公司主体ID'])
    contract_map = c.build_fw_contract_code_map_for_ids(df['相关合同ID'])
    currency_map = c.build_fw_currency_name_map_for_ids(df['付款币种ID'])
    subject_map = c.build_fw_budget_subject_path_map_for_ids(df['预算科目ID'])
    bank_account_map = c.build_fw_supplier_bank_account_map_for_ids(df['银行账号'])

    # [主表] jbr -> hrmresource / hrmjobtitles
    df['经办人'] = df['经办人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('name', ''))
    df['经办人工号'] = df['经办人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('code', ''))
    # [主表] gszt -> uf_gstt.gsmc
    df['公司主体'] = df['公司主体ID'].map(lambda value: company_map.get(c.format_code(value), ''))
    # [主表] xght -> uf_htsp.htbh
    df['相关合同'] = df['相关合同ID'].map(lambda value: _lookup_first_browser_value(contract_map, value))
    # [主表] yhzh 是供应商银行账号浏览框 ID,先转成账号文本,后续再按 Hand 供应商银行卡校验。
    df['银行账号'] = df['银行账号'].map(
        lambda value: _lookup_first_browser_value(bank_account_map, value)
        or ('' if pd.isna(value) else str(value).strip()))
    # [主表] fkbz -> fnacurrency.CURRENCYNAME
    df['付款币种'] = df['付款币种ID'].map(lambda value: currency_map.get(c.format_code(value), ''))
    # [主表] zfzt: 0=已支付
    df['支付状态'] = df['支付状态ID'].map(resolve_payment_status_name)
    # [明细表] yskm -> fnabudgetfeetype 层级路径
    df['预算科目'] = df['预算科目ID'].map(lambda value: subject_map.get(c.format_code(value), ''))
    return df


# ============================ 模板输出 ============================
def build_output(merged_df):
    """DB 源数据 -> 导入模版 24 列。
    能从泛微 DB 原始字段/ID 直接取得的字段直接取;跨系统编码再查中台/规则表。
    """
    def lookup_by_name(mapping, value):
        return '' if pd.isna(value) else mapping.get(c.normalize_name(value), '')

    # 供应商编码: [主表] gys -> common 供应商判断 -> Hand hfbs_system_vender.vender_code。
    # 同时输出 Hand 按 ID 查不到的供应商诊断清单。
    vendor_map = c.build_supplier_vendor_info_map_for_rows(
        merged_df['供应商ID'],
        supplier_texts=merged_df['供应商-文本'],
        document_numbers=merged_df['流程编号'],
        missing_report_file=SUPPLIER_VENDOR_MISSING_FILE,
        log_prefix='[应付期初-供应商付款-DB]',
    )
    # 核算主体编号: [主表] gszt -> 公司主体名称 -> Hand hfac_accounting_entity.acc_entity_code。
    entity_map = c.build_accounting_entity_map_for_names(merged_df['公司主体'])
    # 费用项目编码/描述: [明细] yskm -> 预算科目路径 -> 规则表「业财项目_数据映射规则.xlsx」。
    subject_map = c.build_subject_map()

    def lookup_vendor(index, field):
        return vendor_map.get(index, {}).get(field, '')

    def vendor_description(index, fallback_text):
        name = lookup_vendor(index, 'name')
        if name:
            return name
        return '' if pd.isna(fallback_text) else str(fallback_text).strip()

    def subject_item(subject_path, index):
        if pd.isna(subject_path):
            return ''
        return subject_map.get(c.remove_slashes(subject_path), ('', ''))[index]

    payment_amount = pd.to_numeric(merged_df['付款金额'], errors='coerce')
    paid_mask = merged_df['支付状态ID'].map(c.format_code) == '0'

    output_df = pd.DataFrame(index=merged_df.index)

    # 固定值:当前任务来源就是泛微,单据类型固定为期初对公付款单。
    output_df['来源系统'] = 'FW'
    output_df['单据类型'] = DOCUMENT_TYPE

    # 泛微主表直取字段。
    output_df['来源单据编号'] = merged_df['流程编号']                  # [主表] lcbh
    output_df['申请日期'] = merged_df['申请日期'].map(c.format_date)   # [主表] sqrq
    output_df['申请人工号'] = merged_df['经办人工号']                  # [主表] jbr -> hrmjobtitles.JOBTITLENAME
    output_df['申请人姓名'] = merged_df['经办人']                      # [主表] jbr -> hrmresource.LASTNAME
    output_df['备注'] = merged_df['备注'].astype(str).where(
        merged_df['备注'].notna(), '').str.slice(0, 150)              # [主表] bz,导入限制截前150字符
    output_df['合同号'] = merged_df['相关合同'].where(
        merged_df['相关合同'].notna(), '')                            # [主表] xght -> uf_htsp.htbh
    output_df['银行账号'] = merged_df['银行账号'].where(
        merged_df['银行账号'].notna(), '')                            # [主表] yhzh 原值
    output_df['计划付款日期'] = merged_df['预计付款日期'].map(c.format_date)  # [主表] yjfkrq

    # 当前源数据没有直接可用的订单/计划行/银行备注/主播房间号,按模板留空。
    output_df['订单编号'] = ''                                         # 模板字段,当前口径不取项目编号
    output_df['订单名称'] = ''
    output_df['合同收支计划行'] = ''
    output_df['银行转账备注'] = ''
    output_df['主播房间号'] = ''

    # 跨系统映射字段。
    output_df['核算主体编号'] = merged_df['公司主体'].map(
        lambda value: lookup_by_name(entity_map, value))              # [主表] gszt -> Hand 核算主体编号
    output_df['核算主体描述'] = merged_df['公司主体']                  # [主表] gszt -> uf_gstt.gsmc
    output_df['收款方编码'] = [lookup_vendor(index, 'code') for index in merged_df.index]  # [主表] gys -> Hand vender_code
    output_df['收款方描述'] = [
        vendor_description(index, supplier_text)
        for index, supplier_text in zip(merged_df.index, merged_df['供应商-文本'])
    ]                                                                  # 优先 Hand description,兜底 [主表] gyswb
    output_df['银行账号'] = c.resolve_hand_vendor_bank_accounts(
        output_df['收款方编码'], merged_df['银行账号'])                  # 按收款方 Hand 供应商银行卡校验;为空/不匹配时取默认账号

    # 金额和费用项目。
    output_df['实际已支付金额'] = [
        c.round_amount(value) if is_paid else 0 for value, is_paid in zip(payment_amount, paid_mask)
    ]                                                                  # [明细] fkje;仅 zfzt=0(已支付) 计入
    output_df['费用项目编码'] = merged_df['预算科目'].map(
        lambda value: subject_item(value, 0))                          # [明细] yskm -> 规则表编码
    output_df['费用项目描述'] = merged_df['预算科目'].map(
        lambda value: subject_item(value, 1))                          # [明细] yskm -> 规则表描述
    output_df['报账币种'] = merged_df['付款币种'].map(c.to_iso_currency)  # [主表] fkbz -> ISO币种
    output_df['报账金额（支付币种）'] = payment_amount.map(c.round_amount)  # [明细] fkje
    output_df['泛微费用项目编码'] = merged_df['预算科目'].where(
        merged_df['预算科目'].notna(), '')                             # [明细] yskm -> 原泛微预算科目路径
    # write_to_template 按 DataFrame 顺序写入模板,这里显式固定列序。
    return output_df[OUTPUT_COLUMNS]


def read_merged_source():
    """从 DB 直接读取过滤后的主子合并数据,再补充输出构建需要的展示值。"""
    c.validate_fw_fields(FW_TABLE, EXPECTED_FW_FIELDS)
    stats = _query_fw(STATS_SQL).iloc[0]
    source_names = ', '.join(f'{code}={FLOW_SOURCE_MEANINGS.get(code, "")}' for code in SOURCE_CODES)
    status_name = FLOW_STATUS_MEANINGS.get(APPROVED_STATUS_CODE, '')
    void_name = VOID_FLAG_MEANINGS.get(VOID_CODE, '')
    print(f"[应付期初-供应商付款-DB] SQL过滤: 流程来源∈{SOURCE_CODES}({source_names}) 且 申请日期>={DATE_FROM} "
          f"且 流程状态={APPROVED_STATUS_CODE}({status_name}) 且 是否作废≠{VOID_CODE}({void_name})")
    print(f"  满足前三项 {int(stats['matched_count'] or 0)} 单; "
          f"其中剔除作废 {int(stats['void_count'] or 0)} 单; "
          f"最终保留主表 {int(stats['kept_count'] or 0)} 单")

    merged_df = resolve_source_values(_query_fw(SOURCE_SQL))
    print('[应付期初-供应商付款-DB] SQL主子合并明细行数:', len(merged_df))
    return merged_df


def run():
    # 1. SQL 直接查过滤后的主子合并源数据
    merged_df = read_merged_source()

    # 2. 构建输出
    output_df = build_output(merged_df)
    print('[应付期初-供应商付款-DB] 输出明细行数:', len(output_df))

    # 3. 填充率(必输字段以规则表「是否必填」=Y 为准)
    required_cols = c.required_columns(RULE_SHEET, RULE_TABLE)
    c.report_fill(output_df, required_cols)

    # 4. 写模版
    c.write_to_template(output_df, TEMPLATE_FILE, OUTPUT_FILE, TEMPLATE_SHEET)
    print('已写出:', OUTPUT_FILE)

    # 5. 问题清单
    sheets = {'必输字段未达100%': c.fill_summary(output_df, required_cols, RULE_SHEET, RULE_TABLE)}
    sheets.update(c.collect_field_issues(output_df, merged_df, required_cols, ISSUE_SOURCE_FIELDS))
    bank_issues = c.collect_hand_vendor_bank_account_issues(output_df, merged_df['银行账号'])
    if not bank_issues.empty:
        sheets['银行账号_校验异常'] = bank_issues
    c.write_exceptions(EXCEPTION_FILE, sheets)
    print('已写出:', EXCEPTION_FILE, '| 各清单条数:', {
        sheet_name: len(sheet_df) for sheet_name, sheet_df in sheets.items()
    })


if __name__ == '__main__':
    run()
