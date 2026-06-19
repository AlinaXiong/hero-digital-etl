# -*- coding: utf-8 -*-
"""应付期初 —— 对公付款单 / 批量费用流程 / 只转入外部成本(DB 直连版)。

按《业财项目_数据映射规则.xlsx》-「应付期初」生成期初对公付款单导入模板的三个 sheet:
    1. 期初对公付款单导入
    2. 批量费用流程
    3. 只转入外部成本

跑法:在项目根执行  python run.py ap_payment_opening_extra_db
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

if __package__ is None or __package__ == '':
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from etl import common as c
from etl.tasks import ap_payment_opening_db as base_payment
from etl.tasks.ap_prepayment_opening_db import (
    _order_mapping_value,
    collect_order_mapping_issues,
)


# ============================ 文件 / 模板 ============================
TASK_NAME = 'ap_payment_opening_extra_db'
TEMPLATE_DIR = c.TPL_DIR / 'ap_payment_opening'
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
DATE_SUFFIX = c.today_suffix()

TEMPLATE_FILE = TEMPLATE_DIR / '英雄期初对公付款单导入模版.xlsx'
OUTPUT_FILE = OUTPUT_DIR / f'英雄期初对公付款单导入_应付期初_补充_{DATE_SUFFIX}.xlsx'
EXCEPTION_FILE = OUTPUT_DIR / f'未匹配清单_应付期初_补充_{DATE_SUFFIX}.xlsx'
BASE_SUPPLIER_VENDOR_MISSING_FILE = OUTPUT_DIR / f'Hand按ID查不到的供应商_期初对公付款单导入_{DATE_SUFFIX}.xlsx'
BATCH_SUPPLIER_VENDOR_MISSING_FILE = OUTPUT_DIR / f'Hand按ID查不到的供应商_批量费用流程_{DATE_SUFFIX}.xlsx'

BASE_TEMPLATE_SHEET = '期初对公付款单导入'
SHEET_BATCH = '批量费用流程'
SHEET_EXTERNAL_COST = '只转入外部成本'
RULE_SHEET = '应付期初'
RULE_TABLE = '期初对公付款单'

DOCUMENT_TYPE = 'AP01-1'
DATE_FROM = '2026-01-01'
SOURCE_SYSTEM = 'FW'
VIRTUAL_VENDOR_NAME = '外部成本转移虚拟供应商'

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
    '泛微项目编号',
    '泛微费用项目编码',
]

BATCH_ISSUE_SOURCE_FIELDS = {
    '申请人工号': '申请人',
    '收款方编码': '供应商',
    '核算主体编号': '公司主体',
    '费用项目编码': '预算科目',
    '订单编号': '项目编号',
}
EXTERNAL_COST_ISSUE_SOURCE_FIELDS = {
    '申请人工号': '申请人',
    '收款方编码': '收款方描述',
    '核算主体编号': '公司主体',
    '费用项目编码': '预算科目',
    '订单编号': '项目编号',
}


# ============================ 泛微源 SQL ============================
BATCH_SOURCE_SQL = """
SELECT
    m.id AS `ID`,
    d.id AS `明细ID`,
    COALESCE(NULLIF(m.fybh, ''), LEFT(d.fymxbh, CHAR_LENGTH(d.fymxbh) - 4), d.fymxbh) AS `流程编号`,
    d.jlrq AS `申请日期`,
    d.jsr AS `申请人ID`,
    d.xmbh AS `项目编号ID`,
    m.gszt AS `公司主体ID`,
    COALESCE(NULLIF(d.bz, ''), NULLIF(m.pcbz, ''), NULLIF(m.fypcbz, ''), NULLIF(m.fypcmc, '')) AS `备注`,
    m.dwfkdw AS `供应商ID`,
    d.je AS `金额`,
    COALESCE(d.yskm, d.fyxh) AS `预算科目ID`
FROM uf_plfy m
JOIN uf_plfy_dt1 d ON d.mainid = m.id
WHERE d.sfqr = 0
  AND (d.sfzf IS NULL OR d.sfzf <> 0)
  AND d.jlrq >= %(date_from)s
ORDER BY d.jlrq, m.id, d.id
"""

BATCH_STATS_SQL = """
SELECT
    COUNT(*) AS kept_count,
    COUNT(DISTINCT m.id) AS document_count,
    SUM(d.je) AS amount_total
FROM uf_plfy m
JOIN uf_plfy_dt1 d ON d.mainid = m.id
WHERE d.sfqr = 0
  AND (d.sfzf IS NULL OR d.sfzf <> 0)
  AND d.jlrq >= %(date_from)s
"""

EXTERNAL_COST_SOURCE_SQL = """
SELECT
    m.id AS `ID`,
    d.id AS `明细ID`,
    m.lcbh AS `流程编号`,
    m.sqrq AS `申请日期`,
    m.sqr AS `申请人ID`,
    m.zrxmbh AS `转入项目编号ID`,
    m.zcxmbh AS `转出项目编号ID`,
    m.zrxmbhmcnss AS `转入MCN赛事项目编号ID`,
    m.zcxmbhmcnss AS `转出MCN赛事项目编号ID`,
    m.jsxmmc AS `转入项目名称`,
    m.zcxmmc AS `转出项目名称`,
    m.zrgszt AS `转入公司主体ID`,
    m.zcgszt AS `转出公司主体ID`,
    m.zrzcjehz AS `转入总金额`,
    m.zczcjehz AS `转出总金额`,
    m.fyd AS `费用单ID`,
    d.stzjid AS `费用明细视图ID`,
    d.mxid AS `费用明细ID`,
    d.zczcje AS `转出明细金额`,
    d.srje AS `转出收入金额`,
    v.lcbh AS `费用单据号`,
    COALESCE(v.yskm, d.yslx) AS `预算科目ID`,
    v.fkbz AS `付款币种ID`,
    v.fkje AS `费用原金额`,
    v.sywzje AS `费用剩余未占金额`,
    v.yzje AS `费用已占金额`,
    CASE m.ly
        WHEN 5 THEN '赛事只转入外部成本'
        WHEN 2 THEN 'MCN只转入外部成本'
        ELSE CONCAT('内部收支来源', m.ly)
    END AS `内部收支来源`
FROM uf_xtyynbsz m
JOIN uf_xtyynbsz_dt10 d ON d.mainid = m.id
LEFT JOIN view_costlist_ys v
    ON v.id = COALESCE(NULLIF(d.stzjid, ''), NULLIF(d.mxid, ''))
WHERE m.ly IN (5, 2)
  AND m.lczt IN (1, 2)
  AND m.sfzf IS NULL
  AND m.sqrq >= %(date_from)s
ORDER BY m.sqrq, m.id, d.id
"""

EXTERNAL_COST_STATS_SQL = """
SELECT
    COUNT(*) AS document_count,
    SUM(CASE WHEN ly = 5 THEN 1 ELSE 0 END) AS event_document_count,
    SUM(CASE WHEN ly = 2 THEN 1 ELSE 0 END) AS mcn_document_count,
    SUM(COALESCE(zrzcjehz, 0)) AS in_amount_total,
    SUM(COALESCE(zczcjehz, 0)) AS out_amount_total
FROM uf_xtyynbsz
WHERE ly IN (5, 2)
  AND lczt IN (1, 2)
  AND sfzf IS NULL
  AND sqrq >= %(date_from)s
"""


# ============================ 小工具 ============================
def _query_fw(sql):
    return c.query_db('FW', 'vspn_xtyy', sql, {'date_from': DATE_FROM})


def _text(value):
    if pd.isna(value):
        return ''
    text = str(value).strip()
    return '' if text in ('', 'nan', 'None', 'NaT') else text


def _first_browser_value(mapping, value):
    for item_id in c.parse_browser_ids(value):
        mapped = mapping.get(item_id, '')
        if mapped:
            return mapped
    return ''


def _lookup_by_name(mapping, value):
    text = _text(value)
    return mapping.get(c.normalize_name(text), '') if text else ''


def _first_non_blank(*values):
    for value in values:
        text = _text(value)
        if text:
            return text
    return ''


PROJECT_TABLES = ('uf_xtyyxmkp', 'uf_xmkp', 'view_xmjkzb')


def build_fw_project_info_map_for_ids(project_values):
    """泛微项目浏览框 ID -> 项目编号/名称,兼容 xmkp/xmjk/MCN赛事项目等来源。"""
    project_ids = c.clean_codes(
        project_id
        for value in project_values
        for project_id in c.parse_browser_ids(value)
    )
    if not project_ids:
        return {}

    result = {}
    remaining = set(project_ids)
    for table in PROJECT_TABLES:
        if not remaining:
            break
        ids = [project_id for project_id in project_ids if project_id in remaining]
        try:
            project_df = c.query_db(
                'FW',
                'vspn_xtyy',
                f'SELECT id, xmbh AS project_code, xmmc AS project_name FROM {table} '
                f'WHERE id IN ({c.in_placeholders(ids)})',
                ids,
            )
        except Exception:
            try:
                project_df = c.query_db(
                    'FW',
                    'vspn_xtyy',
                    f"SELECT id, xmbh AS project_code, '' AS project_name FROM {table} "
                    f'WHERE id IN ({c.in_placeholders(ids)})',
                    ids,
                )
            except Exception:
                continue
        for _, row in project_df.iterrows():
            project_id = c.format_code(row['id'])
            project_code = _text(row['project_code'])
            if project_id and project_code and project_id not in result:
                result[project_id] = {
                    'code': project_code,
                    'name': _text(row.get('project_name', '')),
                }
                remaining.discard(project_id)
    return result


def build_fw_project_code_map_for_ids(project_values):
    return {
        project_id: info.get('code', '')
        for project_id, info in build_fw_project_info_map_for_ids(project_values).items()
    }


def _resolve_project_codes(values, project_map):
    return values.map(lambda value: _first_browser_value(project_map, value) or _text(value))


def _resolve_project_names(values, project_name_map):
    return values.map(lambda value: _first_browser_value(project_name_map, value))


def _with_resolved_project_fields(source_df, project_column='项目编号'):
    df = source_df.copy()
    if project_column not in df.columns:
        df['项目编号'] = ''
        df['项目名称'] = df['项目名称'] if '项目名称' in df.columns else ''
        return df

    project_info_map = build_fw_project_info_map_for_ids(df[project_column])
    project_code_map = {
        project_id: info.get('code', '')
        for project_id, info in project_info_map.items()
    }
    project_name_map = {
        project_id: info.get('name', '')
        for project_id, info in project_info_map.items()
    }

    df['项目编号'] = _resolve_project_codes(df[project_column], project_code_map)
    mapped_project_names = _resolve_project_names(df[project_column], project_name_map)
    existing_project_names = df['项目名称'] if '项目名称' in df.columns else pd.Series('', index=df.index)
    df['项目名称'] = [
        _first_non_blank(existing_name, mapped_name)
        for existing_name, mapped_name in zip(existing_project_names, mapped_project_names)
    ]
    return df


def _apply_order_project_columns(output_df, source_df):
    """按预付期初口径补充泛微项目编号,并用项目&订单清洗表映射订单字段。"""
    df = output_df.copy()
    project_source_df = _with_resolved_project_fields(source_df)
    project_codes = project_source_df['项目编号'].map(_text)
    df['泛微项目编号'] = project_codes
    df['订单编号'] = project_codes.map(lambda value: _order_mapping_value(value, '订单编号'))
    df['订单名称'] = project_codes.map(lambda value: _order_mapping_value(value, '订单标题'))
    return df[OUTPUT_COLUMNS]


def _enrich_missing_order_issue(sheets, output_df, source_df):
    """让「缺失_订单编号」清单同时带出项目编号/名称,便于回查原单。"""
    sheet_name = '缺失_订单编号'
    if sheet_name not in sheets or '订单编号' not in output_df.columns:
        return sheets

    project_source_df = _with_resolved_project_fields(source_df)
    blank_mask = output_df['订单编号'].astype(str).str.strip() == ''
    data = {
        '来源单据编号': output_df.loc[blank_mask, '来源单据编号'].astype(str),
        '订单编号': output_df.loc[blank_mask, '订单编号'].astype(str),
        '泛微项目编号': project_source_df.loc[blank_mask, '项目编号'].astype(str),
        '泛微项目名称': project_source_df.loc[blank_mask, '项目名称'].astype(str),
    }
    sheets[sheet_name] = pd.DataFrame(data).drop_duplicates().reset_index(drop=True)
    return sheets


def _resolve_subject_paths(subject_ids):
    subject_map = c.build_fw_budget_subject_path_map_for_ids(subject_ids)
    return subject_ids.map(lambda value: subject_map.get(c.format_code(value), ''))


def _subject_item(subject_lookup, value, index):
    subject_path = _text(value)
    return subject_lookup.get(c.remove_slashes(subject_path), ('', ''))[index] if subject_path else ''


def _supplier_name(value, supplier_status_map):
    selected_id = c.choose_fw_supplier_id(c.parse_browser_ids(value), supplier_status_map)
    return supplier_status_map.get(selected_id, {}).get('name', '')


def _format_signed_amount(value, sign=1):
    if pd.isna(value):
        return ''
    return c.round_amount(float(value) * sign)


def _copy_template_sheet(workbook, target_sheet_name):
    if target_sheet_name in workbook.sheetnames:
        del workbook[target_sheet_name]
    source = workbook[BASE_TEMPLATE_SHEET]
    copied = workbook.copy_worksheet(source)
    copied.title = target_sheet_name
    return copied


def _fill_sheet(worksheet, output_df):
    for col_idx, column_name in enumerate(output_df.columns, start=1):
        worksheet.cell(row=1, column=col_idx).value = column_name
    if worksheet.max_row > 1:
        worksheet.delete_rows(2, worksheet.max_row)
    for _, row in output_df.iterrows():
        worksheet.append(['' if pd.isna(value) else value for value in row.tolist()])


def write_output_workbook(base_output_df, batch_output_df, external_output_df):
    wb = load_workbook(TEMPLATE_FILE)
    _fill_sheet(wb[BASE_TEMPLATE_SHEET], base_output_df)
    _fill_sheet(_copy_template_sheet(wb, SHEET_BATCH), batch_output_df)
    _fill_sheet(_copy_template_sheet(wb, SHEET_EXTERNAL_COST), external_output_df)
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_FILE)
    return OUTPUT_FILE


def _allocate_group_amounts(group, total_col):
    total = pd.to_numeric(group[total_col], errors='coerce').dropna()
    direct = pd.to_numeric(group.get('转出明细金额'), errors='coerce')
    if direct.notna().any():
        return direct.fillna(0)

    base = pd.to_numeric(group['费用已占金额'], errors='coerce')
    if base.isna().all() or float(base.abs().sum()) == 0:
        base = pd.to_numeric(group['费用原金额'], errors='coerce')
    if base.isna().all() or float(base.abs().sum()) == 0:
        base = pd.Series([1] * len(group), index=group.index, dtype='float64')
    base = base.fillna(0).abs()
    base_sum = float(base.sum())

    if len(total) > 0 and base_sum:
        return base / base_sum * float(total.iloc[0])
    return base


def allocate_external_cost_amounts(source_df):
    df = source_df.copy()
    if df.empty:
        df['转入金额'] = []
        df['转出金额'] = []
        return df

    in_amounts = []
    out_amounts = []
    for _, group in df.groupby('ID', sort=False):
        allocated_in = _allocate_group_amounts(group, '转入总金额')
        allocated_out = _allocate_group_amounts(group, '转出总金额')
        in_amounts.append(allocated_in)
        out_amounts.append(allocated_out)
    df['转入金额'] = pd.concat(in_amounts).reindex(df.index)
    df['转出金额'] = pd.concat(out_amounts).reindex(df.index)
    return df


# ============================ 批量费用流程 ============================
def read_base_source():
    """复用原应付期初 DB 任务的源读取逻辑。"""
    base_payment.SUPPLIER_VENDOR_MISSING_FILE = BASE_SUPPLIER_VENDOR_MISSING_FILE
    return base_payment.read_merged_source()


def read_batch_source():
    stats = _query_fw(BATCH_STATS_SQL).iloc[0]
    print(f"[应付期初-批量费用流程] SQL过滤: 是否确认=是 且 是否作废≠是 且 记录日期>={DATE_FROM}")
    print(f"  保留批量费用 {int(stats['document_count'] or 0)} 单 / 明细 {int(stats['kept_count'] or 0)} 行; "
          f"金额合计 {float(stats['amount_total'] or 0):.2f}")
    source_df = _query_fw(BATCH_SOURCE_SQL)
    print('[应付期初-批量费用流程] SQL明细行数:', len(source_df))
    return resolve_batch_values(source_df)


def resolve_batch_values(source_df):
    df = source_df.copy()
    employee_map = c.build_fw_employee_info_map_for_ids(df['申请人ID'])
    company_map = c.build_fw_company_name_map_for_ids(df['公司主体ID'])
    supplier_status_map = c.build_fw_supplier_status_map(df['供应商ID'])

    df['申请人'] = df['申请人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('name', ''))
    df['申请人工号'] = df['申请人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('code', ''))
    df['公司主体'] = df['公司主体ID'].map(lambda value: company_map.get(c.format_code(value), ''))
    df['供应商'] = df['供应商ID'].map(lambda value: _supplier_name(value, supplier_status_map))
    df = _with_resolved_project_fields(df, '项目编号ID')
    df['预算科目'] = _resolve_subject_paths(df['预算科目ID'])
    return df


def build_batch_output(source_df):
    vendor_info_map = c.build_supplier_vendor_info_map_for_rows(
        source_df['供应商ID'],
        supplier_texts=source_df['供应商'],
        document_numbers=source_df['流程编号'],
        missing_report_file=BATCH_SUPPLIER_VENDOR_MISSING_FILE,
        log_prefix='[应付期初-批量费用流程]',
    )
    entity_map = c.build_accounting_entity_map_for_names(source_df['公司主体'])
    subject_lookup = c.build_subject_map()
    amount = pd.to_numeric(source_df['金额'], errors='coerce')

    def vendor_field(index, field):
        return vendor_info_map.get(index, {}).get(field, '')

    output_df = pd.DataFrame(index=source_df.index)
    output_df['来源系统'] = SOURCE_SYSTEM
    output_df['来源单据编号'] = source_df['流程编号']
    output_df['申请日期'] = source_df['申请日期'].map(c.format_date)
    output_df['单据类型'] = DOCUMENT_TYPE
    output_df['申请人工号'] = source_df['申请人工号']
    output_df['申请人姓名'] = source_df['申请人']
    output_df['核算主体编号'] = source_df['公司主体'].map(lambda value: _lookup_by_name(entity_map, value))
    output_df['核算主体描述'] = source_df['公司主体']
    output_df['备注'] = source_df['备注'].map(lambda value: _text(value)[:150])
    output_df['合同号'] = ''
    output_df['合同收支计划行'] = ''
    output_df['收款方编码'] = [vendor_field(index, 'code') for index in source_df.index]
    output_df['收款方描述'] = [
        vendor_field(index, 'name') or supplier_name
        for index, supplier_name in zip(source_df.index, source_df['供应商'])
    ]
    output_df['银行账号'] = ''
    output_df['计划付款日期'] = ''
    output_df['银行转账备注'] = ''
    output_df['实际已支付金额'] = amount.map(c.round_amount)
    output_df['费用项目编码'] = source_df['预算科目'].map(lambda value: _subject_item(subject_lookup, value, 0))
    output_df['费用项目描述'] = source_df['预算科目'].map(lambda value: _subject_item(subject_lookup, value, 1))
    output_df['主播房间号'] = ''
    output_df['报账币种'] = 'CNY'
    output_df['报账金额（支付币种）'] = amount.map(c.round_amount)
    output_df['泛微费用项目编码'] = source_df['预算科目'].where(source_df['预算科目'].notna(), '')
    return _apply_order_project_columns(output_df, source_df)


# ============================ 只转入外部成本 ============================
def read_external_cost_source():
    stats = _query_fw(EXTERNAL_COST_STATS_SQL).iloc[0]
    print(f"[应付期初-只转入外部成本] SQL过滤: 来源=赛事只转入外部成本(ly=5)+MCN只转入外部成本(ly=2) 且 流程状态∈审批中/审批完成 且 未作废 且 申请日期>={DATE_FROM}")
    print(f"  保留内部收支 {int(stats['document_count'] or 0)} 单; "
          f"赛事 {int(stats['event_document_count'] or 0)} 单; "
          f"MCN {int(stats['mcn_document_count'] or 0)} 单; "
          f"转入金额合计 {float(stats['in_amount_total'] or 0):.2f}; "
          f"转出金额合计 {float(stats['out_amount_total'] or 0):.2f}")
    source_df = _query_fw(EXTERNAL_COST_SOURCE_SQL)
    print('[应付期初-只转入外部成本] SQL费用明细行数:', len(source_df))
    return resolve_external_cost_values(allocate_external_cost_amounts(source_df))


def resolve_external_cost_values(source_df):
    df = source_df.copy()
    employee_map = c.build_fw_employee_info_map_for_ids(df['申请人ID'])
    company_ids = pd.concat([df['转入公司主体ID'], df['转出公司主体ID']], ignore_index=True)
    company_map = c.build_fw_company_name_map_for_ids(company_ids)
    currency_map = c.build_fw_currency_name_map_for_ids(df['付款币种ID'])
    project_values = pd.concat([
        df['转入项目编号ID'], df['转出项目编号ID'],
        df['转入MCN赛事项目编号ID'], df['转出MCN赛事项目编号ID'],
    ], ignore_index=True)
    project_info_map = build_fw_project_info_map_for_ids(project_values)
    project_code_map = {
        project_id: info.get('code', '')
        for project_id, info in project_info_map.items()
    }
    project_name_map = {
        project_id: info.get('name', '')
        for project_id, info in project_info_map.items()
    }

    df['申请人'] = df['申请人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('name', ''))
    df['申请人工号'] = df['申请人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('code', ''))
    df['转入公司主体'] = df['转入公司主体ID'].map(lambda value: _first_browser_value(company_map, value))
    df['转出公司主体'] = df['转出公司主体ID'].map(lambda value: _first_browser_value(company_map, value))
    df['转入项目编号'] = [
        _first_non_blank(
            _first_browser_value(project_code_map, in_project),
            _first_browser_value(project_code_map, in_mcn_project),
            _text(in_project),
            _text(in_mcn_project),
        )
        for in_project, in_mcn_project in zip(df['转入项目编号ID'], df['转入MCN赛事项目编号ID'])
    ]
    df['转出项目编号'] = [
        _first_non_blank(
            _first_browser_value(project_code_map, out_project),
            _first_browser_value(project_code_map, out_mcn_project),
            _text(out_project),
            _text(out_mcn_project),
        )
        for out_project, out_mcn_project in zip(df['转出项目编号ID'], df['转出MCN赛事项目编号ID'])
    ]
    df['转入项目名称'] = [
        _first_non_blank(
            in_project_name,
            _first_browser_value(project_name_map, in_project),
            _first_browser_value(project_name_map, in_mcn_project),
        )
        for in_project_name, in_project, in_mcn_project
        in zip(df['转入项目名称'], df['转入项目编号ID'], df['转入MCN赛事项目编号ID'])
    ]
    df['转出项目名称'] = [
        _first_non_blank(
            out_project_name,
            _first_browser_value(project_name_map, out_project),
            _first_browser_value(project_name_map, out_mcn_project),
        )
        for out_project_name, out_project, out_mcn_project
        in zip(df['转出项目名称'], df['转出项目编号ID'], df['转出MCN赛事项目编号ID'])
    ]
    df['付款币种'] = df['付款币种ID'].map(lambda value: currency_map.get(c.format_code(value), '人民币'))
    df['预算科目'] = _resolve_subject_paths(df['预算科目ID'])
    df['费用单据号'] = [
        _first_non_blank(fee_doc_no, fee_id, flow_no)
        for fee_doc_no, fee_id, flow_no in zip(df['费用单据号'], df['费用单ID'], df['流程编号'])
    ]
    return df


def _build_external_side_rows(source_df, side):
    is_in = side == 'in'
    entity_col = '转入公司主体' if is_in else '转出公司主体'
    project_col = '转入项目编号' if is_in else '转出项目编号'
    project_name_col = '转入项目名称' if is_in else '转出项目名称'
    amount_col = '转入金额' if is_in else '转出金额'
    sign = 1 if is_in else -1
    side_label = '转入方' if is_in else '转出方'

    df = source_df.copy()
    df['方向'] = side_label
    df['公司主体'] = df[entity_col]
    df['项目编号'] = df[project_col]
    df['项目名称'] = df[project_name_col]
    df['金额'] = pd.to_numeric(df[amount_col], errors='coerce') * sign
    return df


def build_external_cost_output(source_df):
    if source_df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    expanded_df = pd.concat([
        _build_external_side_rows(source_df, 'in'),
        _build_external_side_rows(source_df, 'out'),
    ], ignore_index=True)

    vendor_map = c.build_vendor_map()
    virtual_vendor_code = vendor_map.get(c.normalize_name(VIRTUAL_VENDOR_NAME), '')
    entity_map = c.build_accounting_entity_map_for_names(expanded_df['公司主体'])
    subject_lookup = c.build_subject_map()

    output_df = pd.DataFrame(index=expanded_df.index)
    output_df['来源系统'] = SOURCE_SYSTEM
    output_df['来源单据编号'] = expanded_df['流程编号']
    output_df['申请日期'] = expanded_df['申请日期'].map(c.format_date)
    output_df['单据类型'] = DOCUMENT_TYPE
    output_df['申请人工号'] = expanded_df['申请人工号']
    output_df['申请人姓名'] = expanded_df['申请人']
    output_df['核算主体编号'] = expanded_df['公司主体'].map(lambda value: _lookup_by_name(entity_map, value))
    output_df['核算主体描述'] = expanded_df['公司主体']
    output_df['备注'] = [
        f'{side}:{fee_doc_no}'[:150]
        for side, fee_doc_no in zip(expanded_df['方向'], expanded_df['费用单据号'])
    ]
    output_df['合同号'] = ''
    output_df['合同收支计划行'] = ''
    output_df['收款方编码'] = virtual_vendor_code
    output_df['收款方描述'] = VIRTUAL_VENDOR_NAME
    output_df['银行账号'] = ''
    output_df['计划付款日期'] = ''
    output_df['银行转账备注'] = ''
    output_df['实际已支付金额'] = expanded_df['金额'].map(c.round_amount)
    output_df['费用项目编码'] = expanded_df['预算科目'].map(lambda value: _subject_item(subject_lookup, value, 0))
    output_df['费用项目描述'] = expanded_df['预算科目'].map(lambda value: _subject_item(subject_lookup, value, 1))
    output_df['主播房间号'] = ''
    output_df['报账币种'] = expanded_df['付款币种'].map(c.to_iso_currency)
    output_df['报账金额（支付币种）'] = expanded_df['金额'].map(c.round_amount)
    output_df['泛微费用项目编码'] = expanded_df['预算科目'].where(expanded_df['预算科目'].notna(), '')

    # 供问题清单复用。
    expanded_df['收款方描述'] = VIRTUAL_VENDOR_NAME
    return _apply_order_project_columns(output_df, expanded_df), expanded_df


def collect_external_cost_pair_check(expanded_df):
    """校验只转入外部成本:每个单据/明细的正数转入行应有对应负数转出行。"""
    if expanded_df.empty:
        return {
            '转入转出校验': pd.DataFrame([{
                '校验结果': '无数据',
                '说明': '只转入外部成本无输出行',
            }]),
        }

    df = expanded_df.copy()
    df['金额'] = pd.to_numeric(df['金额'], errors='coerce').fillna(0).round(2)
    df['配对明细ID'] = [
        _first_non_blank(detail_id, view_id, fee_doc_no)
        for detail_id, view_id, fee_doc_no in zip(df['明细ID'], df['费用明细视图ID'], df['费用单据号'])
    ]

    detail_rows = []
    failed_detail_keys = set()
    for (doc_no, detail_id), group in df.groupby(['流程编号', '配对明细ID'], sort=False):
        positive = group[group['金额'] > 0]
        negative = group[group['金额'] < 0]
        zero_count = int((group['金额'] == 0).sum())
        if len(positive) == 0 and len(negative) == 0:
            continue
        positive_total = round(float(positive['金额'].sum()), 2)
        negative_total = round(float(negative['金额'].sum()), 2)
        diff = round(positive_total + negative_total, 2)
        is_ok = (
            len(positive) == 1
            and len(negative) == 1
            and abs(diff) <= 0.01
        )
        if is_ok:
            continue

        failed_detail_keys.add((doc_no, detail_id))
        sample = group.iloc[0]
        detail_rows.append({
            '来源单据编号': doc_no,
            '配对明细ID': detail_id,
            '费用单据号': _text(sample.get('费用单据号', '')),
            '预算科目': _text(sample.get('预算科目', '')),
            '正数行数': len(positive),
            '负数行数': len(negative),
            '零金额行数': zero_count,
            '正数金额合计': positive_total,
            '负数金额合计': negative_total,
            '正负合计差额': diff,
            '校验结果': '通过' if is_ok else '不通过',
        })

    summary_rows = []
    for doc_no, group in df.groupby('流程编号', sort=False):
        positive = group[group['金额'] > 0]
        negative = group[group['金额'] < 0]
        zero_count = int((group['金额'] == 0).sum())
        positive_total = round(float(positive['金额'].sum()), 2)
        negative_total = round(float(negative['金额'].sum()), 2)
        diff = round(positive_total + negative_total, 2)
        detail_ids = set(zip(group['流程编号'], group['配对明细ID']))
        failed_count = len(detail_ids & failed_detail_keys)
        is_ok = (
            len(positive) == len(negative)
            and failed_count == 0
            and abs(diff) <= 0.01
        )
        summary_rows.append({
            '来源单据编号': doc_no,
            '正数行数': len(positive),
            '负数行数': len(negative),
            '零金额行数': zero_count,
            '正数金额合计': positive_total,
            '负数金额合计': negative_total,
            '正负合计差额': diff,
            '未配对明细数': failed_count,
            '校验结果': '通过' if is_ok else '不通过',
        })

    sheets = {
        '转入转出校验': pd.DataFrame(summary_rows),
    }
    if detail_rows:
        sheets['转入转出校验异常明细'] = pd.DataFrame(detail_rows)
    return sheets


def run():
    # 1. SQL 读取并解析三个来源
    base_source_df = read_base_source()
    batch_source_df = read_batch_source()
    external_source_df = read_external_cost_source()

    # 2. 构建三个 sheet 输出
    base_issue_source_df = _with_resolved_project_fields(base_source_df)
    base_output_df = _apply_order_project_columns(
        base_payment.build_output(base_source_df),
        base_issue_source_df,
    )
    batch_output_df = build_batch_output(batch_source_df)
    external_output_df, external_issue_source_df = build_external_cost_output(external_source_df)
    print('[应付期初-期初对公付款单导入] 输出明细行数:', len(base_output_df))
    print('[应付期初-批量费用流程] 输出明细行数:', len(batch_output_df))
    print('[应付期初-只转入外部成本] 输出明细行数:', len(external_output_df))

    # 3. 填充率
    required_cols = c.required_columns(RULE_SHEET, RULE_TABLE)
    print('— 期初对公付款单导入 填充率 —')
    c.report_fill(base_output_df, required_cols)
    print('— 批量费用流程 填充率 —')
    c.report_fill(batch_output_df, required_cols)
    print('— 只转入外部成本 填充率 —')
    c.report_fill(external_output_df, required_cols)

    # 4. 写入模板:同一个 Excel 内写入三个数据 sheet,lov 页保留
    write_output_workbook(base_output_df, batch_output_df, external_output_df)
    print('已写出:', OUTPUT_FILE)

    # 5. 问题清单
    exception_sheets = {}

    base_sheets = {'必输字段未达100%': c.fill_summary(
        base_output_df, required_cols, RULE_SHEET, RULE_TABLE)}
    base_sheets.update(c.collect_field_issues(
        base_output_df, base_issue_source_df, required_cols, base_payment.ISSUE_SOURCE_FIELDS))
    base_sheets = _enrich_missing_order_issue(base_sheets, base_output_df, base_issue_source_df)
    base_sheets.update(collect_order_mapping_issues(base_issue_source_df))
    exception_sheets.update({f'期初对公付款单导入_{name}': df for name, df in base_sheets.items()})

    batch_sheets = {'必输字段未达100%': c.fill_summary(
        batch_output_df, required_cols, RULE_SHEET, RULE_TABLE)}
    batch_sheets.update(c.collect_field_issues(
        batch_output_df, batch_source_df, required_cols, BATCH_ISSUE_SOURCE_FIELDS))
    batch_sheets = _enrich_missing_order_issue(batch_sheets, batch_output_df, batch_source_df)
    batch_sheets.update(collect_order_mapping_issues(batch_source_df))
    exception_sheets.update({f'批量费用流程_{name}': df for name, df in batch_sheets.items()})

    external_sheets = {'必输字段未达100%': c.fill_summary(
        external_output_df, required_cols, RULE_SHEET, RULE_TABLE)}
    external_sheets.update(c.collect_field_issues(
        external_output_df, external_issue_source_df, required_cols, EXTERNAL_COST_ISSUE_SOURCE_FIELDS))
    external_sheets = _enrich_missing_order_issue(external_sheets, external_output_df, external_issue_source_df)
    external_sheets.update(collect_order_mapping_issues(external_issue_source_df))
    external_sheets.update(collect_external_cost_pair_check(external_issue_source_df))
    exception_sheets.update({f'只转入外部成本_{name}': df for name, df in external_sheets.items()})

    c.write_exceptions(EXCEPTION_FILE, exception_sheets)
    print('已写出:', EXCEPTION_FILE, '| 各清单条数:', {k: len(v) for k, v in exception_sheets.items()})


if __name__ == '__main__':
    run()
