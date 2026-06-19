# -*- coding: utf-8 -*-
"""预付期初 —— 供应商预付款单 / 零工预付款单(DB 直连版)。

处理流程:
1. 校验泛微字段字典,避免 SQL 字段名/含义写错。
2. 从泛微库读取供应商预付 uf_yfkxx/uf_yfkxx_dt1,以及零工付款 uf_lgptfk + 原流程收款人明细/预算项明细。
3. 只对必须跨表/跨系统的 ID 做批量解析,例如人员、公司主体、币种、预算科目、供应商、银行账号。
4. 按导入模版两个 tab 逐列生成输出,字段旁标注取值来源。

跑法:在项目根执行  python run.py ap_prepayment_opening_db
"""
import sys
import re
import os
from pathlib import Path

import pandas as pd

if __package__ is None or __package__ == '':
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from etl import common as c

# ============================ 文件 / 模板 ============================
TASK_NAME = 'ap_prepayment_opening_db'
TEMPLATE_DIR = c.TPL_DIR / 'ap_prepayment_opening'
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
DATE_SUFFIX = c.today_suffix()

TEMPLATE_FILE = TEMPLATE_DIR / '英雄期初预付款单导入模版.xlsx'
OUTPUT_FILE = OUTPUT_DIR / f'英雄期初预付款单导入_预付期初_{DATE_SUFFIX}.xlsx'
EXCEPTION_FILE = OUTPUT_DIR / f'未匹配清单_预付期初_{DATE_SUFFIX}.xlsx'
SUPPLIER_VENDOR_MISSING_FILE = OUTPUT_DIR / f'Hand按ID查不到的供应商_预付期初_{DATE_SUFFIX}.xlsx'

TEMPLATE_SHEET_SUPPLIER = '期初供应商预付款单&期初投资付款单导入'
TEMPLATE_SHEET_GIG = '期初灵工预付款单导入'
RULE_SHEET = '预付期初'
RULE_TABLE_SUPPLIER = '期初供应商预付款单&期初投资付款单导入'
RULE_TABLE_GIG = '期初灵工预付款单导入'

DOCUMENT_TYPE = 'JK01-2'
GIG_DOCUMENT_TYPE = 'PP01-2'
GIG_PLATFORM_VENDOR = {'云账户': 'V-C-CN-HR-PAY-0001', '赛利得': 'V-C-CN-OT-OTH-6573'}

SUPPLIER_OUTPUT_COLUMNS = [
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
    '保证金标志',
    '收款方编码',
    '收款方描述',
    '银行账号',
    '计划付款日期',
    '银行转账备注',
    '费用项目编码',
    '费用项目描述',
    '主播房间号',
    '预付款支付币种',
    '预付款金额（支付币种）',
    '已到票核销金额（支付币种）',
    '已付未核（支付币种）',
    '泛微项目编号',
]
GIG_OUTPUT_COLUMNS = [
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
    '备注_单头',
    '灵工平台收款方编码',
    '合同号',
    '合同收支计划行',
    '保证金标志',
    '计划付款日期',
    '银行转账备注',
    '费用项目编码',
    '费用项目描述',
    '收款方类别',
    '收款方编码',
    '备注',
    '银行账号',
    '预付款支付币种',
    '预付款金额（支付币种）',
    '传送状态',
    '支付状态',
    '退款状态',
    '核销状态',
    '泛微项目编号',
]

SUPPLIER_ISSUE_SOURCE_FIELDS = {
    '申请人工号': '填单人',
    '收款方编码': '付款对象',
    '核算主体编号': '开票单位',
    '费用项目编码': '预算科目',
    '预付款支付币种': '付款币种',
    '订单编号': '项目编号',
}
GIG_ISSUE_SOURCE_FIELDS = {
    '申请人工号': '经办人',
    '灵工平台收款方编码': '收款方文本',
    '核算主体编号': '公司主体',
    '费用项目编码': '预算科目',
    '收款方编码': '实际收款方',
    '订单编号': '项目编号',
}

FW_SUPPLIER_TABLE = 'uf_yfkxx'
FW_SUPPLIER_DETAIL_TABLE = 'uf_yfkxx_dt1'
FW_GIG_HEADER_TABLE = 'uf_lgptfk'
FW_GIG_WORKFLOW_TABLE = 'formtable_main_279'
FW_GIG_BUDGET_TABLE = 'formtable_main_279_dt3'
FW_GIG_RECIPIENT_TABLE = 'formtable_main_279_dt4'
FW_PROJECT_TABLE = 'uf_xtyyxmkp'
ORDER_MAPPING_ENV = 'PROJECT_ORDER_MAPPING_XLSX'
ORDER_MAPPING_XLSX_NAME = '业财项目_项目&订单清洗_0618.xlsx'
ORDER_MAPPING_SHEETS = {
    '赛事本部订单': '赛事订单主表_清洗后',
    'MCN本部订单': 'MCN订单主表_清洗后',
    '全量协同订单': '全量协同订单_清洗后',
}


# ============================ 枚举 / 过滤口径 ============================
DATE_FROM = '2022-01-01'
APPROVED_STATUS_CODE = 2
VOID_CODE = 0
DEPOSIT_PAYMENT_NATURE_CODES = {'0', '1'}  # uf_yfkxx.fkxz:0=押金,1=质保金

FLOW_STATUS_MEANINGS = {
    0: '未提交',
    1: '审批中',
    2: '审批完成',
}
VOID_FLAG_MEANINGS = {
    0: '是',
    1: '否',
}
PAYMENT_NATURE_MEANINGS = {
    0: '押金',
    1: '质保金',
    2: '一般',
}


# ============================ 泛微源 SQL ============================
SUPPLIER_SOURCE_SQL = """
SELECT
    m.id AS `ID`,
    d.id AS `明细ID`,
    m.lcbh AS `流程编号`,
    m.sqrq AS `申请日期`,
    m.tdr AS `填单人ID`,
    m.xmbh AS `项目编号ID`,
    m.kpdw AS `开票单位ID`,
    m.bz AS `备注`,
    m.xght AS `相关合同ID`,
    m.fkdx AS `付款对象ID`,
    m.yhkh AS `银行卡号ID`,
    m.fkbz AS `付款币种ID`,
    m.fkje AS `付款金额`,
    m.fkxz AS `付款性质ID`,
    m.sycxtkje AS `剩余冲销/退款金额`,
    d.yfje AS `预付金额`,
    d.yskm AS `预算科目ID`
FROM uf_yfkxx m
JOIN uf_yfkxx_dt1 d ON d.mainid = m.id
WHERE m.sqrq >= %(date_from)s
  AND m.lczt = %(approved_status_code)s
  AND (m.sfzf IS NULL OR m.sfzf <> %(void_code)s)
ORDER BY m.id, d.id
"""

SUPPLIER_STATS_SQL = """
SELECT
    SUM(CASE
        WHEN m.sqrq >= %(date_from)s
         AND m.lczt = %(approved_status_code)s
        THEN 1 ELSE 0 END) AS matched_count,
    SUM(CASE
        WHEN m.sqrq >= %(date_from)s
         AND m.lczt = %(approved_status_code)s
         AND m.sfzf = %(void_code)s
        THEN 1 ELSE 0 END) AS void_count,
    SUM(CASE
        WHEN m.sqrq >= %(date_from)s
         AND m.lczt = %(approved_status_code)s
         AND (m.sfzf IS NULL OR m.sfzf <> %(void_code)s)
        THEN 1 ELSE 0 END) AS kept_count
FROM uf_yfkxx m
"""

GIG_RECIPIENT_SOURCE_SQL = """
SELECT
    h.id AS `建模付款ID`,
    w.id AS `流程付款ID`,
    d.id AS `收款人明细ID`,
    h.lcbh AS `流程编号`,
    h.sqrq AS `申请日期`,
    h.lcfsrq AS `流程发生日期`,
    w.requestId AS `流程请求ID`,
    h.lczsjqqid AS `流程转数据请求ID`,
    COALESCE(NULLIF(CAST(h.xmbh AS CHAR), ''), NULLIF(CAST(w.xmbh AS CHAR), '')) AS `项目编号ID`,
    COALESCE(NULLIF(h.xmmc, ''), NULLIF(w.xmmc, '')) AS `项目名称`,
    h.jbr AS `经办人ID`,
    h.gszt AS `公司主体ID`,
    COALESCE(NULLIF(h.skfwb, ''), w.skfmc) AS `收款方文本`,
    h.bz AS `备注`,
    w.htmc AS `合同名称`,
    h.xght AS `相关合同ID`,
    h.yjfkrq AS `预计付款日期`,
    d.skf AS `实际收款方`,
    d.sjh AS `手机号`,
    d.sfzh AS `身份证号`,
    d.yhzh AS `银行账号`,
    d.je AS `税前应付金额`,
    d.sl AS `付给三方平台金额`,
    d.se AS `付款平台税额`,
    d.zzs AS `增值税`,
    d.fjs AS `附加税`,
    d.grsds AS `个人所得税`,
    d.ygdsje AS `预估到手金额`,
    d.gyszxx AS `库内外`,
    d.mainid AS `明细mainid`
FROM uf_lgptfk h
JOIN formtable_main_279 w
  ON CAST(w.requestId AS CHAR) = h.lczsjqqid
JOIN formtable_main_279_dt4 d
  ON d.mainid = w.id
WHERE h.sqrq >= %(date_from)s
  AND h.lczt = %(approved_status_code)s
  AND (h.sfzf IS NULL OR h.sfzf <> %(void_code)s)
ORDER BY h.id, d.id
"""

GIG_BUDGET_SOURCE_SQL = """
SELECT
    h.id AS `建模付款ID`,
    w.id AS `流程付款ID`,
    b.id AS `预算明细ID`,
    h.lcbh AS `流程编号`,
    b.yskm AS `预算科目ID`,
    b.fysx AS `费用事项`,
    b.fkje AS `预算项金额`,
    b.rmbje AS `预算项人民币金额`
FROM uf_lgptfk h
JOIN formtable_main_279 w
  ON CAST(w.requestId AS CHAR) = h.lczsjqqid
JOIN formtable_main_279_dt3 b
  ON b.mainid = w.id
WHERE h.sqrq >= %(date_from)s
  AND h.lczt = %(approved_status_code)s
  AND (h.sfzf IS NULL OR h.sfzf <> %(void_code)s)
ORDER BY h.id, b.id
"""

GIG_STATS_SQL = """
SELECT
    SUM(CASE
        WHEN h.sqrq >= %(date_from)s
         AND h.lczt = %(approved_status_code)s
        THEN 1 ELSE 0 END) AS matched_count,
    SUM(CASE
        WHEN h.sqrq >= %(date_from)s
         AND h.lczt = %(approved_status_code)s
         AND h.sfzf = %(void_code)s
        THEN 1 ELSE 0 END) AS void_count,
    SUM(CASE
        WHEN h.sqrq >= %(date_from)s
         AND h.lczt = %(approved_status_code)s
         AND (h.sfzf IS NULL OR h.sfzf <> %(void_code)s)
        THEN 1 ELSE 0 END) AS kept_count
FROM uf_lgptfk h
"""


# 运行前校验字段真实含义。主表字段 detail_table 用空字符串。
EXPECTED_SUPPLIER_FIELDS = {
    '': {
        'lcbh': '流程编号',
        'sqrq': '申请日期',
        'tdr': '填单人',
        'xmbh': '项目编号',
        'kpdw': '开票单位',
        'bz': '备注',
        'xght': '相关合同',
        'fkdx': '付款对象',
        'yhkh': '银行卡号',
        'fkbz': '付款币种',
        'fkje': '付款金额',
        'fkxz': '付款性质',
        'sycxtkje': '剩余冲销/退款金额',
        'lczt': '流程状态',
        'sfzf': '是否作废',
    },
    FW_SUPPLIER_DETAIL_TABLE: {
        'yfje': '预付金额',
        'yskm': '预算科目',
    },
}
EXPECTED_GIG_HEADER_FIELDS = {
    '': {
        'lcbh': '流程编号',
        'sqrq': '申请日期',
        'lcfsrq': '流程发生日期',
        'lczsjqqid': '流程转数据请求ID',
        'xmbh': '项目编号',
        'xmmc': '项目名称',
        'jbr': '经办人',
        'gszt': '公司主体',
        'skfwb': '收款方-文本',
        'bz': '备注',
        'xght': '相关合同',
        'yjfkrq': '预计付款日期',
        'lczt': '流程状态',
        'sfzf': '是否作废',
    },
}
EXPECTED_GIG_WORKFLOW_FIELDS = {
    '': {
        'lcbh': '流程编号',
        'xmbh': '项目编号',
        'xmmc': '项目名称',
        'skfmc': '收款方名称',
        'htmc': '合同名称',
    },
    FW_GIG_BUDGET_TABLE: {
        'fysx': '费用事项',
        'yskm': '预算科目',
        'fkje': '付给三方平台金额',
        'rmbje': '付款人民币金额',
    },
    FW_GIG_RECIPIENT_TABLE: {
        'skf': '收款方',
        'sjh': '手机号',
        'sfzh': '身份证号',
        'yhzh': '银行账号',
        'je': '税前应付金额',
        'sl': '付给三方平台金额',
        'se': '付款平台税额',
        'zzs': '增值税',
        'fjs': '附加税',
        'grsds': '个人所得税',
        'ygdsje': '预估到手金额',
    },
}


# ============================ DB 查询小工具 ============================
def _query_fw(sql):
    """查询泛微库,统一带上本任务过滤参数。"""
    return c.query_db('FW', 'vspn_xtyy', sql, {
        'date_from': DATE_FROM,
        'approved_status_code': APPROVED_STATUS_CODE,
        'void_code': VOID_CODE,
    })


# ============================ 源值解析 ============================
def _text(value):
    if pd.isna(value):
        return ''
    text = str(value).strip()
    return '' if text in ('', 'nan', 'None', 'NaT') else text


_PROJECT_ORDER_MAPPING_CACHE = None


def _find_project_order_mapping_file():
    configured = os.getenv(ORDER_MAPPING_ENV, '').strip()
    if configured:
        path = Path(configured)
        if path.exists():
            return path
        print(f'[预付期初-供应商预付款-订单映射] {ORDER_MAPPING_ENV} 指向的文件不存在:', configured)

    search_dirs = [
        c.SRC_DIR / 'other_cleaned_data',
        c.SRC_DIR / 'project_order',
        c.SRC_DIR,
        Path.home() / 'Downloads',
    ]
    for search_dir in search_dirs:
        path = search_dir / ORDER_MAPPING_XLSX_NAME
        if path.exists():
            return path
    return None


def _dedupe_headers(headers):
    seen = {}
    deduped_headers = []
    for header in headers:
        header = header or '未命名'
        if header in seen:
            seen[header] += 1
            header = f'{header}.{seen[header]}'
        else:
            seen[header] = 0
        deduped_headers.append(header)
    return deduped_headers


def _read_cleaned_order_sheet(path, sheet_name, required_columns):
    raw_df = pd.read_excel(path, sheet_name=sheet_name, dtype=str, keep_default_na=False, header=None)
    required = set(required_columns)

    header_index = None
    headers = []
    for index, row in raw_df.iterrows():
        normalized = [_text(value).replace('\n', '') for value in row.tolist()]
        if required.issubset(set(normalized)):
            header_index = index
            headers = normalized
            break
    if header_index is None:
        raise ValueError(f'{sheet_name} 未找到表头列: {sorted(required)}')

    df = raw_df.iloc[header_index + 1:].copy()
    df.columns = _dedupe_headers(headers)
    return df


def _append_order_mapping_rows(
        rows, df, source_name, fanwei_column, order_column, title_column,
        project_column, project_name_column=''):
    for _, row in df.iterrows():
        fanwei_project = _text(row.get(fanwei_column, ''))
        order_code = _text(row.get(order_column, ''))
        if not fanwei_project or not order_code:
            continue
        rows.append({
            '泛微项目编号': fanwei_project,
            '订单编号': order_code,
            '订单标题': _text(row.get(title_column, '')),
            '项目编号': _text(row.get(project_column, '')),
            '项目名称': _text(row.get(project_name_column, '')) if project_name_column else '',
            '映射来源': source_name,
        })


def _append_order_presence_rows(
        rows, df, source_name, fanwei_column, order_column, title_column):
    """收集订单表中出现过的泛微项目编号,用于解释未匹配原因。

    这里和 _append_order_mapping_rows 的区别是:
    - 正式映射只保留「泛微项目编号 + 订单编号」都不为空的记录。
    - presence 检查只要求泛微项目编号存在,即使订单编号为空也保留。

    这样未匹配时可以区分两类情况:
    1. 0618 三个订单表里完全没有这个泛微项目编号。
    2. 订单表里有这个泛微项目编号,但订单编号为空,所以不能写入导入模板。
    """
    for _, row in df.iterrows():
        # 不同订单表的泛微项目列名不同,通过 fanwei_column 参数统一读取。
        fanwei_project = _text(row.get(fanwei_column, ''))
        if not fanwei_project:
            continue

        # 订单编号允许为空:为空时也要保留,用于后续生成「订单编号为空」的未匹配原因。
        rows.append({
            '泛微项目编号': fanwei_project,
            '订单编号': _text(row.get(order_column, '')),
            '订单标题': _text(row.get(title_column, '')),
            '映射来源': source_name,
        })


def _build_project_order_presence(mapping_file):
    """读取 0618 项目&订单清洗表,汇总所有出现过的泛微项目编号。

    返回的数据只用于「订单映射_未匹配」sheet 的原因判断,不参与一对一映射。
    一对一/多候选的正式映射仍由 _build_project_order_candidates 生成。
    """
    rows = []

    # 本部订单:赛事订单主表、MCN订单主表都使用同一组列名。
    for source_name, sheet_name in (
            ('赛事本部订单', ORDER_MAPPING_SHEETS['赛事本部订单']),
            ('MCN本部订单', ORDER_MAPPING_SHEETS['MCN本部订单'])):
        order_df = _read_cleaned_order_sheet(
            mapping_file,
            sheet_name,
            ['泛微项目编号', '订单编号', '订单标题'],
        )
        _append_order_presence_rows(
            rows,
            order_df,
            source_name,
            fanwei_column='泛微项目编号',
            order_column='订单编号',
            title_column='订单标题',
        )

    # 协同订单:同一行里既有下单方泛微项目编号,也有协同方泛微项目编号。
    # 两个编号都可能被供应商/灵工单据引用,所以分别纳入 presence 检查。
    coop_df = _read_cleaned_order_sheet(
        mapping_file,
        ORDER_MAPPING_SHEETS['全量协同订单'],
        ['泛微下单方项目编号', '泛微协同方项目编号', '协同订单编号', '协同订单标题'],
    )
    _append_order_presence_rows(
        rows,
        coop_df,
        '协同订单-下单方项目',
        fanwei_column='泛微下单方项目编号',
        order_column='协同订单编号',
        title_column='协同订单标题',
    )
    _append_order_presence_rows(
        rows,
        coop_df,
        '协同订单-协同方项目',
        fanwei_column='泛微协同方项目编号',
        order_column='协同订单编号',
        title_column='协同订单标题',
    )

    # 去重后按统一列返回,便于后续按「泛微项目编号」分组判断是否出现过。
    return pd.DataFrame(
        rows,
        columns=['泛微项目编号', '订单编号', '订单标题', '映射来源'],
    ).drop_duplicates()


def _build_project_order_candidates(mapping_file):
    rows = []

    for source_name, sheet_name in (
            ('赛事本部订单', ORDER_MAPPING_SHEETS['赛事本部订单']),
            ('MCN本部订单', ORDER_MAPPING_SHEETS['MCN本部订单'])):
        order_df = _read_cleaned_order_sheet(
            mapping_file,
            sheet_name,
            ['泛微项目编号', '订单编号', '订单标题', '项目编号'],
        )
        _append_order_mapping_rows(
            rows,
            order_df,
            source_name,
            fanwei_column='泛微项目编号',
            order_column='订单编号',
            title_column='订单标题',
            project_column='项目编号',
            project_name_column='项目名称',
        )

    coop_df = _read_cleaned_order_sheet(
        mapping_file,
        ORDER_MAPPING_SHEETS['全量协同订单'],
        ['泛微下单方项目编号', '泛微协同方项目编号', '协同订单编号', '协同订单标题', '协同方项目编号'],
    )
    _append_order_mapping_rows(
        rows,
        coop_df,
        '协同订单-下单方项目',
        fanwei_column='泛微下单方项目编号',
        order_column='协同订单编号',
        title_column='协同订单标题',
        project_column='协同方项目编号',
        project_name_column='协同方项目名称',
    )
    _append_order_mapping_rows(
        rows,
        coop_df,
        '协同订单-协同方项目',
        fanwei_column='泛微协同方项目编号',
        order_column='协同订单编号',
        title_column='协同订单标题',
        project_column='协同方项目编号',
        project_name_column='协同方项目名称',
    )

    return pd.DataFrame(
        rows,
        columns=['泛微项目编号', '订单编号', '订单标题', '项目编号', '项目名称', '映射来源'],
    ).drop_duplicates()


def load_project_order_mapping():
    """读取项目&订单清洗后的 Excel,返回一对一映射和一对多候选。"""
    global _PROJECT_ORDER_MAPPING_CACHE
    if _PROJECT_ORDER_MAPPING_CACHE is not None:
        return _PROJECT_ORDER_MAPPING_CACHE

    mapping_file = _find_project_order_mapping_file()
    if mapping_file is None:
        print('[预付期初-供应商预付款-订单映射] 未找到项目&订单清洗后 Excel,订单字段保持为空。')
        _PROJECT_ORDER_MAPPING_CACHE = ({}, {}, None)
        return _PROJECT_ORDER_MAPPING_CACHE

    mapping_df = _build_project_order_candidates(mapping_file)

    safe_map = {}
    ambiguous_map = {}
    for project_code, group in mapping_df.groupby('泛微项目编号', sort=False):
        orders = group['订单编号'].drop_duplicates()
        if len(orders) == 1:
            row = group.iloc[0]
            safe_map[project_code] = {
                '订单编号': row['订单编号'],
                '订单标题': row['订单标题'],
                '项目编号': row['项目编号'],
                '项目名称': row['项目名称'],
                '映射来源': row['映射来源'],
            }
        else:
            ambiguous_map[project_code] = group.to_dict('records')

    print(
        '[预付期初-供应商预付款-订单映射] 使用:',
        mapping_file,
        '| 候选记录数:', len(mapping_df),
        '| 一对一项目数:', len(safe_map),
        '| 多候选项目数:', len(ambiguous_map),
    )
    _PROJECT_ORDER_MAPPING_CACHE = (safe_map, ambiguous_map, mapping_file)
    return _PROJECT_ORDER_MAPPING_CACHE


def _order_mapping_value(project_code, field):
    safe_map, _, _ = load_project_order_mapping()
    return safe_map.get(_text(project_code), {}).get(field, '')


def collect_order_mapping_issues(merged_df):
    """输出项目->订单映射的未匹配和多候选清单。"""
    safe_map, ambiguous_map, mapping_file = load_project_order_mapping()
    issue_columns = {
        '来源单据编号': merged_df['流程编号'].map(_text),
        '泛微项目编号': merged_df['项目编号'].map(_text),
    }
    if '项目编号ID' in merged_df.columns:
        issue_columns['泛微项目ID'] = merged_df['项目编号ID'].map(_text)
    if '项目名称' in merged_df.columns:
        issue_columns['泛微项目名称'] = merged_df['项目名称'].map(_text)
    rows = pd.DataFrame(issue_columns).drop_duplicates()
    rows = rows[rows['泛微项目编号'] != '']

    sheets = {}
    if mapping_file is None:
        sheets['订单映射_文件缺失'] = pd.DataFrame([{
            '说明': f'未找到 {ORDER_MAPPING_XLSX_NAME},可通过环境变量 {ORDER_MAPPING_ENV} 指定清洗后 Excel。',
        }])
        return sheets

    ambiguous_rows = []
    for _, row in rows.iterrows():
        candidates = ambiguous_map.get(row['泛微项目编号'])
        if candidates:
            ambiguous_rows.append({
                '来源单据编号': row['来源单据编号'],
                '泛微项目编号': row['泛微项目编号'],
                '候选订单编号': '; '.join(_text(item.get('订单编号')) for item in candidates),
                '候选订单标题': '; '.join(_text(item.get('订单标题')) for item in candidates),
                '候选项目编号': '; '.join(_text(item.get('项目编号')) for item in candidates),
                '候选映射来源': '; '.join(_text(item.get('映射来源')) for item in candidates),
            })
    if ambiguous_rows:
        sheets['订单映射_多候选'] = pd.DataFrame(ambiguous_rows)

    order_presence_df = _build_project_order_presence(mapping_file)
    order_presence = {
        project_code: group.to_dict('records')
        for project_code, group in order_presence_df.groupby('泛微项目编号', sort=False)
    }

    mapped_projects = set(safe_map) | set(ambiguous_map)
    unmatched = rows[~rows['泛微项目编号'].isin(mapped_projects)].copy()
    if len(unmatched) > 0:
        unmatched_reasons = []
        unmatched_sources = []
        unmatched_order_values = []
        for _, row in unmatched.iterrows():
            appearances = order_presence.get(row['泛微项目编号'], [])
            if appearances:
                unmatched_reasons.append('0618订单表中有该泛微项目编号,但订单编号为空')
                unmatched_sources.append('; '.join(_text(item.get('映射来源')) for item in appearances))
                unmatched_order_values.append('; '.join(_text(item.get('订单编号')) or '(空)' for item in appearances))
            else:
                unmatched_reasons.append('0618赛事订单主表、MCN订单主表、全量协同订单均未出现该泛微项目编号')
                unmatched_sources.append('')
                unmatched_order_values.append('')
        unmatched.insert(2, '未匹配原因', unmatched_reasons)
        unmatched['0618订单表出现位置'] = unmatched_sources
        unmatched['0618订单编号字段值'] = unmatched_order_values
        sheets['订单映射_未匹配'] = unmatched
    return sheets


def _lookup_first_browser_value(mapping, value):
    for item_id in c.parse_browser_ids(value):
        mapped = mapping.get(item_id, '')
        if mapped:
            return mapped
    return ''


def build_fw_project_code_map_for_ids(project_values):
    """泛微项目浏览框 ID -> 泛微项目编号。

    供应商预付 uf_yfkxx.xmbh 存的是项目主数据 uf_xtyyxmkp.id,不是展示编码。
    """
    project_ids = c.clean_codes(
        project_id
        for value in project_values
        for project_id in c.parse_browser_ids(value)
    )
    if not project_ids:
        return {}
    project_df = c.query_db(
        'FW',
        'vspn_xtyy',
        'SELECT id, xmbh AS project_code '
        f'FROM {FW_PROJECT_TABLE} '
        f'WHERE id IN ({c.in_placeholders(project_ids)})',
        project_ids,
    )
    return {
        c.format_code(row['id']): _text(row['project_code'])
        for _, row in project_df.iterrows()
        if _text(row['project_code'])
    }


def build_fw_bank_account_map_for_ids(bank_account_values):
    """供应商银行账号浏览框ID -> 银行账号文本。

    uf_yfkxx.yhkh 存 browser.gysyhxx 的明细 ID,对应 uf_khgys_dt1.id。
    """
    bank_account_ids = c.clean_codes(
        bank_account_id
        for value in bank_account_values
        for bank_account_id in c.parse_browser_ids(value)
    )
    if not bank_account_ids:
        return {}
    bank_df = c.query_db(
        'FW',
        'vspn_xtyy',
        'SELECT id, yhzh bank_account '
        'FROM uf_khgys_dt1 '
        f'WHERE id IN ({c.in_placeholders(bank_account_ids)})',
        bank_account_ids,
    )
    return {
        c.format_code(row['id']): _text(row['bank_account'])
        for _, row in bank_df.iterrows()
        if _text(row['bank_account'])
    }


def _selected_supplier_name(value, supplier_status_map):
    selected_id = c.choose_fw_supplier_id(c.parse_browser_ids(value), supplier_status_map)
    return supplier_status_map.get(selected_id, {}).get('name', '')


def resolve_supplier_source_values(source_df):
    """基于供应商预付 SQL 返回的泛微 ID 字段补充输出需要的展示值。"""
    df = source_df.copy()
    employee_map = c.build_fw_employee_info_map_for_ids(df['填单人ID'])
    company_map = c.build_fw_company_name_map_for_ids(df['开票单位ID'])
    currency_map = c.build_fw_currency_name_map_for_ids(df['付款币种ID'])
    subject_map = c.build_fw_budget_subject_path_map_for_ids(df['预算科目ID'])
    contract_map = c.build_fw_contract_code_map_for_ids(df['相关合同ID'])
    bank_account_map = build_fw_bank_account_map_for_ids(df['银行卡号ID'])
    supplier_status_map = c.build_fw_supplier_status_map(df['付款对象ID'])
    project_code_map = build_fw_project_code_map_for_ids(df['项目编号ID'])

    # [主表] tdr(填单人) -> hrmresource / hrmjobtitles
    df['填单人'] = df['填单人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('name', ''))
    df['填单人工号'] = df['填单人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('code', ''))
    # [主表] kpdw(开票单位) -> uf_gstt.gsmc(公司主体名称)
    df['开票单位'] = df['开票单位ID'].map(lambda value: company_map.get(c.format_code(value), ''))
    # [主表] fkbz(付款币种) -> fnacurrency.CURRENCYNAME(币种名称)
    df['付款币种'] = df['付款币种ID'].map(lambda value: currency_map.get(c.format_code(value), ''))
    # [明细表] yskm(预算科目) -> fnabudgetfeetype 层级路径
    df['预算科目'] = df['预算科目ID'].map(lambda value: subject_map.get(c.format_code(value), ''))
    # [主表] xght(相关合同) -> uf_htsp.htbh(合同编号)
    df['相关合同'] = df['相关合同ID'].map(lambda value: _lookup_first_browser_value(contract_map, value))
    # [主表] yhkh(银行卡号) -> uf_khgys_dt1.yhzh(银行账号)
    df['银行卡号'] = df['银行卡号ID'].map(lambda value: _lookup_first_browser_value(bank_account_map, value))
    # [主表] fkdx(付款对象) -> uf_khgys.khmc(供应商名称),仅用于描述兜底和异常清单。
    df['付款对象'] = df['付款对象ID'].map(lambda value: _selected_supplier_name(value, supplier_status_map))
    df['项目编号'] = df['项目编号ID'].map(
        lambda value: _lookup_first_browser_value(project_code_map, value) or _text(value))
    df['付款性质'] = df['付款性质ID'].map(
        lambda value: PAYMENT_NATURE_MEANINGS.get(int(c.format_code(value)), '') if c.format_code(value).isdigit() else '')
    return df


def resolve_gig_source_values(source_df):
    """基于零工 SQL 返回的泛微 ID 字段补充输出需要的展示值。"""
    df = source_df.copy()
    employee_map = c.build_fw_employee_info_map_for_ids(df['经办人ID'])
    company_map = c.build_fw_company_name_map_for_ids(df['公司主体ID'])
    contract_map = c.build_fw_contract_code_map_for_ids(df['相关合同ID'])
    project_code_map = build_fw_project_code_map_for_ids(df['项目编号ID'])

    # [建模头表] jbr(经办人) -> hrmresource / hrmjobtitles
    df['经办人'] = df['经办人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('name', ''))
    df['经办人工号'] = df['经办人ID'].map(lambda value: employee_map.get(c.format_code(value), {}).get('code', ''))
    # [建模头表] gszt(公司主体) -> uf_gstt.gsmc(公司主体名称)
    df['公司主体'] = df['公司主体ID'].map(lambda value: company_map.get(c.format_code(value), ''))
    # 原流程表 htmc 为空时,用合同台账编号兜底。
    fallback_contract = df['相关合同ID'].map(lambda value: _lookup_first_browser_value(contract_map, value))
    df['合同名称'] = [
        name if _text(name) else fallback
        for name, fallback in zip(df['合同名称'], fallback_contract)
    ]
    df['项目编号'] = df['项目编号ID'].map(
        lambda value: _lookup_first_browser_value(project_code_map, value) or _text(value))
    return df


def resolve_gig_budget_values(budget_df):
    """基于零工原流程预算项明细补充预算科目路径。

    四合一 Excel 中「零工平台付款」的费用项目来源,对应原流程 formtable_main_279_dt3:
    - yskm(预算科目)
    - fkje(付给三方平台金额)
    """
    df = budget_df.copy()
    subject_map = c.build_fw_budget_subject_path_map_for_ids(df['预算科目ID'])
    df['预算科目'] = df['预算科目ID'].map(lambda value: subject_map.get(c.format_code(value), ''))
    return df


def merge_gig_budget_subjects(budget_df):
    """同一流程内先合并同类预算科目,保留预算科目首次出现顺序。"""
    if budget_df.empty:
        return budget_df.copy()

    df = budget_df.copy()
    df['_预算科目合并键'] = [
        c.format_code(subject_id) or c.remove_slashes(subject_path)
        for subject_id, subject_path in zip(df['预算科目ID'], df['预算科目'])
    ]

    merged_rows = []
    for _, group in df.groupby(['建模付款ID', '_预算科目合并键'], sort=False, dropna=False):
        row = group.iloc[0].copy()
        row['预算项金额'] = c.round_amount(
            pd.to_numeric(group['预算项金额'], errors='coerce').fillna(0).sum())
        if '预算项人民币金额' in group.columns:
            row['预算项人民币金额'] = c.round_amount(
                pd.to_numeric(group['预算项人民币金额'], errors='coerce').fillna(0).sum())
        row = row.drop(labels=['_预算科目合并键'], errors='ignore')
        merged_rows.append(row.to_dict())

    return pd.DataFrame(merged_rows)


def map_gig_budget_expense_items(budget_df):
    """把合并后的零工预算科目映射成新费用项目。"""
    if budget_df.empty:
        return budget_df.copy()

    df = budget_df.copy()
    subject_map = c.build_subject_map()
    mapped_items = df['预算科目'].map(
        lambda value: subject_map.get(c.remove_slashes(value), ('', '')))
    df['费用项目编码'] = mapped_items.map(lambda item: item[0])
    df['费用项目描述'] = mapped_items.map(lambda item: item[1])
    return df


def allocate_gig_budget_to_recipients(recipient_df, budget_df):
    """把原流程预算项分配到实际收款人明细。

    规则表「预付期初」R47 要求:从「对公&报销&零工&批量四合一」取预算科目并分配到
    「零工平台付款_实际收款人明细」。在 DB 中,四合一里「零工平台付款」对应
    formtable_main_279_dt3(预算项明细),实际收款人对应 formtable_main_279_dt4。

    分配口径:
    - 同一流程内先按 dt3 预算科目合并金额,保留该预算科目的首次出现顺序。
    - 合并后的预算科目再映射新费用项目,并以映射后的项目形成预算金额池。
    - 按 dt4 实际收款人明细顺序依次占用预算池; 收款人金额跨预算科目边界时拆成多行。
    - 预算项金额用完后才进入下一预算项; 不再按预算项比例拆分每个收款人。
    """
    rows = []
    raw_budget_count = len(budget_df)
    budget_df = map_gig_budget_expense_items(merge_gig_budget_subjects(budget_df))
    budget_by_payment = {
        key: group.copy()
        for key, group in budget_df.groupby('建模付款ID', dropna=False)
    }
    no_budget_count = 0
    multi_budget_payment_count = 0

    for payment_id, recipient_group in recipient_df.groupby('建模付款ID', sort=False, dropna=False):
        budget_group = budget_by_payment.get(payment_id)
        if budget_group is None or budget_group.empty:
            no_budget_count += len(recipient_group)
            for _, recipient_row in recipient_group.iterrows():
                row = recipient_row.to_dict()
                row.update({
                    '预算明细ID': '',
                    '预算科目ID': '',
                    '预算科目': '',
                    '费用项目编码': '',
                    '费用项目描述': '',
                    '预算项金额': '',
                    '预付款金额分摊': pd.to_numeric(recipient_row['付给三方平台金额'], errors='coerce'),
                })
                rows.append(row)
            continue

        budget_group = budget_group.copy()
        budget_group['_剩余预算项金额'] = (
            pd.to_numeric(budget_group['预算项金额'], errors='coerce')
            .fillna(0)
            .map(c.round_amount)
        )
        budget_total = float(budget_group['_剩余预算项金额'].sum())
        if len(budget_group) > 1:
            multi_budget_payment_count += 1

        if budget_total == 0:
            # 没有可用预算项金额时,保留收款人粒度并带第一条预算科目。
            first_budget = budget_group.iloc[0]
            for _, recipient_row in recipient_group.iterrows():
                row = recipient_row.to_dict()
                row.update({
                    '预算明细ID': first_budget.get('预算明细ID', ''),
                    '预算科目ID': first_budget.get('预算科目ID', ''),
                    '预算科目': first_budget.get('预算科目', ''),
                    '费用项目编码': first_budget.get('费用项目编码', ''),
                    '费用项目描述': first_budget.get('费用项目描述', ''),
                    '预算项金额': first_budget.get('预算项金额', ''),
                    '预付款金额分摊': pd.to_numeric(recipient_row['付给三方平台金额'], errors='coerce'),
                })
                rows.append(row)
            continue

        budget_records = [row.to_dict() for _, row in budget_group.iterrows()]
        budget_index = 0

        def append_allocated_row(recipient_row, budget_row, allocated_amount):
            row = recipient_row.to_dict()
            row.update({
                '预算明细ID': budget_row.get('预算明细ID', ''),
                '预算科目ID': budget_row.get('预算科目ID', ''),
                '预算科目': budget_row.get('预算科目', ''),
                '费用项目编码': budget_row.get('费用项目编码', ''),
                '费用项目描述': budget_row.get('费用项目描述', ''),
                '预算项金额': budget_row.get('预算项金额', ''),
                '预付款金额分摊': allocated_amount,
            })
            rows.append(row)

        for _, recipient_row in recipient_group.iterrows():
            recipient_amount = pd.to_numeric(recipient_row['付给三方平台金额'], errors='coerce')
            current_budget_index = min(budget_index, len(budget_records) - 1)
            if pd.isna(recipient_amount):
                append_allocated_row(recipient_row, budget_records[current_budget_index], recipient_amount)
                continue

            remaining_recipient_amount = c.round_amount(recipient_amount)
            if remaining_recipient_amount <= 0:
                append_allocated_row(recipient_row, budget_records[current_budget_index], remaining_recipient_amount)
                continue

            while remaining_recipient_amount > 0:
                if budget_index >= len(budget_records):
                    append_allocated_row(recipient_row, budget_records[-1], remaining_recipient_amount)
                    break

                budget_row = budget_records[budget_index]
                remaining_budget_amount = budget_row['_剩余预算项金额']
                if remaining_budget_amount <= 0:
                    budget_index += 1
                    continue

                allocated_amount = c.round_amount(min(remaining_recipient_amount, remaining_budget_amount))
                append_allocated_row(recipient_row, budget_row, allocated_amount)

                remaining_recipient_amount = c.round_amount(remaining_recipient_amount - allocated_amount)
                budget_row['_剩余预算项金额'] = c.round_amount(remaining_budget_amount - allocated_amount)
                if budget_row['_剩余预算项金额'] <= 0:
                    budget_index += 1

    allocated_df = pd.DataFrame(rows)
    print('[预付期初-零工预付款-DB] 未匹配预算项的收款人明细行数:', no_budget_count)
    print('[预付期初-零工预付款-DB] 预算项按预算科目合并:', raw_budget_count, '->', len(budget_df))
    print('[预付期初-零工预付款-DB] 多预算科目顺序分配的单据数:', multi_budget_payment_count)
    print('[预付期初-零工预付款-DB] 预算项分配后明细行数:', len(allocated_df))
    return allocated_df


# ============================ 模板输出:供应商预付款 ============================
def build_supplier_output(merged_df):
    """供应商预付 DB 源数据 -> 供应商预付款单导入模版 26 列。"""
    def lookup_by_name(mapping, value):
        return '' if pd.isna(value) else mapping.get(c.normalize_name(value), '')

    vendor_map = c.build_supplier_vendor_info_map_for_rows(
        merged_df['付款对象ID'],
        supplier_texts=merged_df['付款对象'],
        document_numbers=merged_df['流程编号'],
        missing_report_file=SUPPLIER_VENDOR_MISSING_FILE,
        log_prefix='[预付期初-供应商预付款-DB]',
    )
    entity_map = c.build_accounting_entity_map_for_names(merged_df['开票单位'])
    subject_map = c.build_subject_map()

    def lookup_vendor(index, field):
        return vendor_map.get(index, {}).get(field, '')

    def vendor_description(index, fallback_text):
        name = lookup_vendor(index, 'name')
        if name:
            return name
        return _text(fallback_text)

    def subject_item(subject_path, index):
        if pd.isna(subject_path):
            return ''
        return subject_map.get(c.remove_slashes(subject_path), ('', ''))[index]

    main_amount = pd.to_numeric(merged_df['付款金额'], errors='coerce').fillna(0)
    detail_amount = pd.to_numeric(merged_df['预付金额'], errors='coerce').fillna(0)
    remaining = pd.to_numeric(merged_df['剩余冲销/退款金额'], errors='coerce').fillna(0)
    ratio = (detail_amount / main_amount).where(main_amount != 0, 0)
    unsettled_amount = remaining * ratio
    settled_amount = detail_amount - unsettled_amount
    is_deposit = merged_df['付款性质ID'].map(c.format_code).isin(DEPOSIT_PAYMENT_NATURE_CODES)

    output_df = pd.DataFrame(index=merged_df.index)

    # 固定值。
    output_df['来源系统'] = 'FW'
    output_df['单据类型'] = DOCUMENT_TYPE

    # 泛微主表/明细表直取或解析字段。
    output_df['来源单据编号'] = merged_df['流程编号']                    # [主表] lcbh(流程编号)
    output_df['申请日期'] = merged_df['申请日期'].map(c.format_date)     # [主表] sqrq(申请日期)
    output_df['申请人工号'] = merged_df['填单人工号']                    # [主表] tdr(填单人) -> hrmjobtitles.JOBTITLENAME(工号)
    output_df['申请人姓名'] = merged_df['填单人']                        # [主表] tdr(填单人) -> hrmresource.LASTNAME(姓名)
    output_df['备注'] = merged_df['备注'].astype(str).where(
        merged_df['备注'].notna(), '').str.slice(0, 150)                # [主表] bz(备注),导入限制截前150字符
    output_df['合同号'] = merged_df['相关合同'].where(
        merged_df['相关合同'].notna(), '')                              # [主表] xght(相关合同) -> uf_htsp.htbh(合同编号)
    output_df['保证金标志'] = ['是' if flag else '否' for flag in is_deposit]  # [主表] fkxz(付款性质) 0=押金/1=质保金 -> 是
    output_df['银行账号'] = merged_df['银行卡号'].where(
        merged_df['银行卡号'].notna(), '')                              # [主表] yhkh(银行卡号) -> uf_khgys_dt1.yhzh(银行账号)

    # 项目&订单清洗结果:泛微项目编号 -> 订单编号/订单标题。
    output_df['泛微项目编号'] = merged_df['项目编号'].map(_text)
    output_df['订单编号'] = merged_df['项目编号'].map(lambda value: _order_mapping_value(value, '订单编号'))
    output_df['订单名称'] = merged_df['项目编号'].map(lambda value: _order_mapping_value(value, '订单标题'))

    # 当前源数据没有直接可用的计划行/付款计划日期/银行备注/主播房间号,按模板留空。
    output_df['合同收支计划行'] = ''
    output_df['计划付款日期'] = ''
    output_df['银行转账备注'] = ''
    output_df['主播房间号'] = ''

    # 跨系统映射字段。
    output_df['核算主体编号'] = merged_df['开票单位'].map(
        lambda value: lookup_by_name(entity_map, value))                # [主表] kpdw(开票单位) -> Hand 核算主体编号
    output_df['核算主体描述'] = merged_df['开票单位']                    # [主表] kpdw(开票单位) -> uf_gstt.gsmc(公司主体名称)
    output_df['收款方编码'] = [lookup_vendor(index, 'code') for index in merged_df.index]  # [主表] fkdx(付款对象) -> Hand vender_code(供应商编码)
    output_df['收款方描述'] = [
        vendor_description(index, supplier_text)
        for index, supplier_text in zip(merged_df.index, merged_df['付款对象'])
    ]                                                                    # 优先 Hand description(供应商描述),兜底 [主表] fkdx(付款对象) -> uf_khgys.khmc(供应商名称)
    output_df['费用项目编码'] = merged_df['预算科目'].map(
        lambda value: subject_item(value, 0))                            # [明细] yskm(预算科目) -> 规则表费用项目编码
    output_df['费用项目描述'] = merged_df['预算科目'].map(
        lambda value: subject_item(value, 1))                            # [明细] yskm(预算科目) -> 规则表费用项目描述
    output_df['预付款支付币种'] = merged_df['付款币种'].map(c.to_iso_currency)  # [主表] fkbz(付款币种) -> ISO币种

    # 金额拆分。
    output_df['预付款金额（支付币种）'] = detail_amount.map(c.round_amount)       # [明细] yfje(预付金额)
    output_df['已到票核销金额（支付币种）'] = settled_amount.map(c.round_amount)  # [明细] yfje(预付金额) - 按比例分摊的 [主表] sycxtkje(剩余冲销/退款金额)
    output_df['已付未核（支付币种）'] = unsettled_amount.map(c.round_amount)     # [主表] sycxtkje(剩余冲销/退款金额) 按明细占比分摊
    return output_df[SUPPLIER_OUTPUT_COLUMNS]


# ============================ 模板输出:零工预付款 ============================
def gig_platform_vendor(name):
    """收款方文本 -> 灵工平台收款方编码。"""
    text = _text(name)
    for keyword, code in GIG_PLATFORM_VENDOR.items():
        if keyword in text:
            return code
    return ''


def gig_recipient_remark(name, id_number, phone):
    """收款人备注:姓名-身份证-手机号 拼接,限 30 字。"""
    parts = [_text(value) for value in (name, id_number, phone) if _text(value)]
    return '-'.join(parts)[:30]


def build_gig_output(merged_df):
    """零工 DB 源数据 -> 灵工预付款单导入模版 29 列。"""
    vendor_map = c.build_vendor_map()
    entity_map = c.build_accounting_entity_map_for_names(merged_df['公司主体'])
    has_mapped_expense_items = {'费用项目编码', '费用项目描述'}.issubset(merged_df.columns)
    subject_map = {} if has_mapped_expense_items else c.build_subject_map()

    def gig_payee_code(value):
        payee = _text(value)
        if not payee:
            return ''
        return vendor_map.get(c.normalize_name(payee)) or payee

    def subject_item(subject_path, index):
        if pd.isna(subject_path):
            return ''
        return subject_map.get(c.remove_slashes(subject_path), ('', ''))[index]

    output_df = pd.DataFrame(index=merged_df.index)

    # 固定值。
    output_df['来源系统'] = 'FW'
    output_df['单据类型'] = GIG_DOCUMENT_TYPE
    output_df['保证金标志'] = '否'
    output_df['收款方类别'] = '供应商'
    output_df['预付款支付币种'] = 'CNY'
    output_df['传送状态'] = '传送成功'
    output_df['支付状态'] = '支付成功'
    output_df['退款状态'] = ''
    output_df['核销状态'] = '已核销'

    # 泛微头表/原流程明细字段。
    output_df['来源单据编号'] = merged_df['流程编号']                    # [建模头/原流程] lcbh(流程编号)
    output_df['申请日期'] = merged_df['申请日期'].map(c.format_date)     # [建模头] sqrq(申请日期)
    output_df['申请人工号'] = merged_df['经办人工号']                    # [建模头] jbr(经办人) -> hrmjobtitles.JOBTITLENAME(工号)
    output_df['申请人姓名'] = merged_df['经办人']                        # [建模头] jbr(经办人) -> hrmresource.LASTNAME(姓名)
    output_df['备注_单头'] = merged_df['备注'].astype(str).where(
        merged_df['备注'].notna(), '').str.slice(0, 150)                # [建模头] bz(备注)
    output_df['灵工平台收款方编码'] = merged_df['收款方文本'].map(gig_platform_vendor)  # [建模头] skfwb(收款方-文本)
    output_df['合同号'] = merged_df['合同名称'].where(
        merged_df['合同名称'].notna(), '')                              # [原流程] htmc(合同名称);为空时合同编号兜底
    output_df['计划付款日期'] = merged_df['预计付款日期'].map(c.format_date)  # [建模头] yjfkrq(预计付款日期)
    output_df['收款方编码'] = merged_df['实际收款方'].map(gig_payee_code)  # [原流程明细] dt4.skf(收款方)
    output_df['备注'] = [
        gig_recipient_remark(name, id_number, phone)
        for name, id_number, phone in zip(merged_df['实际收款方'], merged_df['身份证号'], merged_df['手机号'])
    ]                                                                    # 姓名-身份证-手机号
    output_df['银行账号'] = merged_df['银行账号'].where(
        merged_df['银行账号'].notna(), '')                              # [原流程明细] dt4.yhzh(银行账号)
    output_df['预付款金额（支付币种）'] = pd.to_numeric(
        merged_df['预付款金额分摊'], errors='coerce').map(c.round_amount)  # [原流程明细] dt4.sl(付给三方平台金额) 按合并预算科目顺序占用预算项金额

    # 项目&订单清洗结果:泛微项目编号 -> 订单编号/订单标题。
    output_df['泛微项目编号'] = merged_df['项目编号'].map(_text)
    output_df['订单编号'] = merged_df['项目编号'].map(lambda value: _order_mapping_value(value, '订单编号'))
    output_df['订单名称'] = merged_df['项目编号'].map(lambda value: _order_mapping_value(value, '订单标题'))
    output_df['合同收支计划行'] = ''
    output_df['银行转账备注'] = ''
    if has_mapped_expense_items:
        output_df['费用项目编码'] = merged_df['费用项目编码'].where(
            merged_df['费用项目编码'].notna(), '')                        # [合并预算科目] 已映射的新费用项目编码
        output_df['费用项目描述'] = merged_df['费用项目描述'].where(
            merged_df['费用项目描述'].notna(), '')                        # [合并预算科目] 已映射的新费用项目描述
    else:
        output_df['费用项目编码'] = merged_df['预算科目'].map(
            lambda value: subject_item(value, 0))                        # [原流程预算项明细] dt3.yskm(预算科目) -> 规则表费用项目编码
        output_df['费用项目描述'] = merged_df['预算科目'].map(
            lambda value: subject_item(value, 1))                        # [原流程预算项明细] dt3.yskm(预算科目) -> 规则表费用项目描述

    # 跨系统映射字段。
    output_df['核算主体编号'] = merged_df['公司主体'].map(
        lambda value: entity_map.get(c.normalize_name(value), '') if _text(value) else '')  # [建模头] gszt(公司主体) -> Hand 核算主体编号
    output_df['核算主体描述'] = merged_df['公司主体']                    # [建模头] gszt(公司主体) -> uf_gstt.gsmc(公司主体名称)
    return output_df[GIG_OUTPUT_COLUMNS]


# ============================ 源读取 ============================
def read_supplier_source():
    """从 DB 直接读取过滤后的供应商预付主子合并数据。"""
    c.validate_fw_fields(FW_SUPPLIER_TABLE, EXPECTED_SUPPLIER_FIELDS)
    stats = _query_fw(SUPPLIER_STATS_SQL).iloc[0]
    status_name = FLOW_STATUS_MEANINGS.get(APPROVED_STATUS_CODE, '')
    void_name = VOID_FLAG_MEANINGS.get(VOID_CODE, '')
    print(f"[预付期初-供应商预付款-DB] SQL过滤: 申请日期>={DATE_FROM} "
          f"且 流程状态={APPROVED_STATUS_CODE}({status_name}) 且 是否作废≠{VOID_CODE}({void_name})")
    print(f"  满足前两项 {int(stats['matched_count'] or 0)} 单; "
          f"其中剔除作废 {int(stats['void_count'] or 0)} 单; "
          f"最终保留主表 {int(stats['kept_count'] or 0)} 单")

    merged_df = resolve_supplier_source_values(_query_fw(SUPPLIER_SOURCE_SQL))
    print('[预付期初-供应商预付款-DB] SQL主子合并明细行数:', len(merged_df))
    return merged_df


def read_gig_source():
    """从 DB 直接读取过滤后的零工付款头 + 原流程收款人明细 + 原流程预算项明细。"""
    c.validate_fw_fields(FW_GIG_HEADER_TABLE, EXPECTED_GIG_HEADER_FIELDS)
    c.validate_fw_fields(FW_GIG_WORKFLOW_TABLE, EXPECTED_GIG_WORKFLOW_FIELDS)
    stats = _query_fw(GIG_STATS_SQL).iloc[0]
    status_name = FLOW_STATUS_MEANINGS.get(APPROVED_STATUS_CODE, '')
    void_name = VOID_FLAG_MEANINGS.get(VOID_CODE, '')
    print(f"[预付期初-零工预付款-DB] SQL过滤: 申请日期>={DATE_FROM} "
          f"且 流程状态={APPROVED_STATUS_CODE}({status_name}) 且 是否作废≠{VOID_CODE}({void_name})")
    print(f"  满足前两项 {int(stats['matched_count'] or 0)} 单; "
          f"其中剔除作废 {int(stats['void_count'] or 0)} 单; "
          f"最终保留主表 {int(stats['kept_count'] or 0)} 单")

    recipient_df = resolve_gig_source_values(_query_fw(GIG_RECIPIENT_SOURCE_SQL))
    print('[预付期初-零工预付款-DB] SQL收款人明细行数:', len(recipient_df))
    budget_df = resolve_gig_budget_values(_query_fw(GIG_BUDGET_SOURCE_SQL))
    print('[预付期初-零工预付款-DB] SQL预算项明细行数:', len(budget_df))
    return allocate_gig_budget_to_recipients(recipient_df, budget_df)


def run():
    # 1. SQL 直接查过滤后的源数据
    supplier_merged_df = read_supplier_source()
    gig_merged_df = read_gig_source()

    # 2. 构建两个 tab 输出
    supplier_output_df = build_supplier_output(supplier_merged_df)
    print('[预付期初-供应商预付款-DB] 输出明细行数:', len(supplier_output_df))
    gig_output_df = build_gig_output(gig_merged_df)
    print('[预付期初-零工预付款-DB] 输出明细行数:', len(gig_output_df))

    # 3. 填充率(必输字段以规则表「是否必填」=Y 为准)
    supplier_required = c.required_columns(RULE_SHEET, RULE_TABLE_SUPPLIER)
    gig_required = c.required_columns(RULE_SHEET, RULE_TABLE_GIG)
    print('— 供应商预付款 填充率 —')
    c.report_fill(supplier_output_df, supplier_required)
    print('— 灵工预付款 填充率 —')
    c.report_fill(gig_output_df, gig_required)

    # 4. 写模版(两个 tab 一次写入,lov 页保留)
    c.write_template_sheets(TEMPLATE_FILE, OUTPUT_FILE, {
        TEMPLATE_SHEET_SUPPLIER: supplier_output_df,
        TEMPLATE_SHEET_GIG: gig_output_df,
    })
    print('已写出:', OUTPUT_FILE)

    # 5. 问题清单
    exception_sheets = {}
    supplier_sheets = {'必输字段未达100%': c.fill_summary(
        supplier_output_df, supplier_required, RULE_SHEET, RULE_TABLE_SUPPLIER)}
    supplier_sheets.update(c.collect_field_issues(
        supplier_output_df, supplier_merged_df, supplier_required, SUPPLIER_ISSUE_SOURCE_FIELDS))
    supplier_sheets.update(collect_order_mapping_issues(supplier_merged_df))
    exception_sheets.update({f'供应商_{name}': df for name, df in supplier_sheets.items()})

    gig_sheets = {'必输字段未达100%': c.fill_summary(
        gig_output_df, gig_required, RULE_SHEET, RULE_TABLE_GIG)}
    gig_sheets.update(c.collect_field_issues(
        gig_output_df, gig_merged_df, gig_required, GIG_ISSUE_SOURCE_FIELDS))
    gig_sheets.update(collect_order_mapping_issues(gig_merged_df))
    exception_sheets.update({f'灵工_{name}': df for name, df in gig_sheets.items()})

    c.write_exceptions(EXCEPTION_FILE, exception_sheets)
    print('已写出:', EXCEPTION_FILE, '| 各清单条数:', {k: len(v) for k, v in exception_sheets.items()})


if __name__ == '__main__':
    run()
