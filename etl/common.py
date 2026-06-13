# -*- coding: utf-8 -*-
"""公共能力:路径配置、数据库只读连接、各类映射、归一化、Excel 读写、行过滤/去重。

所有清洗任务(ap_opening_payment 应付期初,以及将来的应收/预付/预收)都从这里取用,
避免在每个任务里重复写数据库查询和映射逻辑。任务文件本身只关心"过滤 + 字段映射"。

数据库账密从环境变量读取;若项目根有 .env 则自动加载(.env 不提交版本库)。
需要的变量:FW_*(泛微 vspn_xtyy 取工号)、ZT_*(中台库取供应商编码和核算主体编码)。
"""
import os
import re
from datetime import date
from pathlib import Path

import pandas as pd
import pymysql
from openpyxl import load_workbook

# ============================ 路径 ============================
ROOT      = Path(__file__).resolve().parents[1]
SRC_DIR   = ROOT / 'data' / 'source'       # 源表(泛微导出)
RULES_DIR = ROOT / 'data' / 'rules'        # 映射规则
TPL_DIR   = ROOT / 'data' / 'templates'    # 导入模版
OUT_DIR   = ROOT / 'output'                # 产出
RULE_XLSX = RULES_DIR / '业财项目_数据映射规则.xlsx'


# ============================ 配置 / 数据库 ============================
def _load_env():
    env = ROOT / '.env'
    if env.exists():
        for line in env.read_text(encoding='utf-8').splitlines():
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                k, v = line.split('=', 1)
                os.environ.setdefault(k.strip(), v.strip())


def _db_connect(prefix, database):
    """按前缀(FW/ZT)建只读连接。只跑 SELECT,不写生产库。"""
    try:
        return pymysql.connect(
            host=os.environ[f'{prefix}_HOST'], port=int(os.environ[f'{prefix}_PORT']),
            user=os.environ[f'{prefix}_USER'], password=os.environ[f'{prefix}_PASS'],
            database=database, charset='utf8mb4', connect_timeout=20)
    except KeyError as e:
        raise RuntimeError(f'缺少数据库环境变量 {e};请在 .env 或环境变量配置 {prefix}_HOST/PORT/USER/PASS') from e


_load_env()
OUT_DIR.mkdir(parents=True, exist_ok=True)


def today_suffix():
    """返回产出文件使用的运行日期后缀。"""
    return date.today().strftime('%Y%m%d')


# ============================ 归一化 / 格式化 ============================
_PARENTHESIS_TRANSLATION = str.maketrans('（）', '()')


def normalize_name(value):
    """名称归一化:全角括号->半角、去所有空格。消除主体/供应商/人名的全半角差异。"""
    return re.sub(r'\s+', '', str(value).strip().translate(_PARENTHESIS_TRANSLATION))


def remove_slashes(value):
    """去掉所有 '/'。费用科目里 '/' 既是层级分隔符又可能是层级名内部字符,两侧统一去掉才能对齐。"""
    return str(value).strip().replace('/', '')


def format_code(value):
    """浮点编码(1000.0)->整数串'1000';空值->''。"""
    text = str(value).strip()
    if text in ('', 'nan', 'None'):
        return ''
    try:
        return str(int(float(text)))
    except ValueError:
        return text


def format_date(value):
    """日期 -> 'yyyy-mm-dd';空值->''。"""
    return '' if pd.isna(value) else pd.to_datetime(value).strftime('%Y-%m-%d')


def round_amount(value):
    """金额保留 2 位小数;空值->''。"""
    return '' if pd.isna(value) else round(float(value), 2)


# ============================ 映射字典 ============================
# 币种 -> ISO 码
CURRENCY_TO_ISO = {
    '人民币': 'CNY', '美元': 'USD', '马来西亚令吉': 'MYR', '泰铢': 'THB', '印尼盾': 'IDR',
    '韩元': 'KRW', '港币': 'HKD', '新加坡元': 'SGD', '沙特里亚尔': 'SAR', '菲律宾比索': 'PHP',
    '欧元': 'EUR', '日元': 'JPY', '英镑': 'GBP', '瑞士法郎': 'CHF',
    '伊拉克第纳尔': 'IQD', '科威特第纳尔': 'KWD', '埃及镑': 'EGP',
}


def to_iso_currency(currency_name):
    text = '' if pd.isna(currency_name) else str(currency_name).strip()
    return CURRENCY_TO_ISO.get(text, text)


def build_employee_code_map():
    """经办人姓名 -> 工号。来源:泛微 vspn_xtyy。
    hrmresource.JOBTITLE 关联 hrmjobtitles.id 后取 hrmjobtitles.JOBTITLENAME;键用 normalize_name 归一化。"""
    conn = _db_connect('FW', 'vspn_xtyy')
    try:
        employee_df = pd.read_sql(
            'SELECT r.LASTNAME employee_name, j.JOBTITLENAME employee_code '
            'FROM hrmresource r LEFT JOIN hrmjobtitles j ON r.JOBTITLE=j.id',
            conn)
    finally:
        conn.close()
    employee_df['key'] = employee_df['employee_name'].map(normalize_name)
    employee_code_map = {}
    for key, employee_codes in employee_df.groupby('key')['employee_code']:
        valid_codes = [str(code).strip() for code in employee_codes.dropna().unique()
                       if str(code).strip() not in ('', 'Default', 'nan')]
        if key and valid_codes:
            employee_code_map[key] = valid_codes[0]
    return employee_code_map


def build_vendor_map():
    """供应商名称 -> 中台供应商编码 vender_code。来源:中台 hfins_base.hfbs_system_vender。
    按 description / taxpayer_name 建键(均 normalize_name 归一化)。"""
    conn = _db_connect('ZT', 'hfins_base')
    try:
        vendor_df = pd.read_sql(
            'SELECT vender_code vendor_code, description vendor_name, taxpayer_name taxpayer_name '
            'FROM hfbs_system_vender',
            conn)
    finally:
        conn.close()
    vendor_map = {}
    for _, row in vendor_df.iterrows():
        for name in (row['vendor_name'], row['taxpayer_name']):
            key = normalize_name(name)
            if key and key not in ('nan', 'None') and key not in vendor_map:
                vendor_map[key] = str(row['vendor_code']).strip()
    return vendor_map


def build_accounting_entity_map():
    """公司主体名称 -> 核算主体编号。来源:中台 hfins_base_account.hfac_accounting_entity。
    按 acc_entity_name 建键。"""
    conn = _db_connect('ZT', 'hfins_base_account')
    try:
        entity_df = pd.read_sql(
            'SELECT acc_entity_code, acc_entity_name '
            'FROM hfac_accounting_entity',
            conn)
    finally:
        conn.close()
    entity_map = {}
    for _, row in entity_df.iterrows():
        code = str(row['acc_entity_code']).strip()
        if not code:
            continue
        key = normalize_name(row.get('acc_entity_name', ''))
        if key and key != 'nan' and key not in entity_map:
            entity_map[key] = code
    return entity_map


def build_subject_map():
    """费用科目(预算科目)-> (费用项目编码, 费用项目描述)。
    来源:规则表「赛事/MCN 新旧预算项科目-调整后」。键=各级科目名拼接去'/';值=调整后三级预算编码+三级费用项。"""
    subject_map = {}
    event_budget_df = pd.read_excel(RULE_XLSX, sheet_name='赛事新旧预算项科目-调整后', header=None)
    for _, row in event_budget_df.iloc[2:].iterrows():
        key = remove_slashes(row[15])
        subject_code = str(row[21]).strip()
        subject_name = str(row[22]).strip()
        if key and key != 'nan' and subject_code != 'nan':
            subject_map[key] = (subject_code, subject_name)
    mcn_budget_df = pd.read_excel(RULE_XLSX, sheet_name='MCN新旧预算项科目-调整后', header=None)
    for _, row in mcn_budget_df.iloc[2:].iterrows():
        key = remove_slashes(''.join(str(row[column]) for column in (1, 3, 5) if str(row[column]) != 'nan'))
        subject_code = str(row[11]).strip() if str(row[11]) != 'nan' else str(row[7]).strip()
        subject_name = str(row[8]).strip()
        if key and subject_code != 'nan':
            subject_map.setdefault(key, (subject_code, subject_name))
    return subject_map


# ============================ 行过滤 / 去重 / 统计 ============================
def filter_main(main_df, sources, date_from='2026-01-01', status='审批完成',
                date_col='申请日期', drop_void=True):
    """主表标准行过滤,返回过滤后副本。"""
    filtered_main_df = main_df.copy()
    filtered_main_df[date_col] = pd.to_datetime(filtered_main_df[date_col], errors='coerce')
    keep_mask = (
        filtered_main_df['流程来源'].isin(sources)
        & (filtered_main_df[date_col] >= date_from)
        & (filtered_main_df['流程状态'] == status)
    )
    total_count, void_count = int(keep_mask.sum()), 0
    if drop_void and '是否作废' in filtered_main_df.columns:
        void_mask = filtered_main_df['是否作废'].astype(str).str.strip() == '是'
        void_count = int((keep_mask & void_mask).sum())
        keep_mask = keep_mask & ~void_mask
    result_df = filtered_main_df[keep_mask].copy()
    print(f'过滤: 范围+日期+状态={total_count}单; 作废剔除={void_count}单; 最终主表={len(result_df)}单')
    return result_df


def dedup_rows(output_df, key_cols):
    """按 key_cols 完全相同则合并为一条。返回(去重后, 参与合并的全部行[供核对])。"""
    group_key = output_df[key_cols].apply(lambda row: '|'.join(map(str, row.tolist())), axis=1)
    duplicate_mask = group_key.duplicated(keep=False)
    merged_rows = output_df[duplicate_mask].assign(_key=group_key).sort_values('_key').drop(columns='_key')
    deduped_rows = output_df[~group_key.duplicated(keep='first')].reset_index(drop=True)
    print(f'分组去重: 参与合并 {int(duplicate_mask.sum())} 行 -> 最终输出 {len(deduped_rows)} 行')
    return deduped_rows, merged_rows


def report_fill(output_df, columns):
    for column in columns:
        filled_count = (output_df[column].astype(str).str.strip() != '').sum()
        print(f'  {column} 填充率: {filled_count}/{len(output_df)} = {filled_count/len(output_df)*100:.1f}%')


# ============================ Excel 输出 ============================
def write_to_template(output_df, template_path, output_path, sheet_name):
    """写进导入模版(保留表头与 lov 下拉页),从第 2 行覆盖写入。"""
    wb = load_workbook(template_path)
    ws = wb[sheet_name]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)
    for _, row in output_df.iterrows():
        ws.append(['' if pd.isna(v) else v for v in row.tolist()])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def write_exceptions(output_path, sheets):
    """导出未匹配/待核对清单。sheets: {sheet名: DataFrame}。"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path) as writer:
        for sheet_name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output_path
