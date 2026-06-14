# -*- coding: utf-8 -*-
"""公共能力:路径配置、数据库只读连接、各类映射、归一化、Excel 读写、行过滤/去重。

所有清洗任务(ap_opening_payment 应付期初,以及将来的应收/预付/预收)都从这里取用,
避免在每个任务里重复写数据库查询和映射逻辑。任务文件本身只关心"过滤 + 字段映射"。

数据库账密从环境变量读取;若项目根有 .env 则自动加载(.env 不提交版本库)。
需要的变量:FW_*(泛微 vspn_xtyy 取工号)、ZT_*(中台库取供应商编码和核算主体编码)。
"""
import os
import re
import unicodedata
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
# NFKD 折不掉的特殊拉丁字母(无组合记号),显式映射到基础字母。覆盖土耳其语等外文供应商/人名。
_SPECIAL_LETTERS = str.maketrans({
    'ı': 'i', 'İ': 'I', 'ø': 'o', 'Ø': 'O', 'ł': 'l', 'Ł': 'L',
    'đ': 'd', 'Đ': 'D', 'ð': 'd', 'Ð': 'D', 'ħ': 'h', 'ŧ': 't',
    'ß': 'ss', 'æ': 'ae', 'Æ': 'AE', 'œ': 'oe', 'Œ': 'OE',
})


def _fold_accents(text):
    """去掉变音符号:ş->s、ç->c、ö->o、ı->i…,把外文名折成基础拉丁字母(并把全角等兼容字符规整)。"""
    text = text.translate(_SPECIAL_LETTERS)
    decomposed = unicodedata.normalize('NFKD', text)
    return ''.join(ch for ch in decomposed if not unicodedata.combining(ch))


def normalize_name(value):
    """名称归一化(用于按名称匹配):折叠变音符号(土耳其语 ş/ı 等)、去掉所有空格/标点/符号
    (含 & 、,。()（）等)、忽略大小写;只保留字母数字与中文。
    消除主体/供应商/人名在两套系统间的全半角、特殊字符、标点、大小写差异。"""
    text = _fold_accents(str(value).strip())
    # 去掉空白、标点(category P*)、符号(category S*,含 &);保留字母/数字/中文
    text = ''.join(ch for ch in text
                   if not (ch.isspace() or unicodedata.category(ch)[0] in ('P', 'S')))
    return text.casefold()


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
    hrmresource.JOBTITLE 关联 hrmjobtitles.id 后取 hrmjobtitles.JOBTITLENAME;键用 normalize_name 归一化。
    取到占位值 'Default' 也算匹配、不留空(同名有真实工号时优先真实工号)-- 20260614Leo确认。"""
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
        values = [str(code).strip() for code in employee_codes.dropna().unique()
                  if str(code).strip() not in ('', 'nan')]
        real_codes = [code for code in values if code != 'Default']
        # 取到 Default 也算匹配、不留空;同名有真实工号时优先真实工号 -- Leo确认
        chosen = real_codes[0] if real_codes else (values[0] if values else None)
        if key and chosen:
            employee_code_map[key] = chosen
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


def build_customer_map():
    """客户名称 -> 中台客户编码 customer_code。来源:中台 hfins_base.hfbs_system_customer。
    按 description / taxpayer_name 建键(均 normalize_name 归一化)。"""
    conn = _db_connect('ZT', 'hfins_base')
    try:
        customer_df = pd.read_sql(
            'SELECT customer_code, description customer_name, taxpayer_name '
            'FROM hfbs_system_customer',
            conn)
    finally:
        conn.close()
    customer_map = {}
    for _, row in customer_df.iterrows():
        for name in (row['customer_name'], row['taxpayer_name']):
            key = normalize_name(name)
            if key and key not in ('nan', 'none') and key not in customer_map:
                customer_map[key] = str(row['customer_code']).strip()
    return customer_map


def build_lov_meaning_map(lov_code):
    """HZero 值集 meaning -> value。用于把业务侧展示值转成汉得编码。"""
    conn = _db_connect('ZT', 'hzero_platform')
    try:
        lov_df = pd.read_sql(
            'SELECT v.value, v.meaning '
            'FROM hpfm_lov l JOIN hpfm_lov_value v ON v.lov_id = l.lov_id '
            'WHERE l.lov_code = %s AND v.enabled_flag = 1 '
            'ORDER BY v.order_seq, v.value',
            conn,
            params=[lov_code])
    finally:
        conn.close()
    return {
        str(row['meaning']).strip(): str(row['value']).strip()
        for _, row in lov_df.iterrows()
        if str(row['meaning']).strip() and str(row['meaning']).strip() != 'nan'
    }


def build_tax_type_description_map(preferred_descriptions=None):
    """税率 -> 汉得税率类型描述。preferred_descriptions 可指定每个税率优先使用的 description。
    键统一为小数税率,例如 6% 为 0.06。"""
    conn = _db_connect('ZT', 'hfins_base')
    try:
        tax_df = pd.read_sql(
            'SELECT tax_type_code, description, tax_type_rate, sale_tax_flag, input_tax_flag, enabled_flag '
            'FROM hfbs_tax_type '
            'WHERE enabled_flag = 1',
            conn)
    finally:
        conn.close()

    tax_df['rate_key'] = pd.to_numeric(tax_df['tax_type_rate'], errors='coerce').round(4)
    description_to_rate = {
        str(row['description']).strip(): float(row['rate_key'])
        for _, row in tax_df.iterrows()
        if pd.notna(row['rate_key']) and str(row['description']).strip()
    }

    tax_map = {}
    for rate, descriptions in (preferred_descriptions or {}).items():
        for description in descriptions:
            if description in description_to_rate:
                tax_map[round(float(rate), 4)] = description
                break

    # 未显式指定的税率,优先取销项税、再取非进项税,最后保底取该税率第一条。
    sort_df = tax_df.assign(
        sale_priority=(tax_df['sale_tax_flag'] == 1).astype(int),
        non_input_priority=(tax_df['input_tax_flag'] == 0).astype(int),
    ).sort_values(['rate_key', 'sale_priority', 'non_input_priority', 'tax_type_code'],
                  ascending=[True, False, False, True])
    for _, row in sort_df.iterrows():
        if pd.isna(row['rate_key']):
            continue
        tax_map.setdefault(round(float(row['rate_key']), 4), str(row['description']).strip())
    return tax_map


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


def build_fw_company_map():
    """泛微公司主体ID -> 公司主体名称。来源:泛微 vspn_xtyy.uf_gstt。"""
    conn = _db_connect('FW', 'vspn_xtyy')
    try:
        company_df = pd.read_sql(
            'SELECT id company_id, gsmc company_name '
            'FROM uf_gstt',
            conn)
    finally:
        conn.close()
    company_map = {}
    for _, row in company_df.iterrows():
        company_id = format_code(row['company_id'])
        company_name = str(row['company_name']).strip()
        if company_id and company_name and company_name != 'nan':
            company_map[company_id] = company_name
    return company_map


def _cell_text(value):
    """Excel 单元格文本归一化: 空值/nan/None -> ''。"""
    if pd.isna(value):
        return ''
    text = str(value).strip()
    return '' if text in ('', 'nan', 'None', 'NaT') else text


def _joined_row_key(row, columns):
    return remove_slashes(''.join(_cell_text(row[column]) for column in columns if column < len(row)))


def _first_subject_pair(row, pairs):
    """按优先级读取 (费用项目编码, 费用项目描述) 列对。"""
    for code_column, name_column in pairs:
        code = _cell_text(row[code_column]) if code_column < len(row) else ''
        name = _cell_text(row[name_column]) if name_column < len(row) else ''
        if code and code != '\\':
            return code, name
    return '', ''


def _read_rule_sheet(sheet_name):
    return pd.read_excel(RULE_XLSX, sheet_name=sheet_name, header=None, engine='calamine')


def build_subject_map():
    """费用科目(预算科目)-> (费用项目编码, 费用项目描述)。
    来源:规则表「赛事/MCN 新旧预算项科目-调整后」。键=各级科目名拼接去'/';值=调整后三级预算编码+三级费用项。"""
    subject_map = {}
    event_base_df = _read_rule_sheet('赛事泛微新旧科目映射底表')
    event_base_map = {}
    for _, row in event_base_df.iloc[1:].iterrows():
        key = remove_slashes(_cell_text(row[5]))
        code, name = _first_subject_pair(row, ((6, 7),))
        if key and code:
            event_base_map[key] = (code, name)

    event_budget_df = _read_rule_sheet('赛事新旧预算项科目-调整后')
    for _, row in event_budget_df.iloc[2:].iterrows():
        keys = {_joined_row_key(row, (2, 5, 8, 11, 14)), remove_slashes(_cell_text(row[15]))}
        for key in (k for k in keys if k):
            subject_code, subject_name = _first_subject_pair(row, ((21, 22), (19, 20), (17, 18)))
            if not subject_code:
                subject_code, subject_name = event_base_map.get(key, ('', ''))
            if subject_code:
                subject_map[key] = (subject_code, subject_name)

    mcn_base_df = _read_rule_sheet('MCN泛微新旧科目映射底表')
    mcn_code_to_name = {}
    mcn_code_to_subject = {}
    mcn_name_to_subject = {}
    for _, row in mcn_base_df.iloc[2:].iterrows():
        old_code = _cell_text(row[1])
        old_name = _cell_text(row[2])
        subject_code, subject_name = _first_subject_pair(row, ((5, 6),))
        if old_code and old_name:
            mcn_code_to_name[old_code] = old_name
        if old_code and subject_code:
            mcn_code_to_subject[old_code] = (subject_code, subject_name)
        if old_name and subject_code:
            mcn_name_to_subject[old_name] = (subject_code, subject_name)

    mcn_budget_df = _read_rule_sheet('MCN新旧预算项科目-调整后')
    for _, row in mcn_budget_df.iloc[2:].iterrows():
        codes = [_cell_text(row[column]) for column in (0, 2, 4) if column < len(row)]
        names = [_cell_text(row[column]) or mcn_code_to_name.get(code, '') for column, code in zip((1, 3, 5), codes)]
        keys = {
            _joined_row_key(row, (1, 3, 5)),
            remove_slashes(''.join(name for name in names if name)),
            remove_slashes(''.join(code for code in codes if code)),
        }
        subject_code, subject_name = _first_subject_pair(row, ((11, 12), (9, 10), (7, 8)))
        if not subject_code:
            deepest_code = next((code for code in reversed(codes) if code), '')
            deepest_name = next((name for name in reversed(names) if name), '')
            subject_code, subject_name = (
                mcn_code_to_subject.get(deepest_code)
                or mcn_name_to_subject.get(deepest_name)
                or ('', '')
            )
        if subject_code:
            for key in (k for k in keys if k):
                subject_map.setdefault(key, (subject_code, subject_name))
    return subject_map


# ============================ 去重 / 统计 ============================
# 行过滤口径(流程来源/日期/状态等)各任务差异较大,放在各任务文件内的 filter_main,不在此公共层。
def dedup_rows(output_df, key_cols):
    """按 key_cols 完全相同则合并为一条。返回(去重后, 参与合并的全部行[供核对])。"""
    group_key = output_df[key_cols].apply(lambda row: '|'.join(map(str, row.tolist())), axis=1)
    duplicate_mask = group_key.duplicated(keep=False)
    # 注意:key 要对齐到 mask 子集;否则 0 重复时空表 .assign 整列会把空表撑回全部行
    merged_rows = output_df[duplicate_mask].copy()
    merged_rows['_key'] = group_key[duplicate_mask]
    merged_rows = merged_rows.sort_values('_key').drop(columns='_key')
    deduped_rows = output_df[~group_key.duplicated(keep='first')].reset_index(drop=True)
    print(f'分组去重: 参与合并 {int(duplicate_mask.sum())} 行 -> 最终输出 {len(deduped_rows)} 行')
    return deduped_rows, merged_rows


def required_columns(rule_sheet, table_name):
    """必输字段以【规则表的「是否必填」列=Y】。
    从规则表 rule_sheet 内、表名列=table_name 的行里,取「是否必填」=Y 的字段名。
    规则列:0=模块 1=表名 2=字段名 3=是否必填。"""
    df = pd.read_excel(RULE_XLSX, sheet_name=rule_sheet, header=None, engine='calamine')
    columns = []
    current_table = ''
    for _, row in df.iloc[2:].iterrows():
        table = str(row[1]).strip()
        if table and table != 'nan':
            current_table = table
        if current_table == table_name and str(row[3]).strip() == 'Y':
            field = str(row[2]).strip()
            if field and field != 'nan' and field not in columns:
                columns.append(field)
    return columns


def required_column_remarks(rule_sheet, table_name):
    """读取规则表目标表下各字段的备注。用于问题清单汇总页补充说明。"""
    df = pd.read_excel(RULE_XLSX, sheet_name=rule_sheet, header=None, engine='calamine')
    remarks = {}
    current_table = ''
    for _, row in df.iloc[2:].iterrows():
        table = str(row[1]).strip()
        if table and table != 'nan':
            current_table = table
        if current_table != table_name:
            continue
        field = str(row[2]).strip()
        if not field or field == 'nan' or field in remarks:
            continue
        remarks[field] = _cell_text(row[5] if len(row) > 5 else '')
    return remarks


def report_fill(output_df, columns):
    """打印各列非空填充率。只统计 output_df 里存在的列。"""
    for column in columns:
        if column not in output_df.columns:
            continue
        filled_count = (output_df[column].astype(str).str.strip() != '').sum()
        print(f'  {column} 填充率: {filled_count}/{len(output_df)} = {filled_count/len(output_df)*100:.1f}%')


def collect_field_issues(output_df, source_df, required_cols, source_field_map, doc_col='来源单据编号'):
    """驱动于必输字段:遍历所有必输字段,凡【部分缺失】(0<缺失<全部)的,各生成一张
    「缺失_<字段>」明细 sheet。每张只两列:来源单据编号 + 泛微原表-<源字段>(没匹配上的原始值)。
    不写死具体字段;以后新增必输字段自动纳入。
    全空字段(如规则说明无需填写/暂未映射的字段)只在「必输字段未达100%」汇总里体现,不导出整表明细。
        output_df         输出宽表(含 doc_col;用于判断缺失)
        source_df         与 output_df 同索引的主子合并表(取泛微原始字段值)
        required_cols     必输字段(通常来自 required_columns(模版))
        source_field_map  {输出必输字段: source_df 里对应的泛微源字段名}
        doc_col           单据编号列(默认输出里的「来源单据编号」)"""
    sheets = {}
    total = len(output_df)
    for column in required_cols:
        if column not in output_df.columns:
            continue
        blank_mask = output_df[column].astype(str).str.strip() == ''
        missing_count = int(blank_mask.sum())
        if not (0 < missing_count < total):
            continue
        data = {doc_col: output_df.loc[blank_mask, doc_col].astype(str)}
        source_field = source_field_map.get(column)
        if source_field and source_field in source_df.columns:
            data[f'泛微原表-{source_field}'] = source_df.loc[blank_mask, source_field].astype(str)
        sheets[f'缺失_{column}'] = pd.DataFrame(data).drop_duplicates().reset_index(drop=True)
    return sheets


def fill_summary(output_df, columns, rule_sheet=None, table_name=None):
    """返回必输字段中【填充率未达100%】的汇总(供问题清单)。全部满 100% 时返回空表。
    若规则备注含「无需填写」且该必输字段整列为空,汇总页备注写「无需填写」。"""
    total = len(output_df)
    rule_remarks = required_column_remarks(rule_sheet, table_name) if rule_sheet and table_name else {}
    rows = []
    for column in columns:
        column_exists = column in output_df.columns
        filled = int((output_df[column].astype(str).str.strip() != '').sum()) if column_exists else 0
        if filled < total:
            if filled == 0 and '无需填写' in rule_remarks.get(column, ''):
                remark = '无需填写'
            elif not column_exists:
                remark = '输出表缺少该列'
            else:
                remark = ''
            rows.append({'必输字段': column, '填充数': filled, '缺失数': total - filled,
                         '总数': total, '填充率': f'{filled / total * 100:.1f}%', '备注': remark})
    return pd.DataFrame(rows, columns=['必输字段', '填充数', '缺失数', '总数', '填充率', '备注'])


# ============================ Excel 输出 ============================
def _fill_sheet(worksheet, output_df):
    """把 output_df 按列顺序写进 worksheet(保留表头第1行,清空旧数据行)。列名不影响,只看顺序。"""
    if worksheet.max_row > 1:
        worksheet.delete_rows(2, worksheet.max_row)
    for _, row in output_df.iterrows():
        worksheet.append(['' if pd.isna(v) else v for v in row.tolist()])


def write_to_template(output_df, template_path, output_path, sheet_name):
    """写进导入模版单个 sheet(保留表头与 lov 下拉页),从第 2 行覆盖写入。"""
    wb = load_workbook(template_path)
    _fill_sheet(wb[sheet_name], output_df)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def write_template_sheets(template_path, output_path, sheet_to_df):
    """一次把多个 sheet 写进同一个导入模版(保留各表头与 lov 页)。
    sheet_to_df: {sheet名: DataFrame};DataFrame 列顺序需与该 sheet 表头一致(列名不影响)。"""
    wb = load_workbook(template_path)
    for sheet_name, output_df in sheet_to_df.items():
        _fill_sheet(wb[sheet_name], output_df)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def write_exceptions(output_path, sheets):
    """导出未匹配/待核对清单。sheets: {sheet名: DataFrame}。空表的 sheet 不生成。"""
    non_empty = {name: df for name, df in sheets.items() if len(df) > 0}
    if not non_empty:
        print('  (无任何未匹配/待核对项,跳过清单文件)')
        return None
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path) as writer:
        for sheet_name, sheet_df in non_empty.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output_path
