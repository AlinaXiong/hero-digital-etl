# -*- coding: utf-8 -*-
"""公共能力:路径配置、数据库只读连接、各类映射、归一化、Excel 读写、行过滤/去重。

所有清洗任务(ap_opening_payment 应付期初,以及将来的应收/预付/预收)都从这里取用,
避免在每个任务里重复写数据库查询和映射逻辑。任务文件本身只关心"过滤 + 字段映射"。

数据库账密从环境变量读取;若项目根有 .env 则自动加载(.env 不提交版本库)。
需要的变量:FW_*(泛微 vspn_xtyy 取工号)、ZT_*(中台 hfins_base 取供应商编码)。
"""
import os
import re
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


# ============================ 归一化 / 格式化 ============================
_PAREN = str.maketrans('（）', '()')


def nz(s):
    """名称归一化:全角括号->半角、去所有空格。消除主体/供应商/人名的全半角差异。"""
    return re.sub(r'\s+', '', str(s).strip().translate(_PAREN))


def no_slash(s):
    """去掉所有 '/'。费用科目里 '/' 既是层级分隔符又可能是层级名内部字符,两侧统一去掉才能对齐。"""
    return str(s).strip().replace('/', '')


def fmt_code(v):
    """浮点编码(1000.0)->整数串'1000';空值->''。"""
    s = str(v).strip()
    if s in ('', 'nan', 'None'):
        return ''
    try:
        return str(int(float(s)))
    except ValueError:
        return s


def fdate(x):
    """日期 -> 'yyyy-mm-dd';空值->''。"""
    return '' if pd.isna(x) else pd.to_datetime(x).strftime('%Y-%m-%d')


def amt2(x):
    """金额保留 2 位小数;空值->''。"""
    return '' if pd.isna(x) else round(float(x), 2)


# ============================ 映射字典 ============================
# 币种 -> ISO 码
CUR = {
    '人民币': 'CNY', '美元': 'USD', '马来西亚令吉': 'MYR', '泰铢': 'THB', '印尼盾': 'IDR',
    '韩元': 'KRW', '港币': 'HKD', '新加坡元': 'SGD', '沙特里亚尔': 'SAR', '菲律宾比索': 'PHP',
    '欧元': 'EUR', '日元': 'JPY', '英镑': 'GBP', '瑞士法郎': 'CHF',
    '伊拉克第纳尔': 'IQD', '科威特第纳尔': 'KWD', '埃及镑': 'EGP',
}


def to_iso_currency(name):
    s = '' if pd.isna(name) else str(name).strip()
    return CUR.get(s, s)


def build_gonghao_map():
    """经办人姓名 -> 工号。来源:泛微 vspn_xtyy。
    hrmresource.LASTNAME 经 JOBTITLE=id 关联 hrmjobtitles.JOBTITLENAME(工号,如 V81982);
    剔除占位 'Default';键用 nz 归一化(外籍名带空格也能对上)。"""
    conn = _db_connect('FW', 'vspn_xtyy')
    try:
        emp = pd.read_sql('SELECT r.LASTNAME nm, j.JOBTITLENAME gh '
                          'FROM hrmresource r LEFT JOIN hrmjobtitles j ON r.JOBTITLE=j.id', conn)
    finally:
        conn.close()
    emp['key'] = emp['nm'].map(nz)
    gh = {}
    for key, g in emp.groupby('key')['gh']:
        codes = [str(x).strip() for x in g.dropna().unique()
                 if str(x).strip() not in ('', 'Default', 'nan')]
        if key and codes:
            gh[key] = codes[0]
    return gh


def build_vendor_map():
    """供应商名称 -> 中台供应商编码 vender_code。来源:中台 hfins_base.hfbs_system_vender。
    按 description / taxpayer_name 建键(均 nz 归一化)。"""
    conn = _db_connect('ZT', 'hfins_base')
    try:
        ven = pd.read_sql('SELECT vender_code code, description nm, taxpayer_name tnm '
                          'FROM hfbs_system_vender', conn)
    finally:
        conn.close()
    vm = {}
    for _, r in ven.iterrows():
        for name in (r['nm'], r['tnm']):
            k = nz(name)
            if k and k not in ('nan', 'None') and k not in vm:
                vm[k] = str(r['code']).strip()
    return vm


def build_entity_map(code_col='新主体编码'):
    """公司主体名称 -> 核算主体编号。来源:规则表「新旧主体映射」。
    code_col 指定取哪列做编号(默认 新主体编码)。匹配键:中台主体名称/主体名称更正/纳税人名称/更新纳税人名称。"""
    ent = pd.read_excel(RULE_XLSX, sheet_name='新旧主体映射')
    em = {}
    for _, r in ent.iterrows():
        code = fmt_code(r[code_col])
        if not code:
            continue
        for col in ('中台主体名称', '主体名称更正', '纳税人名称', '更新纳税人名称'):
            k = nz(r.get(col, ''))
            if k and k != 'nan' and k not in em:
                em[k] = code
    return em


def build_subject_map():
    """费用科目(预算科目)-> (费用项目编码, 费用项目描述)。
    来源:规则表「赛事/MCN 新旧预算项科目-调整后」。键=各级科目名拼接去'/';值=调整后三级预算编码+三级费用项。"""
    sm = {}
    s1 = pd.read_excel(RULE_XLSX, sheet_name='赛事新旧预算项科目-调整后', header=None)
    for _, r in s1.iloc[2:].iterrows():
        k = no_slash(r[15])
        code, nm = str(r[21]).strip(), str(r[22]).strip()
        if k and k != 'nan' and code != 'nan':
            sm[k] = (code, nm)
    s2 = pd.read_excel(RULE_XLSX, sheet_name='MCN新旧预算项科目-调整后', header=None)
    for _, r in s2.iloc[2:].iterrows():
        key = no_slash(''.join(str(r[c]) for c in (1, 3, 5) if str(r[c]) != 'nan'))
        code = str(r[11]).strip() if str(r[11]) != 'nan' else str(r[7]).strip()
        nm = str(r[8]).strip()
        if key and code != 'nan':
            sm.setdefault(key, (code, nm))
    return sm


# ============================ 行过滤 / 去重 / 统计 ============================
def filter_main(m, sources, date_from='2026-01-01', status='审批完成',
                date_col='申请日期', drop_void=True):
    """主表标准行过滤,返回过滤后副本。"""
    m = m.copy()
    m[date_col] = pd.to_datetime(m[date_col], errors='coerce')
    keep = m['流程来源'].isin(sources) & (m[date_col] >= date_from) & (m['流程状态'] == status)
    n_total, n_void = int(keep.sum()), 0
    if drop_void and '是否作废' in m.columns:
        void = m['是否作废'].astype(str).str.strip() == '是'
        n_void = int((keep & void).sum())
        keep = keep & ~void
    out = m[keep].copy()
    print(f'过滤: 范围+日期+状态={n_total}单; 作废剔除={n_void}单; 最终主表={len(out)}单')
    return out


def dedup_rows(out_df, key_cols):
    """按 key_cols 完全相同则合并为一条。返回(去重后, 参与合并的全部行[供核对])。"""
    k = out_df[key_cols].apply(lambda r: '|'.join(map(str, r.tolist())), axis=1)
    dup = k.duplicated(keep=False)
    collapsed = out_df[dup].assign(_k=k).sort_values('_k').drop(columns='_k')
    deduped = out_df[~k.duplicated(keep='first')].reset_index(drop=True)
    print(f'分组去重: 参与合并 {int(dup.sum())} 行 -> 最终输出 {len(deduped)} 行')
    return deduped, collapsed


def report_fill(out_df, cols):
    for c in cols:
        n = (out_df[c].astype(str).str.strip() != '').sum()
        print(f'  {c} 填充率: {n}/{len(out_df)} = {n/len(out_df)*100:.1f}%')


# ============================ Excel 输出 ============================
def write_to_template(out_df, template_path, out_path, sheet_name):
    """写进导入模版(保留表头与 lov 下拉页),从第 2 行覆盖写入。"""
    wb = load_workbook(template_path)
    ws = wb[sheet_name]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)
    for _, row in out_df.iterrows():
        ws.append(['' if pd.isna(v) else v for v in row.tolist()])
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def write_exceptions(out_path, sheets):
    """导出未匹配/待核对清单。sheets: {sheet名: DataFrame}。"""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path) as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name, index=False)
    return out_path
