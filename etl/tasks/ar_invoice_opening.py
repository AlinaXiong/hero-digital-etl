# -*- coding: utf-8 -*-
"""应收期初 —— 应收报账单。整条 ETL 从上到下读完即可。

数据源:
    开票记录 data/source/ar_invoice_opening/uf_xtyykp开票.xlsx          一行=一条开票记录
    收款登记 data/source/ar_invoice_opening/uf_skdj收款登记.xlsx      按「开票/预收单号」汇总核销金额
    规则 data/rules/业财项目_数据映射规则.xlsx「应收期初」sheet
    泛微 vspn_xtyy(工号)   中台 hfins_base(客户编码/税率类型) / hfins_base_account(核算主体编码)
模版 data/templates/ar_invoice_opening/应收报账单期初数据导入模板.xlsx
产出 output/ar_invoice_opening/英雄应收报账单期初数据导入_应收期初_<YYYYMMDD>.xlsx + 未匹配清单_应收期初_<YYYYMMDD>.xlsx

行过滤:申请日期>=2026-01-01 且 开票状态=已开票 且 非作废
行粒度:一行=一条开票记录;收款登记按 开票/预收单号=流程编号 聚合核销金额后回填。

跑法:在项目根执行  python run.py ar_invoice_opening
"""
import sys
from pathlib import Path

import pandas as pd

if __package__ is None or __package__ == '':
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from etl import common as c

# ---- 文件路径 ----
TASK_NAME = 'ar_invoice_opening'
SOURCE_DIR = c.SRC_DIR / TASK_NAME
TEMPLATE_DIR = c.TPL_DIR / TASK_NAME
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
DATE_SUFFIX = c.today_suffix()

SOURCE_INVOICE_FILE = SOURCE_DIR / 'uf_xtyykp开票.xlsx'
SOURCE_RECEIPT_FILE = SOURCE_DIR / 'uf_skdj收款登记.xlsx'
TEMPLATE_FILE = TEMPLATE_DIR / '应收报账单期初数据导入模板.xlsx'
OUTPUT_FILE = OUTPUT_DIR / f'英雄应收报账单期初数据导入_应收期初_{DATE_SUFFIX}.xlsx'
EXCEPTION_FILE = OUTPUT_DIR / f'未匹配清单_应收期初_{DATE_SUFFIX}.xlsx'
TEMPLATE_SHEET = '应收报账单期初数据导入'
RULE_SHEET = '应收期初'
RULE_TABLE = '应收报账单期初数据导入'

# ---- 口径 ----
DATE_FROM = '2026-01-01'
ISSUED_STATUS = '已开票'
DOCUMENT_TYPE = 'OPEN10'
MANAGEMENT_COMPANY = 'Hero'
PAYER_TYPE = '客户'
INCOME_ITEM = '项目收款'
BUSINESS_TYPE_LOV = 'HERO.BUSINESS_TYPE'
INVOICE_TYPE_LOV = 'HERO.INVOICE_TYPE'
CONTRACT_INVOICE_MEANING = '合同'

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

# 缺失明细第二列:{输出必输字段: 泛微源字段};缺失明细展示 来源单据号 + 泛微原表-<源字段>
ISSUE_SOURCE_FIELDS = {
    '来源单据号': '流程编号',
    '核算主体': '公司主体',
    '申请人': '申请人',
    '支付币种': '开票币种',
    '付款对象': '客户',
    '合同编号': '开票合同',
    '业务类型编码': '业务类型',
    '核销金额': '收款登记已收款金额',
    '金额': '开票金额（含税价）',
    '税率类型': '税率',
}


def _text(value):
    if pd.isna(value):
        return ''
    text = str(value).strip()
    return '' if text in ('', 'nan', 'None', 'NaT') else text


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


def filter_invoice(invoice_df):
    """开票记录行过滤:申请日期>=DATE_FROM 且 开票状态=已开票 且 非作废。"""
    df = invoice_df.copy()
    df['申请日期'] = pd.to_datetime(df['申请日期'], errors='coerce')
    keep_mask = (df['申请日期'] >= DATE_FROM) & (df['开票状态'] == ISSUED_STATUS)
    matched_count = int(keep_mask.sum())
    void_mask = df['是否作废'].astype(str).str.strip() == '是'
    void_count = int((keep_mask & void_mask).sum())
    result_df = df[keep_mask & ~void_mask].copy()
    print(f"[应收期初-应收报账单] 过滤条件: 申请日期>={DATE_FROM} 且 开票状态='{ISSUED_STATUS}' 且 是否作废≠是")
    print(f'  满足前两项 {matched_count} 行; 其中剔除作废 {void_count} 行; 最终保留开票记录 {len(result_df)} 行')
    return result_df


def build_receipt_amount_map(receipt_df):
    """收款登记:按 开票/预收单号 汇总已收款金额,作为应收报账单核销金额。"""
    df = receipt_df.copy()
    df['_receipt_key'] = df['开票/预收单号'].astype(str).str.strip()
    df['_receipt_amount'] = pd.to_numeric(df['已收款金额'], errors='coerce').fillna(0)
    return df.groupby('_receipt_key', dropna=False)['_receipt_amount'].sum().to_dict()


def add_receipt_amount(invoice_df, receipt_amount_map):
    """把收款登记汇总金额补回开票记录,供输出和缺失清单复用。"""
    df = invoice_df.copy()
    key = df['流程编号'].astype(str).str.strip()
    df['收款登记已收款金额'] = key.map(receipt_amount_map).fillna(0)
    return df


def build_output(invoice_df, employee_code_map, customer_map, entity_map,
                 business_type_map, invoice_type_map, tax_description_map):
    """开票记录 -> 应收报账单期初导入模板 71 列。每行注释说明该字段取数来源。"""
    contract_invoice_code = invoice_type_map.get(CONTRACT_INVOICE_MEANING, '')  # HERO.INVOICE_TYPE:合同 -> 开票类型编码
    tax_rate_source = invoice_df['税率'].where(  # [开票记录] 优先税率,为空时取税率（明细）
        invoice_df['税率'].notna(), invoice_df.get('税率（明细）'))
    tax_amount = invoice_df['税额（明细）'].where(  # [开票记录] 优先税额（明细）,为空时取税额
        invoice_df['税额（明细）'].notna(), invoice_df['税额'])

    output_df = pd.DataFrame(index=invoice_df.index)
    output_df['来源单据号'] = invoice_df['流程编号']  # [开票记录] 流程编号
    output_df['应收报账单类型'] = DOCUMENT_TYPE  # 缺省 OPEN10
    output_df['核算主体'] = invoice_df['公司主体'].map(lambda value: _lookup_by_name(entity_map, value))  # [开票记录] 公司主体 -> 中台核算主体编码
    output_df['管理公司'] = MANAGEMENT_COMPANY  # 默认 Hero
    output_df['部门'] = invoice_df['申请人部门'].where(invoice_df['申请人部门'].notna(), '')  # [开票记录] 申请人部门
    output_df['岗位'] = ''  # 不涉及
    output_df['申请人'] = invoice_df['申请人'].map(lambda value: _lookup_by_name(employee_code_map, value))  # [开票记录] 申请人 -> 泛微工号
    output_df['申请日期'] = invoice_df['申请日期'].map(c.format_date)  # [开票记录] 申请日期
    output_df['支付币种'] = invoice_df['开票币种'].map(c.to_iso_currency)  # [开票记录] 开票币种 -> ISO
    output_df['付款对象类型'] = PAYER_TYPE  # 默认客户
    output_df['付款对象'] = invoice_df['客户'].map(lambda value: _lookup_by_name(customer_map, value))  # [开票记录] 客户 -> 中台客户编码
    output_df['合同编号'] = invoice_df['开票合同'].where(invoice_df['开票合同'].notna(), '')  # [开票记录] 开票合同
    output_df['里程碑阶段'] = ''  # 汉得顾问后续统一赋值
    output_df['平台'] = ''  # 汉得顾问后续统一赋值
    output_df['业务类型编码'] = invoice_df['业务类型'].map(lambda value: _business_type_code(value, business_type_map))  # [开票记录] 业务类型 -> HERO.BUSINESS_TYPE 编码
    output_df['开票类型编码'] = contract_invoice_code  # 默认合同开票(HERO.INVOICE_TYPE:合同)
    output_df['核销金额'] = pd.to_numeric(invoice_df['收款登记已收款金额'], errors='coerce').fillna(0).map(c.round_amount)  # [收款登记] 已收款金额按单号汇总
    output_df['头备注'] = invoice_df['开票备注'].astype(str).where(invoice_df['开票备注'].notna(), '').str.slice(0, 150)  # [开票记录] 开票备注,截 150 字
    output_df['自审批'] = ''  # 汉得顾问后续统一赋值
    output_df['自审核'] = ''  # 汉得顾问后续统一赋值
    output_df['凭证推送'] = ''  # 汉得顾问后续统一赋值
    output_df['凭证日期'] = ''  # 不涉及
    output_df['行号'] = ''  # 汉得顾问后续统一赋值
    output_df['收入分类'] = ''  # 汉得顾问后续统一赋值
    output_df['收入项目'] = INCOME_ITEM  # 默认项目收款
    output_df['数量'] = ''  # 不涉及
    output_df['单价'] = ''  # 不涉及
    output_df['金额'] = pd.to_numeric(invoice_df['开票金额（含税价）'], errors='coerce').map(c.round_amount)  # [开票记录] 开票金额（含税价）
    output_df['税率类型'] = tax_rate_source.map(lambda value: _tax_description(value, tax_description_map))  # [开票记录] 税率 -> 汉得税率类型描述
    output_df['税额'] = pd.to_numeric(tax_amount, errors='coerce').map(c.round_amount)  # [开票记录] 税额（明细）/税额
    output_df['行备注'] = ''  # 不涉及

    for column in [f'头维度{i}' for i in range(1, 21)]:
        output_df[column] = ''  # 头维度1-20不涉及,留空
    output_df['项目'] = ''  # 待项目/订单清洗结果匹配
    output_df['订单'] = ''  # 待项目/订单清洗结果匹配
    for column in [f'行维度{i}' for i in range(3, 21)]:
        output_df[column] = ''  # 行维度3-20不涉及,留空
    return output_df


def run():
    # 1. 读源表 + 过滤开票记录
    invoice_df = pd.read_excel(SOURCE_INVOICE_FILE, engine='calamine')
    receipt_df = pd.read_excel(SOURCE_RECEIPT_FILE, engine='calamine')
    filtered_invoice_df = filter_invoice(invoice_df)

    # 2. 构建映射(数据库 + 值集)
    employee_code_map = c.build_employee_code_map()
    customer_map = c.build_customer_map()
    entity_map = c.build_accounting_entity_map()
    business_type_map = c.build_lov_meaning_map(BUSINESS_TYPE_LOV)
    invoice_type_map = c.build_lov_meaning_map(INVOICE_TYPE_LOV)
    tax_description_map = c.build_tax_type_description_map(TAX_PREFERRED_DESCRIPTIONS)

    # 3. 收款登记按单号聚合 + 构建输出
    receipt_amount_map = build_receipt_amount_map(receipt_df)
    enriched_invoice_df = add_receipt_amount(filtered_invoice_df, receipt_amount_map)
    output_df = build_output(enriched_invoice_df, employee_code_map, customer_map, entity_map,
                             business_type_map, invoice_type_map, tax_description_map)
    print('[应收期初-应收报账单] 输出明细行数:', len(output_df))

    # 4. 填充率(必输字段以规则表「是否必填」=Y 为准)
    required_cols = c.required_columns(RULE_SHEET, RULE_TABLE)
    c.report_fill(output_df, required_cols)

    # 5. 写模版
    c.write_to_template(output_df, TEMPLATE_FILE, OUTPUT_FILE, TEMPLATE_SHEET)
    print('已写出:', OUTPUT_FILE)

    # 6. 问题清单:必输字段未达100%汇总 + 每个有缺失的必输字段的缺失明细
    sheets = {'必输字段未达100%': c.fill_summary(output_df, required_cols, RULE_SHEET, RULE_TABLE)}
    sheets.update(c.collect_field_issues(output_df, enriched_invoice_df, required_cols,
                                         ISSUE_SOURCE_FIELDS, doc_col='来源单据号'))
    c.write_exceptions(EXCEPTION_FILE, sheets)
    print('已写出:', EXCEPTION_FILE, '| 各清单条数:', {
        sheet_name: len(sheet_df) for sheet_name, sheet_df in sheets.items()
    })


if __name__ == '__main__':
    run()
