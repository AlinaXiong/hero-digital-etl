# -*- coding: utf-8 -*-
"""应付期初 —— 对公付款单。整条 ETL 从上到下读完即可。

数据源:
    主表 data/source/ap_payment_opening/uf_dgfktz-主表.xlsx        一行=一张付款申请单
    明细 data/source/ap_payment_opening/uf_dgfktz_dt1-明细表.xlsx   一行=一条费用明细(与主表按 ID 关联)
    规则 data/rules/业财项目_数据映射规则.xlsx
    泛微 vspn_xtyy(工号)   中台 hfins_base(供应商编码) / hfins_base_account(核算主体编码)
模版 data/templates/ap_payment_opening/英雄期初对公付款单导入模版.xlsx
产出 output/ap_payment_opening/英雄期初对公付款单导入_应付期初_<YYYYMMDD>.xlsx + output/ap_payment_opening/未匹配清单_应付期初_<YYYYMMDD>.xlsx

行过滤:流程来源∈{对公付款,个人劳务付款} 且 申请日期>=2026-01-01 且 流程状态=审批完成 且 非作废
行粒度:主子按 ID 合并,一行=一条费用明细(不做分组去重)

跑法:在项目根执行  python run.py ap_payment_opening
"""
import sys
from pathlib import Path

import pandas as pd

if __package__ is None or __package__ == '':
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from etl import common as c

# ---- 文件路径 ----
TASK_NAME = 'ap_payment_opening'
SOURCE_DIR = c.SRC_DIR / TASK_NAME
TEMPLATE_DIR = c.TPL_DIR / TASK_NAME
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
DATE_SUFFIX = c.today_suffix()

SOURCE_MAIN_FILE = SOURCE_DIR / 'uf_dgfktz-主表.xlsx'
SOURCE_DETAIL_FILE = SOURCE_DIR / 'uf_dgfktz_dt1-明细表.xlsx'
TEMPLATE_FILE = TEMPLATE_DIR / '英雄期初对公付款单导入模版.xlsx'
OUTPUT_FILE = OUTPUT_DIR / f'英雄期初对公付款单导入_应付期初_{DATE_SUFFIX}.xlsx'
EXCEPTION_FILE = OUTPUT_DIR / f'未匹配清单_应付期初_{DATE_SUFFIX}.xlsx'
TEMPLATE_SHEET = '期初对公付款单导入'

# ---- 口径 ----
SOURCES = ['对公付款', '个人劳务付款']
DATE_FROM = '2026-01-01'
APPROVED_STATUS = '审批完成'


def filter_main(main_df):
    """主表行过滤:流程来源∈SOURCES 且 申请日期>=DATE_FROM 且 流程状态=审批完成 且 非作废。"""
    df = main_df.copy()
    df['申请日期'] = pd.to_datetime(df['申请日期'], errors='coerce')
    keep_mask = (
        df['流程来源'].isin(SOURCES)
        & (df['申请日期'] >= DATE_FROM)
        & (df['流程状态'] == APPROVED_STATUS)
    )
    matched_count = int(keep_mask.sum())
    void_mask = df['是否作废'].astype(str).str.strip() == '是'
    void_count = int((keep_mask & void_mask).sum())
    result_df = df[keep_mask & ~void_mask].copy()
    print(f"过滤条件: 流程来源∈{SOURCES} 且 申请日期>={DATE_FROM} 且 流程状态='{APPROVED_STATUS}' 且 是否作废≠是")
    print(f'  满足前三项 {matched_count} 单; 其中剔除作废 {void_count} 单; 最终保留主表 {len(result_df)} 单')
    return result_df


def build_output(merged_df, employee_code_map, vendor_map, entity_map, subject_map):
    """主子合并表 -> 导入模版 24 列。每行注释说明该字段取数来源。"""
    def lookup_by_name(mapping, value):  # 按归一化名称查映射字典
        return '' if pd.isna(value) else mapping.get(c.normalize_name(value), '')

    def subject_item(value, index):  # 费用科目 -> (编码, 描述)
        return subject_map.get(c.remove_slashes(value), ('', ''))[index] if pd.notna(value) else ''

    payment_amount = merged_df['付款金额']  # [明细] 本行付款金额
    paid_mask = merged_df['支付状态'].astype(str).str.strip() == '已支付'

    output_df = pd.DataFrame(index=merged_df.index)  # 先定行索引,否则首个标量列会变空
    output_df['来源系统'] = 'FW'  # 固定
    output_df['来源单据编号'] = merged_df['流程编号']  # [主表] 流程编号
    output_df['申请日期'] = merged_df['申请日期'].map(c.format_date)  # [主表] 申请日期
    output_df['单据类型'] = 'AP01-1'  # 固定
    output_df['申请人工号'] = merged_df['经办人'].map(
        lambda value: lookup_by_name(employee_code_map, value))  # [主表] 经办人 -> 泛微工号
    output_df['申请人姓名'] = merged_df['经办人']  # [主表] 经办人
    output_df['订单编号'] = ''  # 留空(待项目 -> 订单映射)
    output_df['订单名称'] = ''
    output_df['核算主体编号'] = merged_df['公司主体'].map(
        lambda value: lookup_by_name(entity_map, value))  # [主表] 公司主体 -> 中台核算主体编码
    output_df['核算主体描述'] = merged_df['公司主体']  # [主表] 公司主体
    output_df['备注'] = merged_df['备注'].astype(str).where(
        merged_df['备注'].notna(), '').str.slice(0, 150)  # 截 150 字
    output_df['合同号'] = merged_df['相关合同'].where(merged_df['相关合同'].notna(), '')  # [主表] 相关合同
    output_df['合同收支计划行'] = ''  # 不涉及
    output_df['收款方编码'] = merged_df['供应商-文本'].map(
        lambda value: lookup_by_name(vendor_map, value))  # [主表] 供应商 -> 中台编码
    output_df['收款方描述'] = merged_df['供应商-文本'].where(merged_df['供应商-文本'].notna(), '')
    output_df['银行账号'] = merged_df['银行账号'].where(merged_df['银行账号'].notna(), '')  # [主表] 银行账号
    output_df['计划付款日期'] = merged_df['预计付款日期'].map(c.format_date)  # [主表] 预计付款日期
    output_df['银行转账备注'] = ''  # 不涉及
    output_df['实际已支付金额'] = [
        c.round_amount(value) if is_paid else 0 for value, is_paid in zip(payment_amount, paid_mask)
    ]  # 已支付则取报账金额,否则为 0
    output_df['费用项目编码'] = merged_df['预算科目'].map(
        lambda value: subject_item(value, 0))  # [明细] 预算科目 -> 科目编码
    output_df['费用项目描述'] = merged_df['预算科目'].map(lambda value: subject_item(value, 1))
    output_df['主播房间号'] = ''  # 不涉及(MCN 才有)
    output_df['报账币种'] = merged_df['付款币种'].map(c.to_iso_currency)  # [主表] 付款币种 -> ISO
    output_df['报账金额（支付币种）'] = payment_amount.map(c.round_amount)  # [明细] 付款金额
    return output_df


def run():
    # 1. 读源表 + 过滤主表
    main_df = pd.read_excel(SOURCE_MAIN_FILE)
    detail_df = pd.read_excel(SOURCE_DETAIL_FILE)
    filtered_main_df = filter_main(main_df)

    # 2. 构建映射(数据库 + 规则表)
    employee_code_map = c.build_employee_code_map()
    vendor_map = c.build_vendor_map()
    entity_map = c.build_accounting_entity_map()
    subject_map = c.build_subject_map()

    # 3. 主子按 ID 合并 + 构建输出
    merged_df = detail_df[detail_df['ID'].isin(set(filtered_main_df['ID']))].merge(
        filtered_main_df, on='ID', suffixes=('_detail', ''), how='inner')
    output_df = build_output(merged_df, employee_code_map, vendor_map, entity_map, subject_map)
    print('输出明细行数:', len(output_df))

    # 4. 填充率(模版中所有有底色的必输字段)
    required_cols = c.required_columns(TEMPLATE_FILE, TEMPLATE_SHEET)
    c.report_fill(output_df, required_cols)

    # 5. 写模版
    c.write_to_template(output_df, TEMPLATE_FILE, OUTPUT_FILE, TEMPLATE_SHEET)
    print('已写出:', OUTPUT_FILE)

    # 6. 未匹配清单
    vendor_names = set(filtered_main_df['供应商-文本'].dropna().astype(str).str.strip())
    company_names = set(filtered_main_df['公司主体'].dropna().astype(str).str.strip())
    subject_names = set(merged_df['预算科目'].dropna().astype(str).str.strip())
    employee_names = set(filtered_main_df['经办人'].dropna().astype(str).str.strip())
    # 映射检查清单:(sheet名, 列标题, 名称集合, 映射字典, 归一化函数);新增一类检查在此加一行
    unmatched_checks = [
        ('未匹配_工号', '未匹配_经办人(工号)', employee_names, employee_code_map, c.normalize_name),
        ('未匹配_供应商', '未匹配_供应商(收款方编码)', vendor_names, vendor_map, c.normalize_name),
        ('未匹配_核算主体', '未匹配_公司主体(核算主体)', company_names, entity_map, c.normalize_name),
        ('未匹配_费用科目', '未匹配_预算科目(费用项目)', subject_names, subject_map, c.remove_slashes),
    ]
    sheets = {'必输字段未达100%': c.fill_summary(output_df, required_cols)}
    sheets.update(c.collect_unmatched(unmatched_checks))  # 没匹配上的才会生成对应 sheet
    c.write_exceptions(EXCEPTION_FILE, sheets)
    print('已写出:', EXCEPTION_FILE, '| 各清单条数:', {
        sheet_name: len(sheet_df) for sheet_name, sheet_df in sheets.items()
    })


if __name__ == '__main__':
    run()
