# -*- coding: utf-8 -*-
"""预付期初 —— 供应商预付款单 / 投资付款单。整条 ETL 从上到下读完即可。

数据源:
    主表 data/source/ap_prepayment_opening/uf_yfkxx预付.xlsx     一行=一张预付款单(预付款信息)
    明细 data/source/ap_prepayment_opening/uf_yfkxx_dt1.xlsx     一行=一条预付预算明细(与主表按 ID 关联)
    规则 data/rules/业财项目_数据映射规则.xlsx「预付期初」sheet(R3-R28)
    泛微 vspn_xtyy(工号)   中台 hfins_base(供应商编码) / hfins_base_account(核算主体编码)
模版 data/templates/ap_prepayment_opening/英雄期初预付款单导入模版.xlsx
    - Tab「期初供应商预付款单&期初投资付款单导入」:本任务输出(26 列)
    - Tab「期初灵工预付款单导入」:灵工(零工平台付款)专用,数据源为对外付款主表 + 零工实际收款人明细,
      不在本任务的源文件内,暂不生成(见文末说明)。
产出 output/ap_prepayment_opening/英雄期初预付款单导入_预付期初_<YYYYMMDD>.xlsx + 未匹配清单_预付期初_<YYYYMMDD>.xlsx

行过滤:申请日期>=2026-01-01(今年)且 流程状态=审批完成 且 非作废
行粒度:主子按 ID 合并,一行=一条费用明细(不做分组去重)

金额拆分(规则整体说明 点5):
    预付款金额 = 本明细「预付金额」
    已付未核   = 主表「剩余冲销/退款金额」按明细金额占比分摊到本行
    已到票核销 = 预付款金额 - 已付未核

跑法:在项目根执行  python run.py ap_prepayment_opening
"""
import sys
from pathlib import Path

import pandas as pd

if __package__ is None or __package__ == '':
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from etl import common as c

# ---- 文件路径 ----
TASK_NAME = 'ap_prepayment_opening'
SOURCE_DIR = c.SRC_DIR / TASK_NAME
TEMPLATE_DIR = c.TPL_DIR / TASK_NAME
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
DATE_SUFFIX = c.today_suffix()

SOURCE_MAIN_FILE = SOURCE_DIR / 'uf_yfkxx预付.xlsx'
SOURCE_DETAIL_FILE = SOURCE_DIR / 'uf_yfkxx_dt1.xlsx'
TEMPLATE_FILE = TEMPLATE_DIR / '英雄期初预付款单导入模版.xlsx'
OUTPUT_FILE = OUTPUT_DIR / f'英雄期初预付款单导入_预付期初_{DATE_SUFFIX}.xlsx'
EXCEPTION_FILE = OUTPUT_DIR / f'未匹配清单_预付期初_{DATE_SUFFIX}.xlsx'
TEMPLATE_SHEET_SUPPLIER = '期初供应商预付款单&期初投资付款单导入'

# ---- 口径 ----
APPROVED_STATUS = '审批完成'
DATE_FROM = '2026-01-01'           # 只取今年(申请日期>=本日期)
DOCUMENT_TYPE = 'JK01-2'           # 供应商预付单(期初);投资付款单为 PP02-2,规则未给判定标准,暂统一 JK01-2,待业务确认
DEPOSIT_PAYMENT_NATURES = ['押金', '质保金']  # 付款性质命中这些 -> 保证金标志=是


def filter_main(main_df):
    """预付主表行过滤:申请日期>=DATE_FROM(今年)且 流程状态=审批完成 且 非作废。"""
    df = main_df.copy()
    df['申请日期'] = pd.to_datetime(df['申请日期'], errors='coerce')
    keep_mask = (df['申请日期'] >= DATE_FROM) & (df['流程状态'] == APPROVED_STATUS)
    matched_count = int(keep_mask.sum())
    void_mask = df['是否作废'].astype(str).str.strip() == '是'
    void_count = int((keep_mask & void_mask).sum())
    result_df = df[keep_mask & ~void_mask].copy()
    print(f"过滤条件: 申请日期>={DATE_FROM} 且 流程状态='{APPROVED_STATUS}' 且 是否作废≠是")
    print(f'  满足前两项 {matched_count} 单; 其中剔除作废 {void_count} 单; 最终保留主表 {len(result_df)} 单')
    return result_df


def build_output(merged_df, employee_code_map, vendor_map, entity_map, subject_map):
    """主子合并表 -> 供应商预付款单导入模版 26 列。每行注释说明该字段取数来源。"""
    def lookup_by_name(mapping, value):  # 按归一化名称查映射字典
        return '' if pd.isna(value) else mapping.get(c.normalize_name(value), '')

    def subject_item(value, index):  # 费用科目 -> (编码, 描述)
        return subject_map.get(c.remove_slashes(value), ('', ''))[index] if pd.notna(value) else ''

    # 金额:本费用行金额取明细「预付金额」;已付未核按明细占比从主表「剩余冲销/退款金额」分摊
    main_amount = pd.to_numeric(merged_df['付款金额'], errors='coerce').fillna(0)        # 主表预付款总额
    detail_amount = pd.to_numeric(merged_df['预付金额'], errors='coerce').fillna(0)      # 本行预付款金额
    remaining = pd.to_numeric(merged_df['剩余冲销/退款金额'], errors='coerce').fillna(0)  # 主表剩余冲销/退款=已付未核(单级)
    ratio = (detail_amount / main_amount).where(main_amount != 0, 0)                     # 明细占比
    unsettled_amount = remaining * ratio                                                 # 已付未核(费用行)
    settled_amount = detail_amount - unsettled_amount                                    # 已到票核销(费用行)
    is_deposit = merged_df['付款性质'].isin(DEPOSIT_PAYMENT_NATURES)

    output_df = pd.DataFrame(index=merged_df.index)  # 先定行索引,否则首个标量列会变空
    output_df['来源系统'] = 'FW'  # 固定
    output_df['来源单据编号'] = merged_df['流程编号']  # [主表] 流程编号
    output_df['申请日期'] = merged_df['申请日期'].map(c.format_date)  # [主表] 申请日期
    output_df['单据类型'] = DOCUMENT_TYPE  # 供应商预付单(期初) JK01-2
    output_df['申请人工号'] = merged_df['填单人'].map(
        lambda value: lookup_by_name(employee_code_map, value))  # [主表] 填单人 -> 泛微工号
    output_df['申请人姓名'] = merged_df['填单人']  # [主表] 填单人
    output_df['订单编号'] = ''  # 留空(待项目 -> 订单映射)
    output_df['订单名称'] = ''
    output_df['核算主体编号'] = merged_df['开票单位'].map(
        lambda value: lookup_by_name(entity_map, value))  # [主表] 开票单位 -> 中台核算主体编码
    output_df['核算主体描述'] = merged_df['开票单位']  # [主表] 开票单位
    output_df['备注'] = merged_df['备注'].astype(str).where(
        merged_df['备注'].notna(), '').str.slice(0, 150)  # 截 150 字
    output_df['合同号'] = merged_df['相关合同'].where(merged_df['相关合同'].notna(), '')  # [主表] 相关合同
    output_df['合同收支计划行'] = ''  # 不涉及
    output_df['保证金标志'] = ['是' if flag else '否' for flag in is_deposit]  # [主表] 付款性质=押金/质保金 -> 是
    output_df['收款方编码'] = merged_df['付款对象'].map(
        lambda value: lookup_by_name(vendor_map, value))  # [主表] 付款对象 -> 中台供应商编码
    output_df['收款方描述'] = merged_df['付款对象'].where(merged_df['付款对象'].notna(), '')
    output_df['银行账号'] = merged_df['银行卡号'].where(merged_df['银行卡号'].notna(), '')  # [主表] 银行卡号
    output_df['计划付款日期'] = ''  # 不涉及
    output_df['银行转账备注'] = ''  # 不涉及
    output_df['费用项目编码'] = merged_df['预算科目'].map(
        lambda value: subject_item(value, 0))  # [明细] 预算科目 -> 科目编码
    output_df['费用项目描述'] = merged_df['预算科目'].map(lambda value: subject_item(value, 1))
    output_df['主播房间号'] = ''  # 不涉及(MCN 主播才有)
    output_df['预付款支付币种'] = merged_df['付款币种'].map(c.to_iso_currency)  # [主表] 付款币种 -> ISO
    output_df['预付款金额（支付币种）'] = detail_amount.map(c.round_amount)  # [明细] 预付金额
    output_df['已到票核销金额（支付币种）'] = settled_amount.map(c.round_amount)  # 预付款金额 - 已付未核
    output_df['已付未核（支付币种）'] = unsettled_amount.map(c.round_amount)  # 剩余冲销/退款按占比分摊
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
    c.report_fill(output_df, c.required_columns(TEMPLATE_FILE, TEMPLATE_SHEET_SUPPLIER))

    # 5. 写模版(只写供应商预付款 tab;灵工 tab 与 lov 页保留不动)
    c.write_to_template(output_df, TEMPLATE_FILE, OUTPUT_FILE, TEMPLATE_SHEET_SUPPLIER)
    print('已写出:', OUTPUT_FILE)

    # 6. 未匹配清单
    employee_names = set(filtered_main_df['填单人'].dropna().astype(str).str.strip())
    vendor_names = set(filtered_main_df['付款对象'].dropna().astype(str).str.strip())
    entity_names = set(filtered_main_df['开票单位'].dropna().astype(str).str.strip())
    subject_names = set(merged_df['预算科目'].dropna().astype(str).str.strip())
    sheets = {
        '未匹配_工号': pd.DataFrame({
            '未匹配_填单人(工号)': sorted(
                name for name in employee_names if c.normalize_name(name) not in employee_code_map)
        }),
        '未匹配_供应商': pd.DataFrame({
            '未匹配_付款对象(收款方编码)': sorted(
                name for name in vendor_names if c.normalize_name(name) not in vendor_map)
        }),
        '未匹配_核算主体': pd.DataFrame({
            '未匹配_开票单位(核算主体)': sorted(
                name for name in entity_names if c.normalize_name(name) not in entity_map)
        }),
        '未匹配_费用科目': pd.DataFrame({
            '未匹配_预算科目(费用项目)': sorted(
                name for name in subject_names if c.remove_slashes(name) not in subject_map)
        }),
    }
    c.write_exceptions(EXCEPTION_FILE, sheets)
    print('已写出:', EXCEPTION_FILE, '| 各清单条数:', {
        sheet_name: len(sheet_df) for sheet_name, sheet_df in sheets.items()
    })


if __name__ == '__main__':
    run()
