# -*- coding: utf-8 -*-
"""预付期初 —— 供应商预付款单 / 投资付款单。整条 ETL 从上到下读完即可。

数据源:
    主表 data/source/ap_prepayment_opening/uf_yfkxx预付.xlsx     一行=一张预付款单(预付款信息)
    明细 data/source/ap_prepayment_opening/uf_yfkxx_dt1.xlsx     一行=一条预付预算明细(与主表按 ID 关联)
    规则 data/rules/业财项目_数据映射规则.xlsx「预付期初」sheet(R3-R28)
    泛微 vspn_xtyy(工号)   中台 hfins_base(供应商编码) / hfins_base_account(核算主体编码)
模版 data/templates/ap_prepayment_opening/英雄期初预付款单导入模版.xlsx
    - Tab「期初供应商预付款单&期初投资付款单导入」:本任务输出(26 列)
    - Tab「期初灵工预付款单导入」:灵工(零工平台付款),源 = 零工平台付款_收款人明细_2026.xlsx
      (付款头数据=uf_lgptfk + 实际收款人明细),过滤 流程状态=2(审批完成)+ 申请日期>=2026-01-01 + 非作废。
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
RULE_SHEET = '预付期初'                                       # 规则表 sheet
RULE_TABLE_SUPPLIER = '期初供应商预付款单&期初投资付款单导入'  # 规则表内供应商预付目标表名

# ---- 口径 ----
APPROVED_STATUS = '审批完成'
DATE_FROM = '2026-01-01'           # 只取今年(申请日期>=本日期)
DOCUMENT_TYPE = 'JK01-2'           # 供应商预付单(期初);投资付款单为 PP02-2,规则未给判定标准,暂统一 JK01-2,待业务确认
DEPOSIT_PAYMENT_NATURES = ['押金', '质保金']  # 付款性质命中这些 -> 保证金标志=是
# 缺失明细第二列:{输出必输字段: 泛微源字段};缺失明细展示 来源单据编号 + 泛微原表-<源字段>
ISSUE_SOURCE_FIELDS = {
    '申请人工号': '填单人',
    '收款方编码': '付款对象',
    '核算主体编号': '开票单位',
    '费用项目编码': '预算科目',
    '预付款支付币种': '付款币种',
    '订单编号': '项目编号',
}

# ======== 灵工预付款单(模版第二个 tab:期初灵工预付款单导入)========
GIG_SOURCE_FILE = SOURCE_DIR / '零工平台付款_收款人明细_2026.xlsx'
GIG_HEADER_SHEET = '付款头数据'        # 单头(uf_lgptfk,已含申请人工号V码、公司主体ID等)
GIG_DETAIL_SHEET = '实际收款人明细'    # 收款人(按 建模付款ID 关联单头)
TEMPLATE_SHEET_GIG = '期初灵工预付款单导入'
RULE_TABLE_GIG = '期初灵工预付款单导入'  # 规则表内灵工预付目标表名
GIG_DOCUMENT_TYPE = 'PP01-2'           # 灵工预付款单(期初)
GIG_APPROVED_CODE = 2                  # 付款头数据 流程状态数字码:2=审批完成(经与 uf_dgfktz 中文状态交叉验证)
# 灵工平台收款方编码:按收款方文本关键字映射(规则 R40)
GIG_PLATFORM_VENDOR = {'云账户': 'V-C-CN-HR-PAY-0001', '赛利得': 'V-C-CN-OT-OTH-6573'}
GIG_ISSUE_SOURCE_FIELDS = {
    '申请人工号': '经办人',
    '灵工平台收款方编码': '收款方文本',
    '核算主体编号': '公司主体ID',
    '收款方编码': '实际收款方',
}


def filter_main(main_df):
    """预付主表行过滤:申请日期>=DATE_FROM(今年)且 流程状态=审批完成 且 非作废。"""
    df = main_df.copy()
    df['申请日期'] = pd.to_datetime(df['申请日期'], errors='coerce')
    keep_mask = (df['申请日期'] >= DATE_FROM) & (df['流程状态'] == APPROVED_STATUS)
    matched_count = int(keep_mask.sum())
    void_mask = df['是否作废'].astype(str).str.strip() == '是'
    void_count = int((keep_mask & void_mask).sum())
    result_df = df[keep_mask & ~void_mask].copy()
    print(f"供应商过滤条件: 申请日期>={DATE_FROM} 且 流程状态='{APPROVED_STATUS}' 且 是否作废≠是")
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


# ======================= 灵工预付款单(模版第二个 tab) =======================
def gig_platform_vendor(name):
    """收款方文本 -> 灵工平台收款方编码(规则 R40:云账户/赛利得)。"""
    text = '' if pd.isna(name) else str(name)
    for keyword, code in GIG_PLATFORM_VENDOR.items():
        if keyword in text:
            return code
    return ''


def gig_recipient_remark(name, id_number, phone):
    """收款人备注(规则 R50):姓名-身份证-手机号 拼接,限 30 字。"""
    parts = [str(v).strip() for v in (name, id_number, phone) if pd.notna(v) and str(v).strip() not in ('', 'nan')]
    return '-'.join(parts)[:30]


def build_gig_output(header_df, detail_df, vendor_map, company_map, entity_map):
    """灵工:付款头数据(单头)+ 实际收款人明细(收款人)-> 模版「期初灵工预付款单导入」29 列。
    一行=一个收款人。返回 (输出表, 关联后明细表)。"""
    header = header_df.copy()
    header['申请日期'] = pd.to_datetime(header['申请日期'], errors='coerce')
    matched_mask = (header['流程状态'] == GIG_APPROVED_CODE) & (header['申请日期'] >= DATE_FROM)
    matched_count = int(matched_mask.sum())
    void_mask = header['是否作废'].astype(str).str.strip().isin(['是', '1', '1.0'])
    void_count = int((matched_mask & void_mask).sum())
    header_kept = header[matched_mask & ~void_mask]
    print(f"零工过滤条件: 申请日期>={DATE_FROM} 且 流程状态={GIG_APPROVED_CODE}(审批完成) 且 是否作废≠是")
    print(f'  满足前两项 {matched_count} 单; 其中剔除作废 {void_count} 单; 最终保留主表 {len(header_kept)} 单')

    # 取单头补充字段,与明细按 建模付款ID 关联(明细已含 流程编号/申请日期/收款人/金额/银行账号)
    header_cols = ['建模付款ID', '经办人', '经办人工号', '收款方文本', '备注', '合同名称', '预计付款日期', '公司主体ID']
    merged = detail_df.merge(header_kept[header_cols], on='建模付款ID', how='inner')
    if '公司主体ID' not in merged.columns:
        for column in ('公司主体ID_y', '公司主体ID_x'):
            if column in merged.columns:
                merged['公司主体ID'] = merged[column]
                break
    if '公司主体ID' not in merged.columns:
        merged['公司主体ID'] = ''
    print('[零工预付款] 输出明细行数:', len(merged))

    company_names = merged['公司主体ID'].map(lambda value: company_map.get(c.format_code(value), ''))

    def gig_payee_code(value):
        if pd.isna(value):
            return ''
        payee = str(value).strip()
        if not payee or payee == 'nan':
            return ''
        key = c.normalize_name(value)
        return vendor_map.get(key) or payee

    out = pd.DataFrame(index=merged.index)
    out['来源系统'] = 'FW'                                                      # 固定
    out['来源单据编号'] = merged['流程编号']                                     # [头/明细] 流程编号
    out['申请日期'] = merged['申请日期'].map(c.format_date)                      # [明细] 申请日期
    out['单据类型'] = GIG_DOCUMENT_TYPE                                         # PP01-2
    out['申请人工号'] = merged['经办人工号']                                     # [头] 经办人工号(V码,现成)
    out['申请人姓名'] = merged['经办人']                                         # [头] 经办人
    out['订单编号'] = ''                                                        # 留空(待项目->订单映射)
    out['订单名称'] = ''
    out['核算主体编号'] = company_names.map(
        lambda value: entity_map.get(c.normalize_name(value), '') if value else '')  # 公司主体ID -> 泛微公司名 -> 中台核算主体编码
    out['核算主体描述'] = company_names
    out['备注_单头'] = merged['备注'].astype(str).where(merged['备注'].notna(), '').str.slice(0, 150)  # 模版第11列(单头备注)
    out['灵工平台收款方编码'] = merged['收款方文本'].map(gig_platform_vendor)      # [头] 收款方文本 -> 平台编码
    out['合同号'] = merged['合同名称'].where(merged['合同名称'].notna(), '')      # [头] 合同名称
    out['合同收支计划行'] = ''                                                  # 不涉及
    out['保证金标志'] = '否'                                                    # 默认否(R43)
    out['计划付款日期'] = merged['预计付款日期'].map(c.format_date)              # [头] 预计付款日期
    out['银行转账备注'] = ''                                                    # 不涉及
    out['费用项目编码'] = ''                                                    # 规则R46缺表,无来源,暂空
    out['费用项目描述'] = ''
    out['收款方类别'] = '供应商'                                                # 默认供应商(R48)
    out['收款方编码'] = merged['实际收款方'].map(gig_payee_code)                 # [明细] 实际收款方 -> 中台供应商编码;未命中则保留实际收款方
    out['备注'] = [gig_recipient_remark(n, i, p)                                # 模版第22列:姓名-身份证-手机号(R50)
                  for n, i, p in zip(merged['实际收款方'], merged['身份证号'], merged['手机号'])]
    out['银行账号'] = merged['银行账号'].where(merged['银行账号'].notna(), '')   # [明细] 银行账号
    out['预付款支付币种'] = 'CNY'                                               # 默认CNY(R52)
    out['预付款金额（支付币种）'] = pd.to_numeric(merged['付给三方平台金额'], errors='coerce').map(c.round_amount)  # [明细] 付给三方平台金额
    out['传送状态'] = '传送成功'                                                # 默认(R54)
    out['支付状态'] = '支付成功'                                                # 默认支付成功(R55)
    out['退款状态'] = ''                                                        # 不涉及(R56)
    out['核销状态'] = '已核销'                                                  # 默认已核销(R57)
    return out, merged


def run():
    # 1. 读源表 + 过滤主表
    main_df = pd.read_excel(SOURCE_MAIN_FILE)
    detail_df = pd.read_excel(SOURCE_DETAIL_FILE)
    filtered_main_df = filter_main(main_df)

    # 2. 构建映射(数据库 + 规则表)
    employee_code_map = c.build_employee_code_map()
    vendor_map = c.build_vendor_map()
    entity_map = c.build_accounting_entity_map()
    company_map = c.build_fw_company_map()
    subject_map = c.build_subject_map()

    # 3. 供应商预付款 tab:主子按 ID 合并 + 构建输出
    merged_df = detail_df[detail_df['ID'].isin(set(filtered_main_df['ID']))].merge(
        filtered_main_df, on='ID', suffixes=('_detail', ''), how='inner')
    output_df = build_output(merged_df, employee_code_map, vendor_map, entity_map, subject_map)
    print('[供应商预付款] 输出明细行数:', len(output_df))

    # 4. 灵工预付款 tab:付款头数据 + 实际收款人明细
    gig_header_df = pd.read_excel(GIG_SOURCE_FILE, sheet_name=GIG_HEADER_SHEET)
    gig_detail_df = pd.read_excel(GIG_SOURCE_FILE, sheet_name=GIG_DETAIL_SHEET)
    gig_output_df, gig_merged_df = build_gig_output(gig_header_df, gig_detail_df, vendor_map, company_map, entity_map)

    # 5. 填充率(必输字段以规则表「是否必填」=Y 为准)
    supplier_required = c.required_columns(RULE_SHEET, RULE_TABLE_SUPPLIER)
    gig_required = c.required_columns(RULE_SHEET, RULE_TABLE_GIG)
    print('— 供应商预付款 填充率 —')
    c.report_fill(output_df, supplier_required)
    print('— 灵工预付款 填充率 —')
    c.report_fill(gig_output_df, gig_required)

    # 6. 写模版(两个 tab 一次写入,lov 页保留)
    c.write_template_sheets(TEMPLATE_FILE, OUTPUT_FILE, {
        TEMPLATE_SHEET_SUPPLIER: output_df,
        TEMPLATE_SHEET_GIG: gig_output_df,
    })
    print('已写出:', OUTPUT_FILE)

    # 7. 问题清单(一个文件,sheet 名按 tab 前缀「供应商_」「灵工_」区分):必输字段未达100% + 缺失明细
    exception_sheets = {}
    supplier_sheets = {'必输字段未达100%': c.fill_summary(output_df, supplier_required)}
    supplier_sheets.update(c.collect_field_issues(output_df, merged_df, supplier_required, ISSUE_SOURCE_FIELDS))
    exception_sheets.update({f'供应商_{name}': df for name, df in supplier_sheets.items()})
    gig_sheets = {'必输字段未达100%': c.fill_summary(gig_output_df, gig_required)}
    gig_sheets.update(c.collect_field_issues(gig_output_df, gig_merged_df, gig_required, GIG_ISSUE_SOURCE_FIELDS))
    exception_sheets.update({f'灵工_{name}': df for name, df in gig_sheets.items()})
    c.write_exceptions(EXCEPTION_FILE, exception_sheets)
    print('已写出:', EXCEPTION_FILE, '| 各清单条数:', {k: len(v) for k, v in exception_sheets.items()})


if __name__ == '__main__':
    run()
