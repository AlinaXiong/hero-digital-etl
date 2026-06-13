# -*- coding: utf-8 -*-
"""应付期初 —— 对公付款单。整条 ETL 从上到下读完即可。

数据源:
    主表 data/source/uf_dgfktz-主表.xlsx        一行=一张付款申请单
    明细 data/source/uf_dgfktz_dt1-明细表.xlsx   一行=一条费用明细(与主表按 ID 关联)
    规则 data/rules/业财项目_数据映射规则.xlsx
    泛微 vspn_xtyy(工号)   中台 hfins_base(供应商编码)
模版 data/templates/英雄期初对公付款单导入模版.xlsx
产出 output/英雄期初对公付款单导入_应付期初_v2.xlsx + output/未匹配清单_应付期初_v2.xlsx

行过滤:流程来源∈{对公付款,个人劳务付款} 且 申请日期>=2026-01-01 且 流程状态=审批完成 且 非作废
行粒度:主子按 ID 合并(一行=一条费用明细),再按 单头键+付款行键+费用行键 分组去重

跑法:在项目根执行  python run.py ap_qichu
"""
import pandas as pd

import common as c

# ---- 文件路径 ----
SRC_M = c.SRC_DIR / 'uf_dgfktz-主表.xlsx'
SRC_D = c.SRC_DIR / 'uf_dgfktz_dt1-明细表.xlsx'
TMPL  = c.TPL_DIR / '英雄期初对公付款单导入模版.xlsx'
OUT   = c.OUT_DIR / '英雄期初对公付款单导入_应付期初_v2.xlsx'
EXC   = c.OUT_DIR / '未匹配清单_应付期初_v2.xlsx'
TMPL_SHEET = '期初对公付款单导入'

# ---- 口径 ----
SOURCES = ['对公付款', '个人劳务付款']
# 分组去重键(整体说明 点1/2/3:单头键+付款行键+费用行键)
DEDUP_KEY = ['来源系统', '来源单据编号', '申请日期', '单据类型', '申请人工号', '订单编号',
             '核算主体编号', '备注', '收款方编码', '报账币种', '计划付款日期', '合同号',
             '合同收支计划行', '银行转账备注', '报账金额（支付币种）', '费用项目编码', '主播房间号']


def build_output(df, gh_map, ven_map, ent_map, sub_map):
    """主子合并表 df -> 导入模版 24 列。每行注释 = 该字段取数来源。"""
    def by_name(d, x):                      # 按归一化名称查映射字典
        return '' if pd.isna(x) else d.get(c.nz(x), '')

    def subj(x, i):                         # 费用科目 -> (编码, 描述)
        return sub_map.get(c.no_slash(x), ('', ''))[i] if pd.notna(x) else ''

    pay = df['付款金额']                                                 # [明细] 本行付款金额
    paid = df['支付状态'].astype(str).str.strip() == '已支付'

    out = pd.DataFrame(index=df.index)      # 先定行索引,否则首个标量列会变空
    out['来源系统']        = 'FW'                                          # 固定
    out['来源单据编号']    = df['流程编号']                                 # [主表] 流程编号
    out['申请日期']        = df['申请日期'].map(c.fdate)                    # [主表] 申请日期
    out['单据类型']        = 'AP01-1'                                      # 固定
    out['申请人工号']      = df['经办人'].map(lambda x: by_name(gh_map, x))  # [主表]经办人->泛微工号
    out['申请人姓名']      = df['经办人']                                   # [主表] 经办人
    out['订单编号']        = ''                                            # 留空(待项目->订单映射)
    out['订单名称']        = ''
    out['核算主体编号']    = df['公司主体'].map(lambda x: by_name(ent_map, x))  # [主表]公司主体->新主体编码
    out['核算主体描述']    = df['公司主体']                                 # [主表] 公司主体
    out['备注']            = df['备注'].astype(str).where(df['备注'].notna(), '').str.slice(0, 150)  # 截150字
    out['合同号']          = df['相关合同'].where(df['相关合同'].notna(), '')  # [主表] 相关合同
    out['合同收支计划行']  = ''                                            # 不涉及
    out['收款方编码']      = df['供应商-文本'].map(lambda x: by_name(ven_map, x))  # [主表]供应商->中台编码
    out['收款方描述']      = df['供应商-文本'].where(df['供应商-文本'].notna(), '')
    out['银行账号']        = df['银行账号'].where(df['银行账号'].notna(), '')  # [主表] 银行账号
    out['计划付款日期']    = df['预计付款日期'].map(c.fdate)                 # [主表] 预计付款日期
    out['银行转账备注']    = ''                                            # 不涉及
    out['实际已支付金额']  = [c.amt2(v) if p else 0 for v, p in zip(pay, paid)]  # 已支付?报账金额:0
    out['费用项目编码']    = df['预算科目'].map(lambda x: subj(x, 0))        # [明细]预算科目->科目编码
    out['费用项目描述']    = df['预算科目'].map(lambda x: subj(x, 1))
    out['主播房间号']      = ''                                            # 不涉及(MCN才有)
    out['报账币种']        = df['付款币种'].map(c.to_iso_currency)          # [主表]付款币种->ISO
    out['报账金额（支付币种）'] = pay.map(c.amt2)                           # [明细] 付款金额
    return out


def run():
    # 1. 读源表 + 过滤主表
    m = pd.read_excel(SRC_M)
    d = pd.read_excel(SRC_D)
    mf = c.filter_main(m, SOURCES)

    # 2. 构建映射(数据库 + 规则表)
    gh_map  = c.build_gonghao_map()
    ven_map = c.build_vendor_map()
    ent_map = c.build_entity_map(code_col='新主体编码')
    sub_map = c.build_subject_map()

    # 3. 主子按 ID 合并 + 构建输出
    df = d[d['ID'].isin(set(mf['ID']))].merge(mf, on='ID', suffixes=('_d', ''), how='inner')
    out = build_output(df, gh_map, ven_map, ent_map, sub_map)
    print('合并前明细行数:', len(out))

    # 4. 分组去重 + 填充率
    out, collapsed = c.dedup_rows(out, DEDUP_KEY)
    c.report_fill(out, ['申请人工号', '收款方编码', '核算主体编号', '费用项目编码'])

    # 5. 写模版
    c.write_to_template(out, TMPL, OUT, TMPL_SHEET)
    print('已写出:', OUT)

    # 6. 未匹配 / 待核对清单
    sup  = set(mf['供应商-文本'].dropna().astype(str).str.strip())
    comp = set(mf['公司主体'].dropna().astype(str).str.strip())
    subjs = set(df['预算科目'].dropna().astype(str).str.strip())
    emp  = set(mf['经办人'].dropna().astype(str).str.strip())
    sheets = {
        '未匹配_工号':     pd.DataFrame({'未匹配_经办人(工号)':   sorted(s for s in emp if c.nz(s) not in gh_map)}),
        '未匹配_供应商':   pd.DataFrame({'未匹配_供应商(收款方编码)': sorted(s for s in sup if c.nz(s) not in ven_map)}),
        '未匹配_核算主体': pd.DataFrame({'未匹配_公司主体(核算主体)': sorted(s for s in comp if c.nz(s) not in ent_map)}),
        '未匹配_费用科目': pd.DataFrame({'未匹配_预算科目(费用项目)': sorted(s for s in subjs if c.no_slash(s) not in sub_map)}),
        '分组合并_待核对': collapsed,
    }
    c.write_exceptions(EXC, sheets)
    print('已写出:', EXC, '| 各清单条数:', {k: len(v) for k, v in sheets.items()})


if __name__ == '__main__':
    run()
