# -*- coding: utf-8 -*-
"""按指定 Excel 合同清单补生成一般流程合同导入文件。

本任务不修改 ``contract_general_db`` 的取数、映射或导出逻辑，而是在独立入口中
复用其稳定的字段解析和各 sheet builder，并只覆盖本次需求明确要求的部分：

1. 普通合同的订单编号直接取输入 Excel；
2. H-KF 框架合同保留原合同类型、金额固定为 0、订单取输入 Excel；同时新增
   ``合同编号_VIRTUAL`` 虚拟订单合同，合同分类固定为
   ``订单支出-其他支出订单``，金额及付款计划为该框架合同报账金额合计；
3. H-KJ 暂不处理；
4. 付款计划合计不大于报账金额合计时仅写入错误清单；
5. 仅合同编号后缀子合同（如 ``-S01``、``-N``）单独输出；关联框架关系不算子合同。

运行方式::

    python run.py contract_general_add
    python -m etl.contract.contract_general_add

可用环境变量覆盖输入文件::

    CONTRACT_GENERAL_ADD_MISSING_FILE
    CONTRACT_GENERAL_ADD_AMOUNT_FILE
"""
from __future__ import annotations

import os
from collections import defaultdict
from pathlib import Path

import pandas as pd

from etl.contract import contract_general_db as general
from etl.util import common as c


TASK_NAME = 'contract_general_add'
OUTPUT_DIR = c.OUT_DIR / TASK_NAME

SOURCE_DIR = c.SRC_DIR / 'contract_general_add'
MISSING_INPUT_FILE = Path(os.getenv(
    'CONTRACT_GENERAL_ADD_MISSING_FILE',
    SOURCE_DIR / '缺少的合同.xlsx',
))
AMOUNT_INPUT_FILE = Path(os.getenv(
    'CONTRACT_GENERAL_ADD_AMOUNT_FILE',
    SOURCE_DIR / '金额缺失的合同.xlsx',
))

OUTPUT_MISSING = OUTPUT_DIR / '缺少的合同.xlsx'
OUTPUT_MISSING_CHILDREN = OUTPUT_DIR / '缺少的合同对应的子合同.xlsx'
OUTPUT_AMOUNT = OUTPUT_DIR / '金额缺失的合同.xlsx'
OUTPUT_AMOUNT_CHILDREN = OUTPUT_DIR / '金额缺失的合同对应的子合同.xlsx'

VIRTUAL_ORDER_CATEGORY = '订单支出-其他支出订单'
ERROR_COLUMNS = (
    '错误类型', '来源文件', 'Excel行号', '合同编号', '父合同编号',
    '报账金额合计', '付款计划金额合计', '差额', '说明',
)
CATALOG_COLUMNS = ('source_table', 'source_label', 'source_id', '合同编号', '关联框架协议ID')


def _text(value):
    return general._text(value)


def _contract_key(value):
    return general._contract_number_key(value)


def _is_hkf(value):
    return _contract_key(value).startswith('H-KF')


def _is_hkj(value):
    return _contract_key(value).startswith('H-KJ')


def _is_virtual_contract(value):
    return _contract_key(value).endswith('_VIRTUAL')


def _amount(value):
    return round(abs(general._number(value)), 2)


def _first_existing_column(df, names, required=False):
    normalized = {general._normalize_field_name(column): column for column in df.columns}
    for name in names:
        column = normalized.get(general._normalize_field_name(name))
        if column:
            return column
    if required:
        raise KeyError(f'输入表缺少列: {" / ".join(names)}; 实际列: {list(df.columns)}')
    return None


def read_input_file(path, dataset_name):
    """读取并标准化一份业务清单，保留 Excel 原始行序。"""
    if not path.exists():
        raise FileNotFoundError(f'输入文件不存在: {path}')
    raw = pd.read_excel(path, dtype=object)
    contract_col = _first_existing_column(raw, ('合同号', '合同编号'), required=True)
    order_col = _first_existing_column(raw, ('订单编号',), required=True)
    order_name_col = _first_existing_column(raw, ('订单名称',))
    report_col = _first_existing_column(
        raw,
        ('报账金额（支付币种）', '报账金额(支付币种)', '报账金额'),
        required=True,
    )
    payment_date_col = _first_existing_column(
        raw,
        ('付款时间', '支付日期', '报账日期', '日期'),
    )

    result = pd.DataFrame({
        '来源文件': path.name,
        '数据集': dataset_name,
        'Excel行号': range(2, len(raw) + 2),
        '合同编号': raw[contract_col].map(_text),
        '订单编号': raw[order_col].map(_text),
        '订单名称': raw[order_name_col].map(_text) if order_name_col else '',
        # 保留负数冲销行；框架虚拟合同金额按 Excel 各行带符号累加。
        '报账金额': pd.to_numeric(raw[report_col], errors='coerce').fillna(0).round(2),
        '付款时间': raw[payment_date_col].map(_text) if payment_date_col else '',
    })
    result['合同key'] = result['合同编号'].map(_contract_key)
    result = result[result['合同key'] != ''].copy()
    result['是否H-KF'] = result['合同key'].map(lambda value: value.startswith('H-KF'))
    result['是否H-KJ'] = result['合同key'].map(lambda value: value.startswith('H-KJ'))
    return result


def _catalog_query(table, source_label, relation_expression):
    df = c.query_db(
        'FW',
        'vspn_xtyy',
        f'SELECT id, htbh, {relation_expression} AS relation_id FROM {table} '
        'WHERE htbh IS NOT NULL AND TRIM(htbh) <> \'\'',
    )
    if df.empty:
        return pd.DataFrame(columns=CATALOG_COLUMNS)
    return pd.DataFrame({
        'source_table': table,
        'source_label': source_label,
        'source_id': df['id'].map(c.format_code),
        '合同编号': df['htbh'].map(_text),
        '关联框架协议ID': df['relation_id'].map(_text),
    })


def load_contract_catalog():
    """读取轻量合同目录，供合同编号后缀子合同发现使用。"""
    mcn = _catalog_query(general.FW_TABLE, '泛微(MCN)', 'glkjxy')
    event = _catalog_query(general.FW_TABLE_HTSP, '泛微(赛事)', 'COALESCE(glht, kjht)')
    catalog = pd.concat([mcn, event], ignore_index=True)
    catalog['合同key'] = catalog['合同编号'].map(_contract_key)
    catalog = catalog[catalog['合同key'] != ''].copy()
    catalog['关联ID列表'] = catalog['关联框架协议ID'].map(c.parse_browser_ids)
    print(f'[合同补登] 合同目录: MCN {len(mcn)} 行, 赛事 {len(event)} 行')
    return catalog


def discover_children(catalog, root_codes):
    """仅返回编号后缀子合同，以及每个子合同对应的输入根合同。"""
    roots = {_contract_key(code) for code in root_codes if _contract_key(code)}
    children = {}
    # general._main_contract_code_candidates 仅剥离 -S数字 / -N 后缀，支持递归后缀。
    for code in catalog['合同key'].drop_duplicates():
        ancestors = [_contract_key(item) for item in general._main_contract_code_candidates(code)]
        matched = next((ancestor for ancestor in ancestors if ancestor in roots), '')
        if matched and code not in roots:
            children.setdefault(code, matched)
    children = {code: root for code, root in children.items() if not _is_hkj(code)}
    return children


def _selected_sql(base_sql, contract_count):
    marker = 'ORDER BY h.htbh, h.id'
    if marker not in base_sql:
        raise RuntimeError('源 SQL 结构已变化，找不到 ORDER BY 注入点')
    predicate = f'WHERE h.htbh IN ({c.in_placeholders(range(contract_count))})\n'
    return base_sql.replace(marker, predicate + marker)


def _query_selected_source(base_sql, contract_codes, source_label):
    frames = []
    for batch in general._chunked(sorted(set(contract_codes)), 500):
        sql = _selected_sql(base_sql, len(batch))
        frame = c.query_db('FW', 'vspn_xtyy', sql, list(batch))
        if not frame.empty:
            frame['数据来源'] = source_label
            frame['强制追加导出'] = ''
            frames.append(frame)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def read_and_resolve_sources(contract_codes):
    """精确查询本任务所需合同并复用一般流程字段解析。"""
    codes = sorted({_text(code) for code in contract_codes if _text(code)})
    mcn = _query_selected_source(general.SOURCE_SQL_ALL, codes, '泛微(MCN)')
    event = _query_selected_source(general.SOURCE_SQL_HTSP_ALL, codes, '泛微(赛事)')

    # 和原一般流程一致：同编号优先 MCN。
    if not mcn.empty and not event.empty:
        mcn_keys = set(mcn['合同编号'].map(_contract_key))
        event = event[~event['合同编号'].map(_contract_key).isin(mcn_keys)].copy()

    resolved_mcn = (
        general.resolve_source_values(mcn, option_table=general.FW_TABLE)
        if not mcn.empty else pd.DataFrame()
    )
    resolved_event = (
        general.resolve_source_values(event, option_table=general.FW_TABLE_HTSP)
        if not event.empty else pd.DataFrame()
    )
    resolved = pd.concat([resolved_mcn, resolved_event], ignore_index=True)
    if resolved.empty:
        return resolved
    resolved['合同key'] = resolved['合同编号'].map(_contract_key)
    resolved = resolved.drop_duplicates('合同key', keep='first').copy()
    general._merge_attrs(resolved, [resolved_mcn, resolved_event])

    event_ids = resolved.loc[resolved['数据来源'].map(_text).eq('泛微(赛事)'), 'ID']
    resolved.attrs['saishi_plan_map'] = (
        general.load_htsp_plan_detail_map(event_ids) if not event_ids.empty else {}
    )
    print(f'[合同补登] 已解析合同 {len(resolved)} 个: MCN {len(resolved_mcn)}, 赛事 {len(resolved_event)}')
    return resolved


def _copy_attrs(target, source):
    target.attrs = dict(source.attrs)
    return target


def source_subset(resolved, contract_keys):
    keys = {_contract_key(code) for code in contract_keys if _contract_key(code)}
    if resolved.empty:
        return resolved.copy()
    subset = resolved[resolved['合同key'].isin(keys)].copy()
    _copy_attrs(subset, resolved)
    subset = general._apply_supplement_amount_rollup(subset)
    _copy_attrs(subset, resolved)
    return subset


def _existing_order_map(source):
    result = {}
    for item in general._order_entries_for_source(source):
        code = _text(item.get('订单编号'))
        if code:
            result[_contract_key(code)] = item
    return result


def _input_groups(input_df):
    return {key: group.sort_values('Excel行号') for key, group in input_df.groupby('合同key', sort=False)}


def apply_input_orders(source_df, order_input_df):
    """所有原合同（包括 H-KF 框架合同）的订单均直接取输入 Excel。"""
    if source_df.empty or order_input_df is None or order_input_df.empty:
        return source_df
    groups = _input_groups(order_input_df)
    saved_attrs = dict(source_df.attrs)
    source_df = source_df.copy()
    for index, source in source_df.iterrows():
        key = _contract_key(source.get('合同编号'))
        group = groups.get(key)
        if group is None or group.empty:
            continue

        existing = _existing_order_map(source)
        entries = []
        seen = set()
        for _, input_row in group.iterrows():
            order_code = _text(input_row.get('订单编号'))
            order_key = _contract_key(order_code)
            if not order_code or order_key in seen:
                continue
            seen.add(order_key)
            mapped = existing.get(order_key, {})
            entries.append({
                '订单编号': order_code,
                '订单名称': general._first_non_blank(
                    input_row.get('订单名称'), mapped.get('订单名称'),
                    source.get('项目名称'), source.get('合同标题'), order_code,
                ),
                '成本中心': general._first_non_blank(mapped.get('成本中心'), source.get('成本中心')),
                '订单开始日': general._first_non_blank(
                    mapped.get('订单开始日'), source.get('合同有效期起始时间'), source.get('合同签订日期')),
                '订单结束日': general._first_non_blank(
                    mapped.get('订单结束日'), source.get('合同有效期截止时间'), source.get('合同签订日期')),
            })
        source_df.at[index, '订单编号'] = ';'.join(item['订单编号'] for item in entries)
        source_df.at[index, '订单名称'] = ';'.join(item['订单名称'] for item in entries)
        source_df.at[index, '订单成本中心'] = ';'.join(item['成本中心'] for item in entries)
        source_df.at[index, '订单开始日'] = ';'.join(item['订单开始日'] for item in entries)
        source_df.at[index, '订单结束日'] = ';'.join(item['订单结束日'] for item in entries)
        source_df.at[index, '订单映射来源'] = f'{group.iloc[0]["来源文件"]}:订单编号'
    source_df.attrs = saved_attrs
    return source_df


def append_hkf_virtual_contracts(source_df, input_df):
    """保留 H-KF 框架合同，并为每个输入 H-KF 追加一条独立虚拟订单合同。"""
    if source_df.empty or input_df is None or input_df.empty:
        return source_df
    groups = _input_groups(input_df)
    saved_attrs = dict(source_df.attrs)
    result = source_df.copy()
    virtual_rows = []

    for index, source in result.iterrows():
        parent_key = _contract_key(source.get('合同编号'))
        group = groups.get(parent_key)
        if not (_is_hkf(parent_key) and group is not None and not group.empty):
            continue

        # 框架合同本身始终保持 0 金额；合同分类不修改，订单已由 apply_input_orders 取 Excel。
        for field in (
            '合同金额', '合同预计收入', '合同预计支出',
            '合同总额_解析', '收入总额_解析', '支出总额_解析',
            '合同总额_签名', '收入总额_签名', '支出总额_签名',
            '付款计划汇总金额', '收款计划汇总金额',
        ):
            if field in result.columns:
                result.at[index, field] = 0

        amount = round(group['报账金额'].sum(), 2)
        parent_number = _text(source.get('合同编号'))
        virtual_number = f'{parent_number}_VIRTUAL'
        virtual = source.to_dict()
        virtual.update({
            '合同编号': virtual_number,
            '合同key': _contract_key(virtual_number),
            '合同标题': f'{_text(source.get("合同标题")) or parent_number}-虚拟订单合同',
            '合同分类': VIRTUAL_ORDER_CATEGORY,
            '合同分类依据': (
                'contract_general_add:H-KF新增虚拟订单合同;固定分类=' + VIRTUAL_ORDER_CATEGORY
            ),
            '收支类型': '支出类',
            '合同金额': amount,
            '合同预计收入': 0,
            '合同预计支出': amount,
            '合同总额_解析': amount,
            '收入总额_解析': 0,
            '支出总额_解析': amount,
            '合同总额_签名': amount,
            '收入总额_签名': 0,
            '支出总额_签名': amount,
            '付款计划汇总金额': amount,
            '收款计划汇总金额': 0,
            '主合同编号': '',
            '是否补充协议': '',
            '金额汇总目标合同编号': virtual_number,
            '补充协议数量': 0,
            '补充协议编号': '',
            # Excel 订单属于原框架合同；虚拟合同不把 Excel 订单误当作合同编码。
            '订单编号': '',
            '订单名称': '',
            '订单成本中心': '',
            '订单开始日': '',
            '订单结束日': '',
            '订单映射来源': 'contract_general_add:H-KF虚拟订单合同',
            '关联合同信息': [{
                'id': _text(source.get('ID')),
                'number': parent_number,
                'title': _text(source.get('合同标题')),
                'source_table': _text(source.get('数据来源')),
            }],
            '关联合同编号': parent_number,
            '关联合同名称': _text(source.get('合同标题')),
            '关联框架协议ID': _text(source.get('ID')),
        })
        # 虚拟合同不重复携带框架合同附件/流程附件。
        for field in (
            '合同附件DOCID', '赛事初稿DOCID', '赛事签署稿DOCID', '赛事生效稿DOCID',
            '合同流程ID',
        ):
            if field in virtual:
                virtual[field] = ''
        virtual_rows.append(virtual)

    if virtual_rows:
        result = pd.concat([result, pd.DataFrame(virtual_rows)], ignore_index=True, sort=False)
    result.attrs = saved_attrs
    return result


def inherited_child_order_input(input_df, child_to_root):
    """普通父合同的子合同继承父合同 Excel 订单；H-KF 的 Excel 订单明确不继承。"""
    rows = []
    groups = _input_groups(input_df)
    for child_key, root_key in child_to_root.items():
        if _is_hkf(root_key):
            continue
        group = groups.get(root_key)
        if group is None:
            continue
        for _, row in group.iterrows():
            item = row.to_dict()
            item['合同编号'] = child_key
            item['合同key'] = child_key
            item['来源文件'] = f'{row["来源文件"]}:继承父合同{root_key}'
            item['是否H-KF'] = _is_hkf(child_key)
            item['是否H-KJ'] = _is_hkj(child_key)
            rows.append(item)
    return pd.DataFrame(rows, columns=input_df.columns) if rows else pd.DataFrame(columns=input_df.columns)


def _virtual_payment_date(source, group, existing_rows):
    if group is not None:
        first_input_date = next((_text(value) for value in group['付款时间'] if _text(value)), '')
        if first_input_date:
            return first_input_date
    if not existing_rows.empty:
        date_col = _first_existing_column(existing_rows, ('payment_plan_list[].payment_date（付款时间）',))
        if date_col:
            first_date = next((_text(value) for value in existing_rows[date_col] if _text(value)), '')
            if first_date:
                return first_date
    return general._first_non_blank(
        source.get('合同有效期截止时间'), source.get('合同签订日期'), source.get('合同创建日期'))


def build_payment_plan_output(source_df, headers, input_df=None):
    """复用原付款计划；H-KF 汇总付款计划仅挂到独立 _VIRTUAL 合同。"""
    base = general.build_payment_plan_output(source_df, headers)
    if input_df is None or input_df.empty or source_df.empty:
        return base

    groups = _input_groups(input_df)
    contract_col = _first_existing_column(base, ('contract_number（合同编码）',)) if len(base) else None
    hkf_keys = set(input_df.loc[input_df['是否H-KF'], '合同key'])
    virtual_keys = {f'{key}_VIRTUAL' for key in hkf_keys}
    kept = base
    if contract_col and hkf_keys:
        kept = base[~base[contract_col].map(_contract_key).isin(hkf_keys | virtual_keys)].copy()

    customer_info_map = source_df.attrs.get('customer_info_map', {})
    supplier_info_map = source_df.attrs.get('supplier_info_map', {})
    virtual_rows = []
    for source in source_df.to_dict('records'):
        key = _contract_key(source.get('合同编号'))
        if not _is_virtual_contract(key):
            continue
        parent_key = key[:-len('_VIRTUAL')]
        group = groups.get(parent_key)
        if group is None:
            continue
        merged_amount = round(group['报账金额'].sum(), 2)
        if not merged_amount:
            continue
        existing_rows = (
            base[base[contract_col].map(_contract_key).isin((parent_key, key))]
            if contract_col else pd.DataFrame()
        )
        row = general._new_row(headers)
        general._set(row, 'contract_number（合同编码）', _text(source['合同编号']))
        general._set(row, 'payment_plan_list（付款计划）', '付款计划')
        general._set(row, 'payment_plan_list[].payment_date（付款时间）', c.format_date(
            _virtual_payment_date(source, group, existing_rows)))
        general._set(row, 'payment_plan_list[].prepaid（是否预付）', general.DEFAULT_PREPAID)
        general._set(row, 'payment_plan_list[].payment_amount（付款金额）', merged_amount)
        general._set(row, 'payment_plan_list[].payment_desc（付款说明）', 'H-KF虚拟订单付款计划合并')
        general._set(
            row,
            'payment_plan_list[].payment_custom_attributes/custom_付款性质（付款性质）',
            general.DEFAULT_PAYMENT_NATURE,
        )
        general._set(
            row,
            'payment_plan_list[].payment_counter_party[].counter_party_code（付款对象）',
            general._first_counterparty_code(source, customer_info_map, supplier_info_map),
        )
        general._set(row, '付款计划行id(付款记录传的id)', general._plan_row_id(source, 'P', 'VIRTUAL'))
        virtual_rows.append(row)
    virtual = pd.DataFrame(virtual_rows, columns=headers)
    return pd.concat([kept, virtual], ignore_index=True)


def _error_row(error_type, source_file='', excel_row='', contract_number='', parent_number='',
               report_amount='', plan_amount='', description=''):
    difference = ''
    if report_amount != '' and plan_amount != '':
        difference = round(general._number(plan_amount) - general._number(report_amount), 2)
    return {
        '错误类型': error_type,
        '来源文件': source_file,
        'Excel行号': excel_row,
        '合同编号': contract_number,
        '父合同编号': parent_number,
        '报账金额合计': report_amount,
        '付款计划金额合计': plan_amount,
        '差额': difference,
        '说明': description,
    }


def input_scope_errors(input_df, found_keys):
    rows = []
    found = {_contract_key(value) for value in found_keys}
    for key, group in _input_groups(input_df).items():
        first = group.iloc[0]
        if _is_hkj(key):
            rows.append(_error_row(
                'H-KJ暂不处理', first['来源文件'], first['Excel行号'], first['合同编号'],
                report_amount=round(group['报账金额'].sum(), 2), description='按需求暂不生成导入记录'))
        elif key not in found:
            rows.append(_error_row(
                '数据库未找到合同', first['来源文件'], first['Excel行号'], first['合同编号'],
                report_amount=round(group['报账金额'].sum(), 2),
                description='uf_htk 与 uf_htsp 均未找到该合同编号'))
    return rows


def payment_plan_errors(input_df, payment_df):
    """严格按需求：付款计划合计 <= Excel 报账金额合计时记错，不修金额。"""
    if input_df is None or input_df.empty:
        return []
    contract_col = _first_existing_column(payment_df, ('contract_number（合同编码）',)) if len(payment_df) else None
    amount_col = _first_existing_column(
        payment_df,
        ('payment_plan_list[].payment_amount（付款金额）',),
    ) if len(payment_df) else None
    plan_sum = defaultdict(float)
    if contract_col and amount_col:
        for _, row in payment_df.iterrows():
            key = _contract_key(row[contract_col])
            # H-KF 的付款计划挂在新增虚拟合同；校验仍按输入框架合同汇总。
            if key.endswith('_VIRTUAL') and key[:-len('_VIRTUAL')] in set(input_df['合同key']):
                key = key[:-len('_VIRTUAL')]
            plan_sum[key] += general._number(row[amount_col])

    rows = []
    for key, group in _input_groups(input_df[~input_df['是否H-KJ']]).items():
        report_amount = round(group['报账金额'].sum(), 2)
        payment_amount = round(plan_sum.get(key, 0.0), 2)
        if payment_amount <= report_amount:
            first = group.iloc[0]
            rows.append(_error_row(
                '付款计划金额不足', first['来源文件'], first['Excel行号'], first['合同编号'],
                report_amount=report_amount, plan_amount=payment_amount,
                description='付款计划所有行合计不大于 Excel 报账金额合计；仅记录，不调整付款计划'))
    return rows


def _scope_summary(source_df, input_df, output_file, child_to_root=None):
    return pd.DataFrame([
        {'项目': '输出文件', '值': output_file.name},
        {'项目': '导入合同数', '值': len(source_df)},
        {'项目': 'Excel明细行数', '值': len(input_df) if input_df is not None else 0},
        {'项目': 'Excel合同数', '值': input_df['合同key'].nunique() if input_df is not None and len(input_df) else 0},
        {'项目': 'H-KF虚拟订单合同数', '值': int(source_df['合同编号'].map(_is_virtual_contract).sum()) if len(source_df) else 0},
        {'项目': '子合同数', '值': len(child_to_root or {})},
        {'项目': '付款计划校验规则', '值': '付款计划合计 <= Excel报账金额合计时写错误清单，不调整金额'},
    ])


def _child_relation_audit(source_df, child_to_root):
    rows = []
    for source in source_df.to_dict('records'):
        child = _contract_key(source.get('合同编号'))
        rows.append({
            '子合同编号': _text(source.get('合同编号')),
            '对应输入父合同编号': child_to_root.get(child, ''),
            '数据来源': _text(source.get('数据来源')),
            '关联框架合同编号': _text(source.get('关联合同编号')),
            '是否后缀子合同': 'Y' if general._main_contract_code_candidates(child) else 'N',
        })
    return pd.DataFrame(rows)


def build_one_workbook(source_df, output_file, order_input_df=None, audit_input_df=None,
                       initial_errors=None, child_to_root=None):
    headers = general._template_headers()
    source_df = apply_input_orders(source_df, order_input_df)
    source_df = append_hkf_virtual_contracts(source_df, audit_input_df)

    main_df = general.build_main_output(source_df, headers[general.SHEET_MAIN])
    relation_df, _ = general.build_relation_output(source_df, headers[general.SHEET_RELATION])
    related_order_df = general.build_related_order_output(source_df, headers[general.SHEET_RELATED_ORDER])
    purchase_request_df = general.build_purchase_request_output(source_df, headers[general.SHEET_PURCHASE_REQUEST])
    order_detail_df, _ = general.build_order_detail_output(source_df, headers[general.SHEET_ORDER_DETAIL])
    counterparty_df, _ = general.build_counterparty_output(source_df, headers[general.SHEET_COUNTERPARTY])
    our_party_df, _ = general.build_our_party_output(source_df, headers[general.SHEET_OUR_PARTY])
    payment_df = build_payment_plan_output(source_df, headers[general.SHEET_PAYMENT_PLAN], audit_input_df)
    collection_df = general.build_collection_plan_output(source_df, headers[general.SHEET_COLLECTION_PLAN])
    contract_attachment_df, other_attachment_df, _, _ = general.build_contract_attachment_output(
        source_df,
        headers[general.SHEET_CONTRACT_ATTACHMENT],
        headers[general.SHEET_OTHER_ATTACHMENT],
    )

    errors = list(initial_errors or [])
    errors.extend(payment_plan_errors(audit_input_df, payment_df))
    error_df = pd.DataFrame(errors, columns=ERROR_COLUMNS)
    extra_sheets = {
        '处理范围': _scope_summary(source_df, audit_input_df, output_file, child_to_root),
        '错误清单': error_df,
        '合同分类核对': general.build_category_audit_df(source_df),
        '订单映射核对': general.build_order_audit_df(source_df),
    }
    if child_to_root is not None:
        extra_sheets['子合同归属'] = _child_relation_audit(source_df, child_to_root)

    path = general._write_template_sheets_with_fallback(
        general.TEMPLATE_FILE,
        output_file,
        {
            general.SHEET_MAIN: main_df,
            general.SHEET_RELATION: relation_df,
            general.SHEET_RELATED_ORDER: related_order_df,
            general.SHEET_PURCHASE_REQUEST: purchase_request_df,
            general.SHEET_ORDER_DETAIL: order_detail_df,
            general.SHEET_COUNTERPARTY: counterparty_df,
            general.SHEET_OUR_PARTY: our_party_df,
            general.SHEET_PAYMENT_PLAN: payment_df,
            general.SHEET_COLLECTION_PLAN: collection_df,
            general.SHEET_CONTRACT_ATTACHMENT: contract_attachment_df,
            general.SHEET_OTHER_ATTACHMENT: other_attachment_df,
        },
        extra_sheets=extra_sheets,
    )
    print(
        f'[合同补登] 已生成 {path}: 合同 {len(source_df)}, 订单 {len(order_detail_df)}, '
        f'付款计划 {len(payment_df)}, 错误 {len(error_df)}')
    return Path(path)


def _dataset_plan(input_df, catalog):
    processable = input_df[~input_df['是否H-KJ']].copy()
    root_keys = list(dict.fromkeys(processable['合同key']))
    children = discover_children(catalog, root_keys)
    return processable, root_keys, children


def run():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    missing_input = read_input_file(MISSING_INPUT_FILE, '缺少的合同')
    amount_input = read_input_file(AMOUNT_INPUT_FILE, '金额缺失的合同')
    catalog = load_contract_catalog()

    missing_processable, missing_roots, missing_children = _dataset_plan(missing_input, catalog)
    amount_processable, amount_roots, amount_children = _dataset_plan(amount_input, catalog)
    all_codes = set(missing_roots) | set(amount_roots) | set(missing_children) | set(amount_children)
    resolved = read_and_resolve_sources(all_codes)
    found_keys = set(resolved.get('合同key', pd.Series(dtype=object)))

    missing_main = source_subset(resolved, missing_roots)
    amount_main = source_subset(resolved, amount_roots)
    missing_child = source_subset(resolved, missing_children)
    amount_child = source_subset(resolved, amount_children)

    missing_errors = input_scope_errors(missing_input, found_keys)
    amount_errors = input_scope_errors(amount_input, found_keys)

    missing_child_input = inherited_child_order_input(missing_processable, missing_children)
    amount_child_input = inherited_child_order_input(amount_processable, amount_children)

    outputs = [
        build_one_workbook(
            missing_main, OUTPUT_MISSING,
            order_input_df=missing_processable,
            audit_input_df=missing_input,
            initial_errors=missing_errors,
        ),
        build_one_workbook(
            missing_child, OUTPUT_MISSING_CHILDREN,
            order_input_df=missing_child_input,
            audit_input_df=None,
            child_to_root=missing_children,
        ),
        build_one_workbook(
            amount_main, OUTPUT_AMOUNT,
            order_input_df=amount_processable,
            audit_input_df=amount_input,
            initial_errors=amount_errors,
        ),
        build_one_workbook(
            amount_child, OUTPUT_AMOUNT_CHILDREN,
            order_input_df=amount_child_input,
            audit_input_df=None,
            child_to_root=amount_children,
        ),
    ]
    return outputs


if __name__ == '__main__':
    run()
