# -*- coding: utf-8 -*-
"""混合合同清单增补导入任务。

同一份业务清单里可能同时包含一般流程合同和主播流程合同。本任务只负责:

1. 读取清单并剔除已导入/重复编号;
2. 按 Excel 智书合同类型或泛微合同类型分流:主播类走 contract_anchor_db,其余走 contract_general_db;
3. 复用两个主任务的解析与导出 builder,生成一个混合导入 Excel:一般流程占前 9 个 sheet,
   主播流程占后 4 个 sheet。

运行方式::

    python run.py contract_mixed_add

可用环境变量覆盖输入文件::

    CONTRACT_MIXED_ADD_FILE
"""
from __future__ import annotations

import os
import re
import json
from copy import copy
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from etl.contract import contract_anchor_db as anchor
from etl.contract import contract_general_add as general_add
from etl.contract import contract_general_db as general
from etl.util import common as c


TASK_NAME = 'contract_mixed_add'
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
SOURCE_DIR = c.SRC_DIR / 'contract_general_add'
INPUT_FILE = Path(os.getenv(
    'CONTRACT_MIXED_ADD_FILE',
    SOURCE_DIR / '需技术侧优先导入的合同数据.xlsx',
))

DATE_SUFFIX = general.DATE_SUFFIX
GENERAL_OUTPUT_FILE = OUTPUT_DIR / f'智书合同字段_一般流程_混合增补_{DATE_SUFFIX}.xlsx'
ANCHOR_OUTPUT_FILE = OUTPUT_DIR / f'智书合同字段_主播流程_混合增补_{DATE_SUFFIX}.xlsx'
MIXED_OUTPUT_FILE = OUTPUT_DIR / f'智书合同字段_混合增补_{DATE_SUFFIX}.xlsx'
MIXED_ATTACHMENT_ROOT = OUTPUT_DIR / f'混合增补合同附件_{DATE_SUFFIX}'
ARCHIVED_REQUEST_FILE = OUTPUT_DIR / f'智书合同同步请求_归档_9_{DATE_SUFFIX}.json'
OTHER_REQUEST_FILE = OUTPUT_DIR / f'智书合同同步请求_其他_0_{DATE_SUFFIX}.json'
AUDIT_FILE = OUTPUT_DIR / f'混合增补处理清单_{DATE_SUFFIX}.xlsx'

# ZhiShuSynServiceImpl.SheetRole uses fixed workbook indexes instead of sheet names.
# Keep these slots synchronized with the Java enum: general 0-8, anchor 9-12.
JAVA_GENERAL_SHEET_NAMES = (
    '一般流程主表',
    '关联合同',
    '相关订单-订单信息',
    '采购申请',
    '订单信息明细',
    '对方信息',
    '我方主体列表',
    '付款计划',
    '收款计划',
)
JAVA_ANCHOR_SHEET_NAMES = (
    '主播流程主表',
    '主播流程_对方信息',
    '主播流程_我方信息',
    '主播流程_费用明细',
)
GENERAL_SOURCE_TO_JAVA_SHEET_NAMES = {
    general.SHEET_MAIN: JAVA_GENERAL_SHEET_NAMES[0],
    general.SHEET_RELATION: JAVA_GENERAL_SHEET_NAMES[1],
    general.SHEET_RELATED_ORDER: JAVA_GENERAL_SHEET_NAMES[2],
    general.SHEET_PURCHASE_REQUEST: JAVA_GENERAL_SHEET_NAMES[3],
    general.SHEET_ORDER_DETAIL: JAVA_GENERAL_SHEET_NAMES[4],
    general.SHEET_COUNTERPARTY: JAVA_GENERAL_SHEET_NAMES[5],
    general.SHEET_OUR_PARTY: JAVA_GENERAL_SHEET_NAMES[6],
    general.SHEET_PAYMENT_PLAN: JAVA_GENERAL_SHEET_NAMES[7],
    general.SHEET_COLLECTION_PLAN: JAVA_GENERAL_SHEET_NAMES[8],
}
ANCHOR_SOURCE_TO_JAVA_SHEET_NAMES = {
    anchor.SHEET_MAIN: JAVA_ANCHOR_SHEET_NAMES[0],
    anchor.SHEET_COUNTERPARTY: JAVA_ANCHOR_SHEET_NAMES[1],
    anchor.SHEET_OUR_PARTY: JAVA_ANCHOR_SHEET_NAMES[2],
    anchor.SHEET_FEE_DETAIL: JAVA_ANCHOR_SHEET_NAMES[3],
}

GENERAL_ADD_SOURCE_FILES = (
    general_add.MISSING_INPUT_FILE,
    general_add.AMOUNT_INPUT_FILE,
)

CONTRACT_COLUMN_CANDIDATES = (
    '合同编号', '合同号', '合同编码', 'contract_number（合同编码）', 'contract_number',
)
STATUS_COLUMN_CANDIDATES = ('导入状态', '状态', '是否导入', '导入标记')
TYPE_COLUMN_CANDIDATES = ('智书合同类型', '合同类型', 'contractCategory(智书框架合同类型)')
ORDER_COLUMN_CANDIDATES = ('关联业财订单', '订单编号')
OLD_CODE_COLUMN_CANDIDATES = ('老泛微编码', '泛微编码', 'OA编号')

EXCLUDED_CONTRACT_NUMBERS = frozenset({
    'H-DF2025070326',
    'H-DF2025080282',
    'H-DF2025090217',
    'H-DF2025100218',
    'H-DF2025110097',
    'H-DF2025110417',
    'H-DF2025110418',
    'H-DF2025120085',
    'H-DF2025120203',
    'H-DF2025120221',
    'H-DF2025120227',
    'H-DF2025120228',
    'H-DF2025120229',
    'H-DF2026030454',
    'H-DS2024120042',
    'H-KF2023040006',
    'H-KF2026030029',
    'H-KS2024080001',
    'H-OF2025100064',
    'H-OF2025100071',
    'H-OF2026010010',
    'H-OF2026010043',
    'H-OF2026010064',
    'H-OF2026010065',
    'H-OF2026020005',
    'H-OF2026020009',
    'H-OF2026030008',
    'H-OF2026030009',
    'H-OF2026030010',
    'H-OF2026030014',
    'H-OF2026030028',
    'H-OF2026030049',
    'H-OF2026030072',
    'H-OF2026040005',
    'H-OF2026040017',
    'H-OF2026040033',
    'H-OF2026040047',
    'H-OF2026050001',
    'H-OF2026050006',
    'H-OF2026060001',
    'H-OF2026060015',
    'H-OF2026060043',
    'H-OF2026060047',
    'H-S201804001',
    'H-S201905003',
    'H-S202003039-S01',
})


def _text(value):
    return general._text(value)


def _contract_key(value):
    return general._contract_number_key(value)


def _first_existing_column(df, names, required=False):
    normalized = {general._normalize_field_name(column): column for column in df.columns}
    for name in names:
        column = normalized.get(general._normalize_field_name(name))
        if column:
            return column
    if required:
        raise KeyError(f'输入表缺少列: {" / ".join(names)}; 实际列: {list(df.columns)}')
    return None


def _split_contract_numbers(value):
    text = _text(value)
    if not text:
        return []
    parts = re.split(r'[\r\n,，;；]+', text)
    return [part.strip() for part in parts if _contract_key(part.strip())]


def _read_mixed_input(path):
    if not path.exists():
        raise FileNotFoundError(f'混合增补输入文件不存在: {path}')

    # 输入清单的「创建人」仅是业务备注，禁止参与导入值映射。
    # 合同创建人必须由 general/anchor 原流程从泛微源数据解析并执行离职替换规则。
    sheets = pd.read_excel(path, sheet_name=None, dtype=object)
    rows = []
    skipped_sheets = []
    for sheet_name, raw in sheets.items():
        contract_col = _first_existing_column(raw, CONTRACT_COLUMN_CANDIDATES)
        if not contract_col:
            skipped_sheets.append(sheet_name)
            continue
        status_col = _first_existing_column(raw, STATUS_COLUMN_CANDIDATES)
        type_col = _first_existing_column(raw, TYPE_COLUMN_CANDIDATES)
        order_col = _first_existing_column(raw, ORDER_COLUMN_CANDIDATES)
        old_code_col = _first_existing_column(raw, OLD_CODE_COLUMN_CANDIDATES)

        for excel_index, row in raw.iterrows():
            raw_contract = _text(row.get(contract_col))
            for contract_number in _split_contract_numbers(raw_contract):
                rows.append({
                    '输入文件': path.name,
                    'Sheet': sheet_name,
                    'Excel行号': int(excel_index) + 2,
                    '源表合同编号': raw_contract,
                    '合同编号': contract_number,
                    '合同key': _contract_key(contract_number),
                    '智书合同类型': _text(row.get(type_col)) if type_col else '',
                    '关联业财订单': _text(row.get(order_col)) if order_col else '',
                    '老泛微编码': _text(row.get(old_code_col)) if old_code_col else '',
                    '导入状态': _text(row.get(status_col)) if status_col else '',
                })

    if not rows:
        detail = f'; 已跳过无合同编号列sheet: {", ".join(skipped_sheets)}' if skipped_sheets else ''
        raise RuntimeError(f'输入文件未读到任何合同编号: {path}{detail}')

    result = pd.DataFrame(rows)
    result = result[result['合同key'] != ''].copy()
    result['输入顺序'] = range(1, len(result) + 1)
    return result


def _load_general_add_input_keys():
    rows = []
    for path in GENERAL_ADD_SOURCE_FILES:
        if not path.exists():
            continue
        raw = pd.read_excel(path, dtype=object)
        contract_col = _first_existing_column(raw, ('合同号', '合同编号'), required=True)
        for excel_index, row in raw.iterrows():
            for contract_number in _split_contract_numbers(row.get(contract_col)):
                rows.append({
                    '来源文件': path.name,
                    'Excel行号': int(excel_index) + 2,
                    '合同编号': contract_number,
                    '合同key': _contract_key(contract_number),
                    '排除来源': 'contract_general_add.py输入清单',
                })
    return pd.DataFrame(rows)


def _load_general_add_processed_keys():
    direct_df = _load_general_add_input_keys()
    direct_keys = set(direct_df.get('合同key', pd.Series(dtype=object)))
    direct_keys.discard('')
    processed_keys = set(direct_keys)
    child_rows = []

    if direct_keys:
        try:
            catalog = general_add.load_contract_catalog()
            child_to_root = general_add.discover_children(catalog, direct_keys)
            for child_key, root_key in child_to_root.items():
                processed_keys.add(child_key)
                child_rows.append({
                    '来源文件': 'contract_general_add.py',
                    'Excel行号': '',
                    '合同编号': child_key,
                    '合同key': child_key,
                    '排除来源': f'contract_general_add.py子合同,父合同={root_key}',
                })
            for key in direct_keys:
                if key.startswith('H-KF'):
                    processed_keys.add(f'{key}_VIRTUAL')
        except Exception as error:
            print(f'[合同混合增补] contract_general_add子合同排除池读取失败,仅使用直接输入清单: {error}')

    extra_df = pd.DataFrame(child_rows)
    exclude_df = pd.concat([direct_df, extra_df], ignore_index=True) if not extra_df.empty else direct_df
    exclude_df = exclude_df.drop_duplicates('合同key', keep='first') if not exclude_df.empty else exclude_df
    return processed_keys, exclude_df


def _apply_input_exclusions(input_df, _general_add_keys):
    result = input_df.copy()
    reasons = [[] for _ in range(len(result))]
    marked = result['导入状态'].map(lambda value: '已导入' in _text(value))
    explicitly_excluded = result['合同key'].isin(EXCLUDED_CONTRACT_NUMBERS)
    for pos, is_marked in enumerate(marked):
        if is_marked:
            reasons[pos].append('表格导入状态标注已导入')
    for pos, is_excluded in enumerate(explicitly_excluded):
        if is_excluded:
            reasons[pos].append('用户指定不处理')

    result['排除原因'] = ['; '.join(items) for items in reasons]
    processable_mask = result['排除原因'].eq('')
    duplicate_mask = result.loc[processable_mask].duplicated('合同key', keep='first')
    duplicate_indexes = duplicate_mask[duplicate_mask].index
    result.loc[duplicate_indexes, '排除原因'] = '重复合同编号,本任务按合同编号仅生成一次'
    result['是否剔除'] = result['排除原因'].map(lambda value: 'Y' if _text(value) else 'N')
    return result


def _tuple_param(values):
    return tuple(dict.fromkeys(value for value in values if _text(value)))


def _query_contract_meta(contract_keys):
    keys = _tuple_param(contract_keys)
    if not keys:
        return {}
    frames = []
    for batch in general._chunked(list(keys), 500):
        df = c.query_db(
            'FW',
            'vspn_xtyy',
            'SELECT htbh AS `合同编号`, htlx AS `合同类型ID`, htzt AS `合同签署状态ID` '
            'FROM uf_htk WHERE htbh IN %(contract_codes)s',
            {'contract_codes': tuple(batch)},
        )
        frames.append(df)
    if not frames:
        return {}
    meta = pd.concat(frames, ignore_index=True)
    result = {}
    for _, row in meta.iterrows():
        key = _contract_key(row.get('合同编号'))
        result.setdefault(key, {
            '合同类型ID': c.format_code(row.get('合同类型ID')),
            '合同签署状态ID': c.format_code(row.get('合同签署状态ID')),
        })
    return result


def _is_excel_anchor_type(value):
    return '主播' in _text(value)


def _route_processable_rows(input_df):
    processable = input_df[input_df['排除原因'].eq('')].copy()
    if processable.empty:
        return processable
    meta = _query_contract_meta(processable['合同key'])
    processable['DB合同类型ID'] = processable['合同key'].map(
        lambda key: meta.get(key, {}).get('合同类型ID', ''))
    processable['DB合同签署状态ID'] = processable['合同key'].map(
        lambda key: meta.get(key, {}).get('合同签署状态ID', ''))
    processable['Excel类型含主播'] = processable['智书合同类型'].map(
        lambda value: 'Y' if _is_excel_anchor_type(value) else '')
    processable['DB类型为主播协议'] = processable['DB合同类型ID'].map(
        lambda value: 'Y' if c.format_code(value) == str(anchor.ANCHOR_CONTRACT_TYPE_CODE) else '')
    anchor_mask = processable['Excel类型含主播'].eq('Y') | processable['DB类型为主播协议'].eq('Y')
    processable['路由流程'] = anchor_mask.map({True: '主播流程', False: '一般流程'})
    processable['路由依据'] = processable.apply(_route_basis, axis=1)
    return processable


def _route_basis(row):
    basis = []
    if _text(row.get('Excel类型含主播')) == 'Y':
        basis.append('Excel智书合同类型含主播')
    if _text(row.get('DB类型为主播协议')) == 'Y':
        basis.append('泛微合同类型=主播协议')
    if basis:
        return '; '.join(basis)
    return '默认一般流程'


def _selected_sql_without_base_where(base_sql):
    marker = 'ORDER BY h.htbh, h.id'
    if marker not in base_sql:
        raise RuntimeError('源 SQL 结构已变化,找不到 ORDER BY 注入点')
    return base_sql.replace(marker, 'WHERE h.htbh IN %(contract_codes)s\n' + marker)


def _selected_anchor_sql():
    marker = 'ORDER BY h.htbh, h.id'
    if marker not in anchor.SOURCE_SQL:
        raise RuntimeError('主播源 SQL 结构已变化,找不到 ORDER BY 注入点')
    return anchor.SOURCE_SQL.replace(marker, '  AND h.htbh IN %(contract_codes)s\n' + marker)


def _query_selected_source(base_sql, contract_keys, source_label):
    keys = _tuple_param(contract_keys)
    if not keys:
        return pd.DataFrame()
    sql = _selected_sql_without_base_where(base_sql)
    frames = []
    for batch in general._chunked(list(keys), 500):
        frame = c.query_db('FW', 'vspn_xtyy', sql, {'contract_codes': tuple(batch)})
        if not frame.empty:
            frame['数据来源'] = source_label
            frame['强制追加导出'] = ''
            frames.append(frame)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _query_selected_anchor_source(contract_keys):
    keys = _tuple_param(contract_keys)
    if not keys:
        return pd.DataFrame()
    sql = _selected_anchor_sql()
    frames = []
    for batch in general._chunked(list(keys), 500):
        frame = c.query_db(
            'FW',
            'vspn_xtyy',
            sql,
            {
                'anchor_contract_type_code': anchor.ANCHOR_CONTRACT_TYPE_CODE,
                'migration_status_codes': anchor.MIGRATION_STATUS_CODES,
                'contract_codes': tuple(batch),
            },
        )
        if not frame.empty:
            frames.append(frame)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _copy_attrs(target, source):
    target.attrs = dict(source.attrs)
    return target


def _resolve_general_sources(contract_keys):
    mcn = _query_selected_source(general.SOURCE_SQL_ALL, contract_keys, '泛微(MCN)')
    if not mcn.empty:
        mcn = mcn[mcn['合同类型ID'].map(c.format_code) != str(anchor.ANCHOR_CONTRACT_TYPE_CODE)].copy()
    event = _query_selected_source(general.SOURCE_SQL_HTSP_ALL, contract_keys, '泛微(赛事)')
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
        return resolved, mcn, event

    general._merge_attrs(resolved, [resolved_mcn, resolved_event])
    resolved = general._apply_supplement_amount_rollup(resolved)
    general._merge_attrs(resolved, [resolved_mcn, resolved_event])
    event_ids = resolved.loc[
        resolved.get('数据来源', pd.Series('', index=resolved.index)).map(_text).eq('泛微(赛事)'),
        'ID',
    ] if 'ID' in resolved.columns else pd.Series(dtype=object)
    resolved.attrs['saishi_plan_map'] = (
        general.load_htsp_plan_detail_map(event_ids) if not event_ids.empty else {}
    )
    return resolved, mcn, event


def _resolve_anchor_sources(contract_keys):
    raw = _query_selected_anchor_source(contract_keys)
    if raw.empty:
        return pd.DataFrame(), raw
    resolved = anchor.resolve_source_values(raw)
    return resolved, raw


def _timestamped_path(path):
    return path.with_name(f'{path.stem}_{datetime.now().strftime("%H%M%S")}{path.suffix}')


def _write_template_sheets_with_fallback(template_file, output_file, sheet_to_df, extra_sheets=None):
    try:
        return c.write_template_sheets(template_file, output_file, sheet_to_df, extra_sheets=extra_sheets)
    except PermissionError:
        fallback = _timestamped_path(output_file)
        print(f'输出文件被占用,改写到: {fallback}')
        return c.write_template_sheets(template_file, fallback, sheet_to_df, extra_sheets=extra_sheets)


def _ensure_placeholder_sheets(workbook, sheet_names):
    for sheet_name in sheet_names:
        if sheet_name not in workbook.sheetnames:
            workbook.create_sheet(sheet_name)


def _move_sheet_to_index(workbook, sheet_name, target_index):
    worksheet = workbook[sheet_name]
    current_index = workbook.index(worksheet)
    if current_index != target_index:
        workbook.move_sheet(worksheet, offset=target_index - current_index)


def _rename_general_sheets_for_sync(workbook):
    for source_name, target_name in GENERAL_SOURCE_TO_JAVA_SHEET_NAMES.items():
        if target_name not in workbook.sheetnames and source_name in workbook.sheetnames:
            workbook[source_name].title = target_name


def _rename_anchor_sheets_for_sync(workbook):
    for source_name, target_name in ANCHOR_SOURCE_TO_JAVA_SHEET_NAMES.items():
        if target_name not in workbook.sheetnames and source_name in workbook.sheetnames:
            workbook[source_name].title = target_name


def _align_sheet_order_for_zhishu_sync(output_file, flow_name):
    """Align workbook slots with ZhiShuSynServiceImpl.SheetRole indexes."""
    path = Path(output_file)
    workbook = load_workbook(path)
    if flow_name == '一般流程':
        _rename_general_sheets_for_sync(workbook)
        required_names = JAVA_GENERAL_SHEET_NAMES
        placeholder_names = JAVA_ANCHOR_SHEET_NAMES
        ordered_names = required_names + placeholder_names
        active_index = 0
    elif flow_name == '主播流程':
        _rename_anchor_sheets_for_sync(workbook)
        required_names = JAVA_ANCHOR_SHEET_NAMES
        placeholder_names = JAVA_GENERAL_SHEET_NAMES
        ordered_names = placeholder_names + required_names
        active_index = 9
    elif flow_name == '混合流程':
        _rename_general_sheets_for_sync(workbook)
        _rename_anchor_sheets_for_sync(workbook)
        required_names = JAVA_GENERAL_SHEET_NAMES + JAVA_ANCHOR_SHEET_NAMES
        placeholder_names = ()
        ordered_names = required_names
        active_index = 0
    else:
        raise ValueError(f'不支持的智书同步流程: {flow_name}')

    missing_names = [name for name in required_names if name not in workbook.sheetnames]
    if missing_names:
        raise KeyError(f'{flow_name}输出缺少必需sheet: {missing_names}')

    _ensure_placeholder_sheets(workbook, placeholder_names)
    for target_index, sheet_name in enumerate(ordered_names):
        _move_sheet_to_index(workbook, sheet_name, target_index)
    for worksheet in list(workbook.worksheets):
        if worksheet.title not in ordered_names:
            workbook.remove(worksheet)
    workbook.active = active_index
    workbook.save(path)
    print(f'[合同混合增补] 已按Java SheetRole顺序整理{flow_name}文件: {path}')
    return path


def _build_general_sheet_frames(source_df):
    headers = general._template_headers()
    if source_df.empty:
        return headers, {
            sheet_name: pd.DataFrame(columns=headers[sheet_name])
            for sheet_name in GENERAL_SOURCE_TO_JAVA_SHEET_NAMES
        }

    relation_df, _ = general.build_relation_output(source_df, headers[general.SHEET_RELATION])
    order_detail_df, _ = general.build_order_detail_output(source_df, headers[general.SHEET_ORDER_DETAIL])
    counterparty_df, _ = general.build_counterparty_output(source_df, headers[general.SHEET_COUNTERPARTY])
    our_party_df, _ = general.build_our_party_output(source_df, headers[general.SHEET_OUR_PARTY])
    return headers, {
        general.SHEET_MAIN: general.build_main_output(source_df, headers[general.SHEET_MAIN]),
        general.SHEET_RELATION: relation_df,
        general.SHEET_RELATED_ORDER: general.build_related_order_output(
            source_df, headers[general.SHEET_RELATED_ORDER]),
        general.SHEET_PURCHASE_REQUEST: general.build_purchase_request_output(
            source_df, headers[general.SHEET_PURCHASE_REQUEST]),
        general.SHEET_ORDER_DETAIL: order_detail_df,
        general.SHEET_COUNTERPARTY: counterparty_df,
        general.SHEET_OUR_PARTY: our_party_df,
        general.SHEET_PAYMENT_PLAN: general.build_payment_plan_output(
            source_df, headers[general.SHEET_PAYMENT_PLAN]),
        general.SHEET_COLLECTION_PLAN: general.build_collection_plan_output(
            source_df, headers[general.SHEET_COLLECTION_PLAN]),
    }


def _build_anchor_sheet_frames(source_df):
    headers = anchor._template_headers()
    if source_df.empty:
        return headers, {
            sheet_name: pd.DataFrame(columns=headers[sheet_name])
            for sheet_name in ANCHOR_SOURCE_TO_JAVA_SHEET_NAMES
        }

    counterparty_df, _ = anchor.build_counterparty_output(source_df, headers[anchor.SHEET_COUNTERPARTY])
    our_party_df, _ = anchor.build_our_party_output(source_df, headers[anchor.SHEET_OUR_PARTY])
    fee_detail_df, _ = anchor.build_fee_detail_output(source_df, headers[anchor.SHEET_FEE_DETAIL])
    return headers, {
        anchor.SHEET_MAIN: anchor.build_main_output(source_df, headers[anchor.SHEET_MAIN]),
        anchor.SHEET_COUNTERPARTY: counterparty_df,
        anchor.SHEET_OUR_PARTY: our_party_df,
        anchor.SHEET_FEE_DETAIL: fee_detail_df,
    }


def _build_attachment_audit(source_df, flow_module, headers):
    if source_df.empty:
        return pd.DataFrame(), pd.DataFrame()
    old_output_dir = flow_module.OUTPUT_DIR
    try:
        flow_module.OUTPUT_DIR = OUTPUT_DIR
        _, _, manifest_df, missing_df = flow_module.build_contract_attachment_output(
            source_df,
            headers[flow_module.SHEET_CONTRACT_ATTACHMENT],
            headers[flow_module.SHEET_OTHER_ATTACHMENT],
        )
    finally:
        flow_module.OUTPUT_DIR = old_output_dir
    return manifest_df, missing_df


def _apply_anchor_template_layout(output_file):
    """Copy the four anchor-sheet header layouts into the combined workbook."""
    workbook = load_workbook(output_file)
    template = load_workbook(anchor.TEMPLATE_FILE)
    for source_name, target_name in ANCHOR_SOURCE_TO_JAVA_SHEET_NAMES.items():
        source = template[source_name]
        target = workbook[target_name]
        target.freeze_panes = source.freeze_panes
        target.sheet_view.showGridLines = source.sheet_view.showGridLines
        target.sheet_format.defaultColWidth = source.sheet_format.defaultColWidth
        target.sheet_format.defaultRowHeight = source.sheet_format.defaultRowHeight

        for column_name, dimension in source.column_dimensions.items():
            target_dimension = target.column_dimensions[column_name]
            target_dimension.width = dimension.width
            target_dimension.hidden = dimension.hidden
            target_dimension.outlineLevel = dimension.outlineLevel
        if 1 in source.row_dimensions:
            target.row_dimensions[1].height = source.row_dimensions[1].height

        for source_cell in source[1]:
            target_cell = target.cell(row=1, column=source_cell.column)
            if source_cell.has_style:
                target_cell._style = copy(source_cell._style)
            if source_cell.hyperlink:
                target_cell._hyperlink = copy(source_cell.hyperlink)
            if source_cell.comment:
                target_cell.comment = copy(source_cell.comment)
    workbook.save(output_file)


def _write_mixed_workbook(general_source_df, anchor_source_df):
    general_headers, general_sheets = _build_general_sheet_frames(general_source_df)
    anchor_headers, anchor_sheets = _build_anchor_sheet_frames(anchor_source_df)
    general_manifest_df, general_missing_df = _build_attachment_audit(
        general_source_df, general, general_headers)
    anchor_manifest_df, anchor_missing_df = _build_attachment_audit(
        anchor_source_df, anchor, anchor_headers)

    anchor_java_sheets = {
        ANCHOR_SOURCE_TO_JAVA_SHEET_NAMES[source_name]: output_df
        for source_name, output_df in anchor_sheets.items()
    }
    path = _write_template_sheets_with_fallback(
        general.TEMPLATE_FILE,
        MIXED_OUTPUT_FILE,
        general_sheets,
        extra_sheets=anchor_java_sheets,
    )
    _apply_anchor_template_layout(path)
    path = _align_sheet_order_for_zhishu_sync(path, '混合流程')
    print(f'[合同混合增补] 已生成混合导入文件: {path}')
    return (
        path,
        general_manifest_df,
        general_missing_df,
        anchor_manifest_df,
        anchor_missing_df,
    )


def _contract_numbers_by_archive_status(general_source_df, anchor_source_df):
    status_by_contract = {}
    for source_df in (general_source_df, anchor_source_df):
        if source_df.empty:
            continue
        for _, row in source_df.iterrows():
            contract_number = _text(row.get('合同编号'))
            if contract_number and contract_number not in status_by_contract:
                status_by_contract[contract_number] = _text(row.get('合同审批状态'))
    archived = [
        contract_number
        for contract_number, status in status_by_contract.items()
        if status == '归档'
    ]
    other = [
        contract_number
        for contract_number, status in status_by_contract.items()
        if status != '归档'
    ]
    return archived, other


def _write_json_file(path, payload):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8') as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)
        file.write('\n')
    return path


def _write_sync_request_files(general_source_df, anchor_source_df, mixed_path):
    archived, other = _contract_numbers_by_archive_status(general_source_df, anchor_source_df)
    file_path = Path(mixed_path).resolve().as_posix()
    fallback_root = MIXED_ATTACHMENT_ROOT.resolve().as_posix()
    archived_path = _write_json_file(ARCHIVED_REQUEST_FILE, {
        'filePath': file_path,
        'contractNumbers': archived,
        'contractFileFallbackRoot': fallback_root,
        'contractStatusCode': '9',
        'threadCount': 5,
        'batchSize': 10,
    })
    other_path = _write_json_file(OTHER_REQUEST_FILE, {
        'filePath': file_path,
        'contractNumbers': other,
        'contractFileFallbackRoot': fallback_root,
        'contractStatusCode': '0',
        'threadCount': 5,
        'batchSize': 10,
    })
    print(
        '[合同混合增补] 已生成同步请求:',
        f'归档(9) {len(archived)} 个 -> {archived_path};',
        f'其他(0) {len(other)} 个 -> {other_path}',
    )
    return archived_path, other_path


def _flow_scope_summary(flow_name, source_df, input_df, output_file):
    return pd.DataFrame([
        {'项目': '流程', '值': flow_name},
        {'项目': '输出文件', '值': output_file.name},
        {'项目': '输入明细行数', '值': len(input_df)},
        {'项目': '输入合同数', '值': input_df['合同key'].nunique() if len(input_df) else 0},
        {'项目': '导入合同数', '值': len(source_df)},
        {'项目': '输入来源', '值': str(INPUT_FILE)},
    ])


def _write_general_workbook(source_df, input_df):
    if source_df.empty:
        print('[合同混合增补] 一般流程无可生成合同,跳过一般流程导入文件')
        return None, pd.DataFrame(), pd.DataFrame()

    headers = general._template_headers()
    main_df = general.build_main_output(source_df, headers[general.SHEET_MAIN])
    relation_df, _ = general.build_relation_output(source_df, headers[general.SHEET_RELATION])
    related_order_df = general.build_related_order_output(source_df, headers[general.SHEET_RELATED_ORDER])
    purchase_request_df = general.build_purchase_request_output(source_df, headers[general.SHEET_PURCHASE_REQUEST])
    order_detail_df, _ = general.build_order_detail_output(source_df, headers[general.SHEET_ORDER_DETAIL])
    counterparty_df, _ = general.build_counterparty_output(source_df, headers[general.SHEET_COUNTERPARTY])
    our_party_df, _ = general.build_our_party_output(source_df, headers[general.SHEET_OUR_PARTY])
    payment_df = general.build_payment_plan_output(source_df, headers[general.SHEET_PAYMENT_PLAN])
    collection_df = general.build_collection_plan_output(source_df, headers[general.SHEET_COLLECTION_PLAN])

    old_output_dir = general.OUTPUT_DIR
    try:
        general.OUTPUT_DIR = OUTPUT_DIR
        contract_attachment_df, other_attachment_df, manifest_df, missing_df = general.build_contract_attachment_output(
            source_df,
            headers[general.SHEET_CONTRACT_ATTACHMENT],
            headers[general.SHEET_OTHER_ATTACHMENT],
        )
    finally:
        general.OUTPUT_DIR = old_output_dir

    path = general._write_template_sheets_with_fallback(
        general.TEMPLATE_FILE,
        GENERAL_OUTPUT_FILE,
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
        extra_sheets={
            '处理范围': _flow_scope_summary('一般流程', source_df, input_df, GENERAL_OUTPUT_FILE),
            '输入清单': input_df,
            '合同分类核对': general.build_category_audit_df(source_df),
            '订单映射核对': general.build_order_audit_df(source_df),
        },
    )
    path = _align_sheet_order_for_zhishu_sync(path, '一般流程')
    print(f'[合同混合增补] 已生成一般流程导入文件: {path}')
    return path, manifest_df, missing_df


def _write_anchor_workbook(source_df, input_df):
    if source_df.empty:
        print('[合同混合增补] 主播流程无可生成合同,跳过主播流程导入文件')
        return None, pd.DataFrame(), pd.DataFrame()

    headers = anchor._template_headers()
    main_df = anchor.build_main_output(source_df, headers[anchor.SHEET_MAIN])
    counterparty_df, _ = anchor.build_counterparty_output(source_df, headers[anchor.SHEET_COUNTERPARTY])
    our_party_df, _ = anchor.build_our_party_output(source_df, headers[anchor.SHEET_OUR_PARTY])
    fee_detail_df, _ = anchor.build_fee_detail_output(source_df, headers[anchor.SHEET_FEE_DETAIL])

    old_output_dir = anchor.OUTPUT_DIR
    try:
        anchor.OUTPUT_DIR = OUTPUT_DIR
        contract_attachment_df, other_attachment_df, manifest_df, missing_df = anchor.build_contract_attachment_output(
            source_df,
            headers[anchor.SHEET_CONTRACT_ATTACHMENT],
            headers[anchor.SHEET_OTHER_ATTACHMENT],
        )
    finally:
        anchor.OUTPUT_DIR = old_output_dir

    path = _write_template_sheets_with_fallback(
        anchor.TEMPLATE_FILE,
        ANCHOR_OUTPUT_FILE,
        {
            anchor.SHEET_MAIN: main_df,
            anchor.SHEET_COUNTERPARTY: counterparty_df,
            anchor.SHEET_OUR_PARTY: our_party_df,
            anchor.SHEET_FEE_DETAIL: fee_detail_df,
            anchor.SHEET_CONTRACT_ATTACHMENT: contract_attachment_df,
            anchor.SHEET_OTHER_ATTACHMENT: other_attachment_df,
        },
        extra_sheets={
            '处理范围': _flow_scope_summary('主播流程', source_df, input_df, ANCHOR_OUTPUT_FILE),
            '输入清单': input_df,
        },
    )
    anchor._add_flow_audit_sheet(path, source_df)
    anchor._add_platform_audit_sheet(path, source_df)
    path = _align_sheet_order_for_zhishu_sync(path, '主播流程')
    print(f'[合同混合增补] 已生成主播流程导入文件: {path}')
    return path, manifest_df, missing_df


def _source_keys(source_df):
    if source_df.empty or '合同编号' not in source_df.columns:
        return set()
    return set(source_df['合同编号'].map(_contract_key)) - {''}


def _build_unresolved_df(route_df, general_keys, anchor_keys, anchor_raw_keys):
    generated_keys = set(general_keys) | set(anchor_keys)
    rows = []
    for _, row in route_df.iterrows():
        key = row['合同key']
        if key in generated_keys:
            continue
        if row['路由流程'] == '主播流程':
            reason = (
                '主播流程解析后未保留,请看主播流程异常口径'
                if key in anchor_raw_keys else
                '主播源未找到或不满足主播流程SQL条件'
            )
        else:
            reason = '一般流程源未找到(uf_htk/uf_htsp均未命中)'
        rows.append({
            '合同编号': row['合同编号'],
            '合同key': key,
            '路由流程': row['路由流程'],
            '路由依据': row['路由依据'],
            '未生成原因': reason,
            'Excel行号': row['Excel行号'],
            '智书合同类型': row['智书合同类型'],
        })
    return pd.DataFrame(rows)


def _build_input_audit(input_df, route_df, general_keys, anchor_keys, unresolved_df):
    result = input_df.copy()
    route = route_df.drop_duplicates('合同key').set_index('合同key') if not route_df.empty else pd.DataFrame()
    unresolved = (
        unresolved_df.drop_duplicates('合同key').set_index('合同key')['未生成原因'].to_dict()
        if not unresolved_df.empty else {}
    )
    generated_keys = set(general_keys) | set(anchor_keys)
    for column in ('路由流程', '路由依据', 'DB合同类型ID', 'DB合同签署状态ID'):
        result[column] = result['合同key'].map(
            lambda key: _text(route.at[key, column]) if not route.empty and key in route.index else '')
    result['生成状态'] = result.apply(
        lambda row: _generation_status(row, generated_keys, unresolved),
        axis=1,
    )
    return result


def _generation_status(row, generated_keys, unresolved):
    if _text(row.get('排除原因')):
        return '未生成: ' + _text(row.get('排除原因'))
    key = row.get('合同key')
    if key in generated_keys:
        return '已生成'
    return '未生成: ' + unresolved.get(key, '未命中源数据')


def _write_audit_workbook(input_audit_df, route_df, unresolved_df, general_add_exclude_df,
                          general_manifest_df, general_missing_df, anchor_manifest_df, anchor_missing_df,
                          mixed_path):
    summary = pd.DataFrame([
        {'项目': '输入文件', '值': str(INPUT_FILE)},
        {'项目': '输入展开行数', '值': len(input_audit_df)},
        {'项目': '输入合同数', '值': input_audit_df['合同key'].nunique()},
        {'项目': '剔除行数', '值': int((input_audit_df['是否剔除'] == 'Y').sum())},
        {'项目': '待处理合同数', '值': route_df['合同key'].nunique() if len(route_df) else 0},
        {'项目': '一般流程合同数', '值': int((route_df.get('路由流程', pd.Series(dtype=str)) == '一般流程').sum())},
        {'项目': '主播流程合同数', '值': int((route_df.get('路由流程', pd.Series(dtype=str)) == '主播流程').sum())},
        {'项目': '混合流程输出', '值': str(mixed_path or '')},
    ])
    sheets = {
        '处理汇总': summary,
        '输入清单_处理结果': input_audit_df,
        '待处理_路由结果': route_df,
        '未生成清单': unresolved_df,
        'contract_general_add历史范围': general_add_exclude_df,
        '一般流程合同附件下载清单': general_manifest_df,
        '一般流程合同附件DOCID缺失': general_missing_df,
        '主播流程合同附件下载清单': anchor_manifest_df,
        '主播流程合同附件DOCID缺失': anchor_missing_df,
    }
    try:
        path = c.write_exceptions(AUDIT_FILE, sheets)
    except PermissionError:
        fallback = _timestamped_path(AUDIT_FILE)
        print(f'处理清单被占用,改写到: {fallback}')
        path = c.write_exceptions(fallback, sheets)
    if path:
        print(f'[合同混合增补] 已生成处理清单: {path}')
    return Path(path) if path else None


def run():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    input_df = _read_mixed_input(INPUT_FILE)
    general_add_keys, general_add_exclude_df = _load_general_add_processed_keys()
    input_df = _apply_input_exclusions(input_df, general_add_keys)
    route_df = _route_processable_rows(input_df)

    general_keys = route_df.loc[route_df['路由流程'].eq('一般流程'), '合同key'] if not route_df.empty else []
    anchor_keys = route_df.loc[route_df['路由流程'].eq('主播流程'), '合同key'] if not route_df.empty else []

    print(
        '[合同混合增补] 输入:',
        f'{len(input_df)} 行 / {input_df["合同key"].nunique()} 个合同;',
        f'剔除 {(input_df["是否剔除"] == "Y").sum()} 行;',
        f'一般流程 {len(list(general_keys))} 个;',
        f'主播流程 {len(list(anchor_keys))} 个',
    )

    general_source_df, general_mcn_raw, general_event_raw = _resolve_general_sources(general_keys)
    anchor_source_df, anchor_raw_df = _resolve_anchor_sources(anchor_keys)

    (
        mixed_path,
        general_manifest_df,
        general_missing_df,
        anchor_manifest_df,
        anchor_missing_df,
    ) = _write_mixed_workbook(general_source_df, anchor_source_df)
    archived_request_path, other_request_path = _write_sync_request_files(
        general_source_df,
        anchor_source_df,
        mixed_path,
    )

    general_output_keys = _source_keys(general_source_df)
    anchor_output_keys = _source_keys(anchor_source_df)
    anchor_raw_keys = _source_keys(anchor_raw_df)
    unresolved_df = _build_unresolved_df(route_df, general_output_keys, anchor_output_keys, anchor_raw_keys)
    input_audit_df = _build_input_audit(
        input_df,
        route_df,
        general_output_keys,
        anchor_output_keys,
        unresolved_df,
    )

    audit_path = _write_audit_workbook(
        input_audit_df,
        route_df,
        unresolved_df,
        general_add_exclude_df,
        general_manifest_df,
        general_missing_df,
        anchor_manifest_df,
        anchor_missing_df,
        mixed_path,
    )

    return [
        path
        for path in (
            mixed_path,
            audit_path,
            archived_request_path,
            other_request_path,
        )
        if path
    ]


if __name__ == '__main__':
    run()
