# -*- coding: utf-8 -*-
"""合同补登四类文件附件下载。

只处理 contract_general_add 生成的四个导入 Excel:

1. 缺少的合同
2. 缺少的合同对应的子合同
3. 金额缺失的合同
4. 金额缺失的合同对应的子合同

附件规则:

- _VIRTUAL 虚拟合同使用框架合同附件,但按虚拟合同编号落盘。
- 主文件、归档扫描件各最多 1 条;其他附件可多条。
- 目标稿件内若没有文件名以合同编号开头的主文件/归档扫描件,取该类型任意第一条兜底。
"""
from __future__ import annotations

import html
import os
import time
from pathlib import Path

import pandas as pd

from etl.contract import contract_anchor_db as anchor
from etl.contract import contract_general_add as general_add
from etl.contract import contract_general_db as base
from etl.util import common as c


TASK_NAME = 'contract_general_add_attachments_db'
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
DOWNLOAD_ROOT = OUTPUT_DIR / f'合同补登附件_{base.DATE_SUFFIX}'
MANIFEST_FILE = OUTPUT_DIR / f'合同补登四类附件下载清单_{base.DATE_SUFFIX}.xlsx'

CATEGORY_FILES = (
    ('缺少的合同', general_add.OUTPUT_MISSING),
    ('缺少的合同对应的子合同', general_add.OUTPUT_MISSING_CHILDREN),
    ('金额缺失的合同', general_add.OUTPUT_AMOUNT),
    ('金额缺失的合同对应的子合同', general_add.OUTPUT_AMOUNT_CHILDREN),
)


def _text(value):
    return base._text(value)


def _contract_key(value):
    return base._contract_number_key(value)


def _source_contract_number(contract_number):
    text = _text(contract_number)
    key = _contract_key(text)
    if key.endswith('_VIRTUAL'):
        return text[:-len('_VIRTUAL')]
    return text


def _read_category_contracts():
    rows = []
    for category, path in CATEGORY_FILES:
        if not path.exists():
            print(f'[合同补登附件] 输入文件不存在,跳过: {path}')
            continue
        df = pd.read_excel(path, sheet_name=base.SHEET_MAIN, dtype=object)
        contract_col = _first_contract_column(df)
        for excel_index, row in df.iterrows():
            contract_number = _text(row.get(contract_col))
            if not contract_number:
                continue
            source_contract = _source_contract_number(contract_number)
            rows.append({
                '类别': category,
                '输入文件': path.name,
                'Excel行号': int(excel_index) + 2,
                '合同编号': contract_number,
                '合同key': _contract_key(contract_number),
                '附件来源合同编号': source_contract,
                '附件来源合同key': _contract_key(source_contract),
                '是否虚拟合同': 'Y' if _contract_key(contract_number).endswith('_VIRTUAL') else '',
            })
    if not rows:
        return pd.DataFrame(columns=[
            '类别', '输入文件', 'Excel行号', '合同编号', '合同key',
            '附件来源合同编号', '附件来源合同key', '是否虚拟合同',
        ])
    result = pd.DataFrame(rows)
    result = result[result['合同key'] != ''].copy()
    return result.drop_duplicates(['类别', '合同key'], keep='first').reset_index(drop=True)


def _first_contract_column(df):
    candidates = (
        'contract_number（合同编码）',
        'contract_number(合同编码)',
        '合同编号',
        '合同号',
    )
    normalized = {base._normalize_field_name(column): column for column in df.columns}
    for candidate in candidates:
        column = normalized.get(base._normalize_field_name(candidate))
        if column:
            return column
    raise KeyError(f'字段模板缺少合同编码列; 实际列: {list(df.columns)}')


def _query_contract_type_map(contract_numbers):
    numbers = tuple(dict.fromkeys(_text(value) for value in contract_numbers if _text(value)))
    if not numbers:
        return {}
    frames = []
    for batch in base._chunked(list(numbers), 500):
        df = c.query_db(
            'FW',
            'vspn_xtyy',
            'SELECT htbh AS `合同编号`, htlx AS `合同类型ID` FROM uf_htk '
            'WHERE htbh IN %(contract_numbers)s',
            {'contract_numbers': tuple(batch)},
        )
        if not df.empty:
            frames.append(df)
    if not frames:
        return {}
    result = {}
    for _, row in pd.concat(frames, ignore_index=True).iterrows():
        result[_contract_key(row.get('合同编号'))] = c.format_code(row.get('合同类型ID'))
    return result


def _attach_route(input_df):
    if input_df.empty:
        input_df['附件流程'] = ''
        return input_df
    type_map = _query_contract_type_map(input_df['附件来源合同编号'])
    result = input_df.copy()
    result['附件来源合同类型ID'] = result['附件来源合同key'].map(lambda key: type_map.get(key, ''))
    result['附件流程'] = result['附件来源合同类型ID'].map(
        lambda value: '主播流程' if c.format_code(value) == str(anchor.ANCHOR_CONTRACT_TYPE_CODE) else '一般流程'
    )
    return result


def _selected_general_sources(contract_numbers):
    if not contract_numbers:
        return pd.DataFrame()
    source_df = general_add.read_and_resolve_sources(contract_numbers)
    if source_df.empty:
        return source_df
    source_df = source_df[
        source_df.get('合同类型ID', pd.Series('', index=source_df.index)).map(c.format_code)
        != str(anchor.ANCHOR_CONTRACT_TYPE_CODE)
    ].copy()
    source_df.attrs.update(source_df.attrs)
    return source_df


def _selected_anchor_sql():
    marker = 'ORDER BY h.htbh, h.id'
    if marker not in anchor.SOURCE_SQL:
        raise RuntimeError('主播源 SQL 结构已变化,找不到 ORDER BY 注入点')
    return anchor.SOURCE_SQL.replace(marker, '  AND h.htbh IN %(contract_numbers)s\n' + marker)


def _selected_anchor_sources(contract_numbers):
    numbers = tuple(dict.fromkeys(_text(value) for value in contract_numbers if _text(value)))
    if not numbers:
        return pd.DataFrame()
    sql = _selected_anchor_sql()
    frames = []
    for batch in base._chunked(list(numbers), 500):
        df = c.query_db(
            'FW',
            'vspn_xtyy',
            sql,
            {
                'anchor_contract_type_code': anchor.ANCHOR_CONTRACT_TYPE_CODE,
                'migration_status_codes': anchor.MIGRATION_STATUS_CODES,
                'contract_numbers': tuple(batch),
            },
        )
        if not df.empty:
            frames.append(df)
    if not frames:
        return pd.DataFrame()
    return anchor.resolve_source_values(pd.concat(frames, ignore_index=True))


def _subset_by_source_keys(source_df, source_keys):
    if source_df.empty:
        return source_df
    keys = {_contract_key(value) for value in source_keys if _contract_key(value)}
    subset = source_df[source_df['合同编号'].map(_contract_key).isin(keys)].copy()
    subset.attrs = dict(source_df.attrs)
    return subset


def _build_manifest_with_all_others(source_df, download_root):
    if source_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    docrefs_by_contract, raw_effective_docids, all_docids = base._timed(
        '  ├─附件docref收集(表单/赛事字段)',
        lambda: base._collect_attachment_docrefs(source_df),
    )
    print(f'[计时]   ├─附件 docid 总数: {len(all_docids)}', flush=True)
    docimage_map, imagefile_map, docdetail_map = base._timed(
        '  ├─附件三表JOIN取数(_load_attachment_maps)',
        lambda: base._load_attachment_maps(all_docids),
    )

    candidate_rows = []
    missing_rows = []
    seen = set()
    started_at = time.perf_counter()
    for source in source_df.to_dict('records'):
        contract_number = _text(source['合同编号'])
        fixed_docids = base.FIXED_CONTRACT_ATTACHMENT_DOCIDS.get(contract_number)
        contract_docrefs = docrefs_by_contract.get(contract_number, [])
        if fixed_docids:
            contract_docrefs = [item for item in contract_docrefs if item['docid'] in fixed_docids]
        for docref in contract_docrefs:
            docid = docref['docid']
            doc_rows = docimage_map.get(docid, [])
            if not doc_rows:
                missing_rows.append({
                    'contract_number（合同编码）': contract_number,
                    '合同ID': _text(source.get('ID')),
                    '合同名称': _text(source.get('合同标题')),
                    'raw_docids': raw_effective_docids.get(contract_number, ''),
                    'docid': docid,
                    'source': docref.get('source', ''),
                    'status': 'missing_docimagefile',
                    'error': 'docimagefile 无记录',
                })
                continue
            for doc_row in doc_rows:
                imagefileid = doc_row['imagefileid']
                dedupe_key = (contract_number, docid, imagefileid)
                if dedupe_key in seen:
                    continue
                seen.add(dedupe_key)
                image_info = imagefile_map.get(imagefileid, {})
                doc_info = docdetail_map.get(docid, {})
                attachment_name = html.unescape(base._first_non_blank(
                    doc_row.get('filename'),
                    image_info.get('filename'),
                    doc_info.get('subject'),
                    str(imagefileid),
                ))
                attachment_type = base._classify_attachment_type(
                    doc_row,
                    image_info,
                    doc_info,
                    docref.get('attachment_type', ''),
                )
                skip_reason = '' if fixed_docids else base._attachment_skip_reason(source, attachment_type)
                if skip_reason:
                    continue
                candidate_rows.append({
                    'contract_number（合同编码）': contract_number,
                    '合同ID': _text(source.get('ID')),
                    '合同名称': _text(source.get('合同标题')),
                    'attachment_type': attachment_type,
                    'raw_docids': raw_effective_docids.get(contract_number, ''),
                    'docid': docid,
                    'imagefileid': imagefileid,
                    'attachment_name': attachment_name,
                    'attachment_sheet': '',
                    'attachment_rule': '',
                    'imagefiletype': image_info.get('imagefiletype', ''),
                    'filesize': image_info.get('filesize', ''),
                    'target_path': '',
                    'source': docref.get('source', ''),
                    'nodeid': docref.get('nodeid', ''),
                    'share_id': docref.get('share_id', ''),
                    'doc_created_at': ' '.join(
                        item for item in [doc_info.get('created_date', ''), doc_info.get('created_time', '')] if item
                    ),
                    'status': 'pending',
                    'error': '',
                    '_source': source,
                })

    rows = _assign_main_archive_with_all_others(candidate_rows)
    rows = _set_target_paths(rows, download_root)
    print(f'[计时]   └─附件清单逐行构建({len(rows)}行): {time.perf_counter() - started_at:.1f}s', flush=True)
    return pd.DataFrame(rows), pd.DataFrame(missing_rows)


def _assign_main_archive_with_all_others(candidate_rows):
    grouped = {}
    for row in candidate_rows:
        grouped.setdefault(_text(row.get('contract_number（合同编码）')), []).append(row)

    assigned = []
    for contract_number, items in grouped.items():
        if not items:
            continue
        starts = [
            base._attachment_name_startswith_contract(item.get('attachment_name'), contract_number)
            for item in items
        ]
        start_indexes = [index for index, matched in enumerate(starts) if matched]
        main_index = start_indexes[0] if start_indexes else 0
        archived = base._is_archived_contract(items[main_index].get('_source', {}))

        main_item = items[main_index].copy()
        main_item['attachment_sheet'] = base.ATTACHMENT_FOLDER_MAIN
        main_item['attachment_rule'] = (
            '目标稿件且文件名以合同编号开头,保留1条作为主文件'
            if starts[main_index]
            else '目标稿件未命中合同编号前缀,取第一条作为主文件'
        )
        assigned.append(main_item)

        if archived:
            archive_item = items[main_index].copy()
            archive_item['attachment_sheet'] = base.ATTACHMENT_FOLDER_ARCHIVE_SCAN
            archive_item['attachment_rule'] = (
                '目标稿件且文件名以合同编号开头,保留1条作为归档扫描件'
                if starts[main_index]
                else '目标稿件未命中合同编号前缀,取第一条作为归档扫描件'
            )
            assigned.append(archive_item)

        for index, item in enumerate(items):
            if index == main_index:
                continue
            other_item = item.copy()
            other_item['attachment_sheet'] = base.SHEET_OTHER_ATTACHMENT
            other_item['attachment_rule'] = '目标稿件未作为主文件/归档扫描件,作为其他附件'
            assigned.append(other_item)

    for row in assigned:
        row.pop('_source', None)
    return assigned


def _set_target_paths(rows, download_root):
    used_paths = set()
    for row in rows:
        contract_number = _text(row.get('contract_number（合同编码）'))
        contract_dir = base._sanitize_path_part(contract_number, f'contract_{_text(row.get("合同ID"))}')
        folder_value = _text(row.get('attachment_sheet')) or base.SHEET_OTHER_ATTACHMENT
        target_dir = download_root / contract_dir / base._sanitize_path_part(folder_value, folder_value)
        target_name = base._build_target_filename(row.get('attachment_name'), row.get('imagefileid'))
        target_path = base._unique_attachment_path_preserve_name(
            target_dir,
            target_name,
            row.get('imagefileid'),
            used_paths,
        )
        row['target_path'] = str(target_path)
    return rows


def _remap_manifest_for_requests(source_manifest_df, source_missing_df, request_df, category_root):
    if request_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    manifest_rows = []
    missing_rows = []
    manifest_by_source = {}
    if not source_manifest_df.empty:
        for key, group in source_manifest_df.groupby(
            source_manifest_df['contract_number（合同编码）'].map(_contract_key),
            sort=False,
        ):
            manifest_by_source[key] = group
    missing_by_source = {}
    if not source_missing_df.empty:
        for key, group in source_missing_df.groupby(
            source_missing_df['contract_number（合同编码）'].map(_contract_key),
            sort=False,
        ):
            missing_by_source[key] = group

    for _, request in request_df.iterrows():
        request_number = _text(request['合同编号'])
        source_number = _text(request['附件来源合同编号'])
        source_key = _contract_key(source_number)
        for _, row in manifest_by_source.get(source_key, pd.DataFrame()).iterrows():
            item = row.to_dict()
            item['附件来源合同编号'] = source_number
            item['类别'] = request['类别']
            item['输入文件'] = request['输入文件']
            item['输入Excel行号'] = request['Excel行号']
            item['是否虚拟合同'] = request['是否虚拟合同']
            item['contract_number（合同编码）'] = request_number
            item['target_path'] = _remapped_target_path(item, category_root, request_number)
            manifest_rows.append(item)
        for _, row in missing_by_source.get(source_key, pd.DataFrame()).iterrows():
            item = row.to_dict()
            item['附件来源合同编号'] = source_number
            item['类别'] = request['类别']
            item['输入文件'] = request['输入文件']
            item['输入Excel行号'] = request['Excel行号']
            item['是否虚拟合同'] = request['是否虚拟合同']
            item['contract_number（合同编码）'] = request_number
            missing_rows.append(item)

    manifest_df = pd.DataFrame(manifest_rows)
    missing_df = pd.DataFrame(missing_rows)
    return manifest_df, missing_df


def _remapped_target_path(item, category_root, request_number):
    folder_value = _text(item.get('attachment_sheet')) or base.SHEET_OTHER_ATTACHMENT
    contract_dir = base._sanitize_path_part(request_number, f'contract_{_text(item.get("合同ID"))}')
    target_dir = category_root / contract_dir / base._sanitize_path_part(folder_value, folder_value)
    target_name = Path(_text(item.get('target_path'))).name
    if not target_name:
        target_name = base._build_target_filename(item.get('attachment_name'), item.get('imagefileid'))
    return str(target_dir / target_name)


def _download_enabled(cookie):
    flag = os.getenv(base.ATTACHMENT_DOWNLOAD_ENABLED_ENV, '').strip().lower()
    if flag in ('0', 'false', 'n', 'no', '否'):
        return False
    return bool(_text(cookie))


def _download_manifest(manifest_df, cookie, category):
    if manifest_df.empty:
        return manifest_df
    if not _download_enabled(cookie):
        status = 'download_disabled' if os.getenv(base.ATTACHMENT_DOWNLOAD_ENABLED_ENV, '').strip().lower() in (
            '0', 'false', 'n', 'no', '否'
        ) else 'cookie_missing'
        result = manifest_df.copy()
        result['status'] = status
        result['error'] = (
            f'未配置 {base.ATTACHMENT_COOKIE_ENV},仅生成下载清单'
            if status == 'cookie_missing'
            else '环境变量关闭附件下载'
        )
        print(f'[合同补登附件] {category}: 未下载 {len(result)} 条, status={status}')
        return result
    print(f'[合同补登附件] {category}: 开始下载 {len(manifest_df)} 个文件')
    return base.download_attachment_manifest_16_workers(
        manifest_df,
        cookie,
        log_prefix=f'合同补登附件-{category}',
    )


def _write_outputs(category_results, input_df):
    summary_rows = []
    sheets = {'输入合同范围': input_df}
    for category, manifest_df, missing_df in category_results:
        status_counts = (
            manifest_df['status'].value_counts(dropna=False).to_dict()
            if not manifest_df.empty and 'status' in manifest_df.columns else {}
        )
        summary_rows.append({
            '类别': category,
            '合同数': int(input_df[input_df['类别'] == category]['合同key'].nunique()),
            '附件数': len(manifest_df),
            '缺失DOCID数': len(missing_df),
            'downloaded': status_counts.get('downloaded', 0),
            'skipped_exists': status_counts.get('skipped_exists', 0),
            'failed': status_counts.get('failed', 0),
            'cookie_missing': status_counts.get('cookie_missing', 0),
            'download_disabled': status_counts.get('download_disabled', 0),
        })
        safe_name = category[:24]
        sheets[f'{safe_name}_下载清单'] = manifest_df
        sheets[f'{safe_name}_DOCID缺失'] = missing_df
    sheets = {'处理汇总': pd.DataFrame(summary_rows), **sheets}
    try:
        output_file = c.write_exceptions(MANIFEST_FILE, sheets)
    except PermissionError:
        output_file = c.write_exceptions(base._timestamped_path(MANIFEST_FILE), sheets)
    if output_file:
        print('已写出:', output_file)
    return output_file


def run():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    input_df = _attach_route(_read_category_contracts())
    if input_df.empty:
        print('[合同补登附件] 四个补登文件未读到合同编号')
        return None

    general_numbers = sorted(set(input_df.loc[input_df['附件流程'] == '一般流程', '附件来源合同编号'].map(_text)) - {''})
    anchor_numbers = sorted(set(input_df.loc[input_df['附件流程'] == '主播流程', '附件来源合同编号'].map(_text)) - {''})
    print(
        '[合同补登附件] 输入范围:',
        f'{len(input_df)} 行 / {input_df["合同key"].nunique()} 个合同;',
        f'一般来源 {len(general_numbers)} 个;',
        f'主播来源 {len(anchor_numbers)} 个',
    )

    general_source_df = _selected_general_sources(general_numbers)
    anchor_source_df = _selected_anchor_sources(anchor_numbers)
    cookie = os.getenv(base.ATTACHMENT_COOKIE_ENV, '').strip()

    category_results = []
    for category, _ in CATEGORY_FILES:
        category_df = input_df[input_df['类别'] == category].copy()
        if category_df.empty:
            category_results.append((category, pd.DataFrame(), pd.DataFrame()))
            continue
        category_root = DOWNLOAD_ROOT / category
        manifests = []
        missings = []
        for flow, source_df in (('一般流程', general_source_df), ('主播流程', anchor_source_df)):
            flow_requests = category_df[category_df['附件流程'] == flow].copy()
            if flow_requests.empty:
                continue
            source_subset = _subset_by_source_keys(source_df, flow_requests['附件来源合同key'])
            source_manifest_df, source_missing_df = _build_manifest_with_all_others(source_subset, category_root)
            manifest_df, missing_df = _remap_manifest_for_requests(
                source_manifest_df,
                source_missing_df,
                flow_requests,
                category_root,
            )
            if not manifest_df.empty:
                manifest_df['附件流程'] = flow
                manifests.append(manifest_df)
            if not missing_df.empty:
                missing_df['附件流程'] = flow
                missings.append(missing_df)
        category_manifest = pd.concat(manifests, ignore_index=True) if manifests else pd.DataFrame()
        category_missing = pd.concat(missings, ignore_index=True) if missings else pd.DataFrame()
        category_manifest = _download_manifest(category_manifest, cookie, category)
        category_results.append((category, category_manifest, category_missing))

    return _write_outputs(category_results, input_df)


if __name__ == '__main__':
    run()
