# -*- coding: utf-8 -*-
"""混合增补导入文件附件下载。

只处理 contract_mixed_add 生成的混合导入范围，与
「智书合同字段_混合增补_*.xlsx」保持一致。

处理清单仅作为审计文件，不参与附件下载。本任务直接复用
contract_mixed_add.py 的输入读取、剔除、路由和 source 解析逻辑来确定下载范围。
"""
from __future__ import annotations

import os
from pathlib import Path

import pandas as pd

from etl.contract import contract_anchor_db as anchor
from etl.contract import contract_general_db as general
from etl.contract import contract_mixed_add as mixed
from etl.util import common as c


TASK_NAME = 'contract_mixed_add_attachments_db'
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
MANIFEST_FILE = OUTPUT_DIR / f'混合增补附件下载清单_{mixed.DATE_SUFFIX}.xlsx'

FLOW_GENERAL = '一般流程'
FLOW_ANCHOR = '主播流程'


def _text(value):
    return general._text(value)


def _contract_key(value):
    return general._contract_number_key(value)


def _resolve_scope_from_mixed_logic():
    input_df = mixed._read_mixed_input(mixed.INPUT_FILE)
    general_add_keys, general_add_exclude_df = mixed._load_general_add_processed_keys()
    input_df = mixed._apply_input_exclusions(input_df, general_add_keys)
    route_df = mixed._route_processable_rows(input_df)

    general_scope = (
        route_df[route_df['路由流程'].eq(FLOW_GENERAL)].copy()
        if not route_df.empty else pd.DataFrame()
    )
    anchor_scope = (
        route_df[route_df['路由流程'].eq(FLOW_ANCHOR)].copy()
        if not route_df.empty else pd.DataFrame()
    )
    return input_df, route_df, general_add_exclude_df, general_scope, anchor_scope


def _subset_by_scope(source_df, scope_df):
    if source_df.empty or scope_df.empty:
        return source_df.iloc[0:0].copy() if not source_df.empty else source_df
    keys = set(scope_df['合同key']) - {''}
    result = source_df[source_df['合同编号'].map(_contract_key).isin(keys)].copy()
    result.attrs = dict(source_df.attrs)
    return result


def _build_general_manifest(scope_df):
    if scope_df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    source_df, _, _ = mixed._resolve_general_sources(scope_df['合同key'])
    source_df = _subset_by_scope(source_df, scope_df)
    old_output_dir = general.OUTPUT_DIR
    try:
        general.OUTPUT_DIR = mixed.OUTPUT_DIR
        manifest_df, missing_df = general.build_contract_attachment_manifest(
            source_df,
            retention_mode='main_archive',
        )
    finally:
        general.OUTPUT_DIR = old_output_dir

    if not manifest_df.empty:
        manifest_df = manifest_df.copy()
        manifest_df['流程'] = FLOW_GENERAL
    if not missing_df.empty:
        missing_df = missing_df.copy()
        missing_df['流程'] = FLOW_GENERAL
    return source_df, manifest_df, missing_df


def _build_anchor_manifest(scope_df):
    if scope_df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    source_df, _ = mixed._resolve_anchor_sources(scope_df['合同key'])
    source_df = _subset_by_scope(source_df, scope_df)
    old_output_dir = anchor.OUTPUT_DIR
    try:
        anchor.OUTPUT_DIR = mixed.OUTPUT_DIR
        manifest_df, missing_df = anchor.build_contract_attachment_manifest(source_df)
    finally:
        anchor.OUTPUT_DIR = old_output_dir

    if not manifest_df.empty:
        manifest_df = manifest_df.copy()
        manifest_df['流程'] = FLOW_ANCHOR
    if not missing_df.empty:
        missing_df = missing_df.copy()
        missing_df['流程'] = FLOW_ANCHOR
    return source_df, manifest_df, missing_df


def _download_enabled(cookie):
    flag = os.getenv(general.ATTACHMENT_DOWNLOAD_ENABLED_ENV, '').strip().lower()
    if flag in ('0', 'false', 'n', 'no', '否'):
        return False
    return bool(_text(cookie))


def _retarget_manifest_to_mixed_root(manifest_df):
    if manifest_df.empty or 'target_path' not in manifest_df.columns:
        return manifest_df
    result = manifest_df.copy()

    def target_path(row):
        contract_number = _text(row.get('contract_number（合同编码）'))
        original_path = Path(_text(row.get('target_path')))
        if contract_number and contract_number in original_path.parts:
            contract_index = original_path.parts.index(contract_number)
            relative_parts = original_path.parts[contract_index:]
        else:
            relative_parts = (contract_number or '未识别合同', original_path.name)
        return str(mixed.MIXED_ATTACHMENT_ROOT.joinpath(*relative_parts))

    result['target_path'] = result.apply(target_path, axis=1)
    return result


def _disabled_status(cookie):
    flag = os.getenv(general.ATTACHMENT_DOWNLOAD_ENABLED_ENV, '').strip().lower()
    if flag in ('0', 'false', 'n', 'no', '否'):
        return 'download_disabled', '环境变量关闭附件下载'
    return 'cookie_missing', f'未配置 {general.ATTACHMENT_COOKIE_ENV},仅生成下载清单'


def _download_manifest(manifest_df, cookie, flow_name):
    if manifest_df.empty:
        return manifest_df
    if not _download_enabled(cookie):
        status, error = _disabled_status(cookie)
        result = manifest_df.copy()
        result['status'] = status
        result['error'] = error
        print(f'[混合增补附件] {flow_name}: 未下载 {len(result)} 条, status={status}')
        return result

    print(f'[混合增补附件] {flow_name}: 开始下载 {len(manifest_df)} 个文件')
    if flow_name == FLOW_ANCHOR:
        return anchor._download_attachment_manifest_16_workers(manifest_df, cookie)
    return general.download_attachment_manifest_16_workers(
        manifest_df,
        cookie,
        log_prefix='混合增补一般流程附件',
    )


def _status_summary(manifest_df):
    if manifest_df.empty or 'status' not in manifest_df.columns:
        return {}
    return manifest_df['status'].value_counts(dropna=False).to_dict()


def _build_summary(scope_df, general_source_df, anchor_source_df,
                   general_manifest_df, general_missing_df,
                   anchor_manifest_df, anchor_missing_df):
    rows = []
    flow_column = '流程' if '流程' in scope_df.columns else '路由流程'
    for flow_name, source_df, manifest_df, missing_df in (
        (FLOW_GENERAL, general_source_df, general_manifest_df, general_missing_df),
        (FLOW_ANCHOR, anchor_source_df, anchor_manifest_df, anchor_missing_df),
    ):
        flow_scope = (
            scope_df[scope_df[flow_column].eq(flow_name)]
            if not scope_df.empty and flow_column in scope_df.columns else pd.DataFrame()
        )
        status_counts = _status_summary(manifest_df)
        rows.append({
            '流程': flow_name,
            '导入文件合同数': int(flow_scope['合同key'].nunique()) if not flow_scope.empty else 0,
            '源数据合同数': len(source_df),
            '附件数': len(manifest_df),
            'DOCID缺失数': len(missing_df),
            'downloaded': status_counts.get('downloaded', 0),
            'skipped_exists': status_counts.get('skipped_exists', 0),
            'failed': status_counts.get('failed', 0),
            'cookie_missing': status_counts.get('cookie_missing', 0),
            'download_disabled': status_counts.get('download_disabled', 0),
        })
    rows.append({
        '流程': '处理清单',
        '导入文件合同数': '',
        '源数据合同数': '',
        '附件数': '',
        'DOCID缺失数': '',
        'downloaded': '',
        'skipped_exists': '',
        'failed': '',
        'cookie_missing': '',
        'download_disabled': '',
        '说明': f'{mixed.AUDIT_FILE.name} 仅用于审计,不下载附件',
    })
    return pd.DataFrame(rows)


def _write_outputs(input_df, route_df, general_add_exclude_df,
                   general_source_df, anchor_source_df,
                   general_manifest_df, general_missing_df,
                   anchor_manifest_df, anchor_missing_df):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    sheets = {
        '处理汇总': _build_summary(
            route_df,
            general_source_df,
            anchor_source_df,
            general_manifest_df,
            general_missing_df,
            anchor_manifest_df,
            anchor_missing_df,
        ),
        '输入清单_处理结果': input_df,
        '待处理_路由结果': route_df,
        'contract_general_add历史范围': general_add_exclude_df,
        '一般流程下载清单': general_manifest_df,
        '一般流程DOCID缺失': general_missing_df,
        '主播流程下载清单': anchor_manifest_df,
        '主播流程DOCID缺失': anchor_missing_df,
    }
    try:
        output_file = c.write_exceptions(MANIFEST_FILE, sheets)
    except PermissionError:
        output_file = c.write_exceptions(general._timestamped_path(MANIFEST_FILE), sheets)
    if output_file:
        print('已写出:', output_file)
    return output_file


def run(suppress_manifest=False):
    """下载附件；API 调用可跳过仅供审计的附件下载清单 Excel。"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    mixed.MIXED_ATTACHMENT_ROOT.mkdir(parents=True, exist_ok=True)
    input_df, route_df, general_add_exclude_df, general_scope, anchor_scope = _resolve_scope_from_mixed_logic()
    if route_df.empty:
        print('[混合增补附件] 混合增补输入经剔除后无待处理合同')
        return None

    print(
        '[混合增补附件] 输入范围:',
        f'{len(input_df)} 行 / {input_df["合同key"].nunique()} 个合同;',
        f'剔除 {(input_df["是否剔除"] == "Y").sum()} 行;',
        f'一般流程 {general_scope["合同key"].nunique() if len(general_scope) else 0} 个;',
        f'主播流程 {anchor_scope["合同key"].nunique() if len(anchor_scope) else 0} 个',
    )

    general_source_df, general_manifest_df, general_missing_df = _build_general_manifest(general_scope)
    anchor_source_df, anchor_manifest_df, anchor_missing_df = _build_anchor_manifest(anchor_scope)
    general_manifest_df = _retarget_manifest_to_mixed_root(general_manifest_df)
    anchor_manifest_df = _retarget_manifest_to_mixed_root(anchor_manifest_df)

    cookie = os.getenv(general.ATTACHMENT_COOKIE_ENV, '').strip()
    general_manifest_df = _download_manifest(general_manifest_df, cookie, FLOW_GENERAL)
    anchor_manifest_df = _download_manifest(anchor_manifest_df, cookie, FLOW_ANCHOR)

    if suppress_manifest:
        print('[混合增补附件] API 调用跳过附件下载清单 Excel')
        return None
    return _write_outputs(
        input_df,
        route_df,
        general_add_exclude_df,
        general_source_df,
        anchor_source_df,
        general_manifest_df,
        general_missing_df,
        anchor_manifest_df,
        anchor_missing_df,
    )


if __name__ == '__main__':
    run()
