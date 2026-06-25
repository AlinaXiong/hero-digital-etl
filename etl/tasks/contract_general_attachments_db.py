# -*- coding: utf-8 -*-
"""合同迁移 - 一般流程合同附件下载(DB直连版)。

只下载附件并输出下载清单,不生成/改写智书导入 Excel。
跑法: 在项目根执行 python run.py contract_general_attachments_db
"""
import os

from etl.tasks import contract_general_db as base


TASK_NAME = 'contract_general_attachments_db'
MANIFEST_FILE = base.OUTPUT_DIR / f'一般流程合同附件下载清单_{base.DATE_SUFFIX}.xlsx'


def _download_enabled(cookie):
    flag = os.getenv(base.ATTACHMENT_DOWNLOAD_ENABLED_ENV, '').strip().lower()
    if flag in ('0', 'false', 'n', 'no', '否'):
        return False
    return bool(base._text(cookie))


def run():
    source_df = base.read_source()
    manifest_df, missing_df = base.build_contract_attachment_manifest(source_df)
    download_root = base._attachment_download_root()
    cookie = os.getenv(base.ATTACHMENT_COOKIE_ENV, '').strip()

    if manifest_df.empty:
        print('[一般流程合同附件] 没有可下载附件。')
        base._write_exceptions_with_fallback(MANIFEST_FILE, {
            '合同附件下载清单': manifest_df,
            '合同附件DOCID_缺失映射': missing_df,
        })
        return

    manifest_df = manifest_df.copy()
    if not _download_enabled(cookie):
        status = 'download_disabled' if os.getenv(base.ATTACHMENT_DOWNLOAD_ENABLED_ENV, '').strip().lower() in (
            '0', 'false', 'n', 'no', '否'
        ) else 'cookie_missing'
        manifest_df['status'] = status
        manifest_df['error'] = (
            f'未配置 {base.ATTACHMENT_COOKIE_ENV},仅生成下载清单'
            if status == 'cookie_missing'
            else '环境变量关闭附件下载'
        )
        print(f'[一般流程合同附件] 未下载: {status}; 下载清单 {len(manifest_df)} 条 -> {download_root}')
    else:
        print(f'[一般流程合同附件] 开始下载 {len(manifest_df)} 个文件 -> {download_root}')
        manifest_df = base.download_attachment_manifest(manifest_df, cookie, log_prefix='一般流程合同附件')

    output_file = base._write_exceptions_with_fallback(MANIFEST_FILE, {
        '合同附件下载清单': manifest_df,
        '合同附件DOCID_缺失映射': missing_df,
    })
    print('已写出:', output_file)

