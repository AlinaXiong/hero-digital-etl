# -*- coding: utf-8 -*-
"""合同迁移 - 主播流程合同附件下载(DB直连版)。

只下载附件并输出下载清单,不生成/改写智书导入 Excel。
跑法: 在项目根执行 python run.py contract_anchor_attachments_db
"""
import os

from etl.contract import contract_anchor_db as anchor
from etl.contract import contract_general_db as base


TASK_NAME = 'contract_anchor_attachments_db'


def run():
    source_df = anchor.read_source()
    manifest_df, missing_df = anchor.build_contract_attachment_manifest(source_df)
    download_root = anchor._anchor_attachment_download_root()
    cookie = os.getenv(base.ATTACHMENT_COOKIE_ENV, '').strip()

    if manifest_df.empty:
        print('[主播流程合同附件] 没有可下载附件。')
        output_file = anchor._write_exceptions_with_fallback(anchor.MANIFEST_FILE, {
            '合同附件下载清单': manifest_df,
            '合同附件DOCID_缺失映射': missing_df,
        })
        if output_file:
            print('已写出:', output_file)
        return

    manifest_df = manifest_df.copy()
    if not anchor._download_enabled(cookie):
        status = 'download_disabled' if os.getenv(base.ATTACHMENT_DOWNLOAD_ENABLED_ENV, '').strip().lower() in (
            '0', 'false', 'n', 'no', '否'
        ) else 'cookie_missing'
        manifest_df['status'] = status
        manifest_df['error'] = (
            f'未配置 {base.ATTACHMENT_COOKIE_ENV},仅生成下载清单'
            if status == 'cookie_missing'
            else '环境变量关闭附件下载'
        )
        print(f'[主播流程合同附件] 未下载: {status}; 下载清单 {len(manifest_df)} 条 -> {download_root}')
    else:
        print(f'[主播流程合同附件] 开始下载 {len(manifest_df)} 个文件 -> {download_root}')
        manifest_df = anchor._download_attachment_manifest_16_workers(manifest_df, cookie)

    output_file = anchor._write_exceptions_with_fallback(anchor.MANIFEST_FILE, {
        '合同附件下载清单': manifest_df,
        '合同附件DOCID_缺失映射': missing_df,
    })
    if output_file:
        print('已写出:', output_file)


if __name__ == '__main__':
    run()
