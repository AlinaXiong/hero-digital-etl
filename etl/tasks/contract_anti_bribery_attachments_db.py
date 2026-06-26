# -*- coding: utf-8 -*-
"""合同迁移 - 反商业贿赂协议附件下载(DB直连版)。"""
import os

from etl.tasks import contract_anti_bribery_db as anti
from etl.tasks import contract_general_db as base


TASK_NAME = 'contract_anti_bribery_attachments_db'


def _download_enabled(cookie):
    flag = os.getenv(base.ATTACHMENT_DOWNLOAD_ENABLED_ENV, '').strip().lower()
    if flag in ('0', 'false', 'n', 'no', '否'):
        return False
    return bool(base._text(cookie))


def _download_attachment_manifest_16_workers(manifest_df, cookie):
    return base.download_attachment_manifest_16_workers(
        manifest_df,
        cookie,
        log_prefix='反商业贿赂协议附件',
    )


def run():
    template_rows, contract_df, main_output_df, manifest_df, missing_df = anti.build_outputs()
    anti.write_outputs(main_output_df, contract_df, template_rows)

    cookie = os.getenv(base.ATTACHMENT_COOKIE_ENV, '').strip()
    if manifest_df.empty:
        print('[反商业贿赂协议附件] 没有可下载附件。')
    elif not _download_enabled(cookie):
        status = 'download_disabled' if os.getenv(base.ATTACHMENT_DOWNLOAD_ENABLED_ENV, '').strip().lower() in (
            '0', 'false', 'n', 'no', '否'
        ) else 'cookie_missing'
        manifest_df = manifest_df.copy()
        manifest_df['status'] = status
        manifest_df['error'] = (
            f'未配置 {base.ATTACHMENT_COOKIE_ENV},仅生成下载清单'
            if status == 'cookie_missing'
            else '环境变量关闭附件下载'
        )
        print(f'[反商业贿赂协议附件] 未下载: {status}; 仅生成下载清单')
    else:
        print(f'[反商业贿赂协议附件] 开始下载 {len(manifest_df)} 个文件 -> {anti._anti_attachment_download_root()}')
        manifest_df = _download_attachment_manifest_16_workers(manifest_df, cookie)

    anti.write_exception_outputs(template_rows, contract_df, main_output_df, manifest_df, missing_df)


if __name__ == '__main__':
    run()
