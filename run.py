# -*- coding: utf-8 -*-
"""任务入口。用法:
    python run.py            # 默认跑 ap_payment_opening
    python run.py ap_payment_opening # 跑指定任务
    python run.py --list     # 列出所有任务

新增任务:在 etl/tasks 下加一个 <任务名>.py,再在下面 TASKS 登记。
"""
import shutil
import sys
from datetime import datetime

from etl import common as c
from etl import export_feishu_employees

from etl.tasks import (
    anti_bribery_signers_db,
    ap_payment_opening,
    ap_payment_opening_db,
    ap_payment_opening_extra_db,
    ap_prepayment_opening,
    ap_prepayment_opening_db,
    ar_invoice_opening,
    ar_invoice_opening_db,
    contract_anchor_db,
    contract_general_attachments_db,
    contract_general_db,
    invoice_info_db,
)

# 登记任务:任务名 -> run 函数。新增任务在这里加一行。
TASKS = {
    'anti_bribery_signers_db': anti_bribery_signers_db.run,  # 反商业贿赂协议签署情况补登(DB直连版)
    'ap_payment_opening': ap_payment_opening.run,         # 应付期初 对公付款单
    'ap_payment_opening_db': ap_payment_opening_db.run,   # 应付期初 对公付款单(DB直连版)
    'ap_payment_opening_extra_db': ap_payment_opening_extra_db.run,  # 应付期初 批量费用流程/只转入外部成本(DB直连版)
    'ap_prepayment_opening': ap_prepayment_opening.run,   # 预付期初 供应商预付款单
    'ap_prepayment_opening_db': ap_prepayment_opening_db.run,  # 预付期初 供应商预付款单/零工预付款单(DB直连版)
    'ar_invoice_opening': ar_invoice_opening.run,         # 应收期初 应收报账单
    'ar_invoice_opening_db': ar_invoice_opening_db.run,   # 应收期初 应收报账单(DB直连版)
    'contract_anchor_db': contract_anchor_db.run,         # 合同迁移 主播流程(DB直连版)
    'contract_general_attachments_db': contract_general_attachments_db.run,  # 合同迁移 一般流程附件下载(DB直连版)
    'contract_general_db': contract_general_db.run,       # 合同迁移 一般流程(DB直连版)
    'invoice_info_db': invoice_info_db.run,               # 发票信息(DB直连版)
    'export_feishu_employees': export_feishu_employees.run,  # 飞书全量员工信息导出Excel
}

ALL_TASK_NAMES = (
    'ap_payment_opening_extra_db',
    'ap_prepayment_opening_db',
    'ar_invoice_opening_db',
    'invoice_info_db',
)

ALL_SUMMARY_GROUPS = (
    ('应付期初', 'ap_payment_opening_extra_db'),
    ('预付期初', 'ap_prepayment_opening_db'),
    ('应收期初', 'ar_invoice_opening_db'),
    ('发票信息', 'invoice_info_db'),
)


def _summary_dir_name():
    return f'应收-应付-预付-发票-{c.today_suffix()[4:]}汇总'


def _copy_recent_task_outputs(task_name, target_dir, started_at):
    source_dir = c.OUT_DIR / task_name
    if not source_dir.exists():
        return 0

    copied_count = 0
    threshold = started_at.timestamp() - 2
    for source_file in sorted(source_dir.glob('*.xlsx'), key=lambda path: path.name):
        if source_file.name.startswith('~$'):
            continue
        if source_file.stat().st_mtime < threshold:
            continue
        target_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_file, target_dir / source_file.name)
        copied_count += 1
    return copied_count


def create_all_summary(started_at):
    summary_dir = c.OUT_DIR / _summary_dir_name()
    if summary_dir.exists():
        try:
            shutil.rmtree(summary_dir)
        except PermissionError:
            summary_dir = summary_dir.with_name(f'{summary_dir.name}_{datetime.now().strftime("%H%M%S")}')
    summary_dir.mkdir(parents=True, exist_ok=True)

    total_count = 0
    for folder_name, task_name in ALL_SUMMARY_GROUPS:
        copied_count = _copy_recent_task_outputs(task_name, summary_dir / folder_name, started_at)
        total_count += copied_count
        print(f'[all汇总] {folder_name}: {copied_count} 个文件')
    print(f'[all汇总] 已生成: {summary_dir} (共 {total_count} 个文件)')
    return summary_dir


def run_all():
    started_at = datetime.now()
    for index, task_name in enumerate(ALL_TASK_NAMES, start=1):
        print(f'\n=== 运行 all 子任务 {index}/{len(ALL_TASK_NAMES)}: {task_name} ===')
        TASKS[task_name]()
    create_all_summary(started_at)


TASKS['all'] = run_all


def main():
    arg = sys.argv[1] if len(sys.argv) > 1 else 'ap_payment_opening'
    if arg in ('--list', '-l'):
        print('可用任务:', ', '.join(TASKS))
        return
    if arg not in TASKS:
        print(f'未知任务: {arg}\n可用任务: {", ".join(TASKS)}')
        sys.exit(1)
    print(f'=== 运行任务: {arg} ===')
    TASKS[arg]()


if __name__ == '__main__':
    main()
