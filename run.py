# -*- coding: utf-8 -*-
"""任务入口。用法:
    python run.py            # 默认跑 ap_payment_opening_extra_db
    python run.py ap_payment_opening_extra_db # 跑指定任务
    python run.py --list     # 列出所有任务

新增任务:在 etl/<分组> 下加一个 <任务名>.py,再在下面 TASKS 登记。
"""
import shutil
import sys
from datetime import datetime

from etl.util import common as c
from etl.lark import export_feishu_employees

from etl.contract import (
    anti_bribery_signers_db,
    contract_anchor_attachments_db,
    contract_anchor_db,
    contract_anti_bribery_attachments_db,
    contract_anti_bribery_db,
    contract_general_attachments_db,
    contract_general_db,
)
from etl.process import (
    ap_payment_opening_extra_db,
    ap_prepayment_opening_db,
    ar_invoice_opening_db,
)
from etl.invoice import invoice_info_db

# 登记任务:任务名 -> run 函数。新增任务在这里加一行。
TASKS = {
    'anti_bribery_signers_db': anti_bribery_signers_db.run,  # 反商业贿赂协议签署情况补登(DB直连版)
    'ap_payment_opening_extra_db': ap_payment_opening_extra_db.run,  # 应付期初 对公付款单/批量费用流程/只转入外部成本/MCN(DB直连版)
    'ap_prepayment_opening_db': ap_prepayment_opening_db.run,  # 预付期初 供应商预付款单/零工预付款单(DB直连版)
    'ar_invoice_opening_db': ar_invoice_opening_db.run,   # 应收期初 应收报账单(DB直连版)
    'contract_anti_bribery_db': contract_anti_bribery_db.run,  # 合同迁移 反商业贿赂协议(DB直连版)
    'contract_anti_bribery_attachments_db': contract_anti_bribery_attachments_db.run,  # 反商业贿赂协议附件下载(DB直连版)
    'contract_anchor_db': contract_anchor_db.run,         # 合同迁移 主播流程(DB直连版)
    'contract_anchor_attachments_db': contract_anchor_attachments_db.run,  # 合同迁移 主播流程附件下载(DB直连版)
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

# 合同类任务:数据/导入 Excel(不含附件下载)
CONTRACT_ALL_TASKS = (
    'contract_general_db',       # 一般流程 Excel
    'contract_anchor_db',        # 主播流程 Excel
    'contract_anti_bribery_db',  # 反商业贿赂协议 Excel
)

# 合同类任务:附件下载(每个流程一个;依赖 cookie,放在数据任务之后跑)
CONTRACT_ATTACHMENT_TASKS = (
    'contract_general_attachments_db',
    'contract_anchor_attachments_db',
    'contract_anti_bribery_attachments_db',
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


def _run_task_sequence(task_names, label):
    for index, task_name in enumerate(task_names, start=1):
        print(f'\n=== 运行 {label} 子任务 {index}/{len(task_names)}: {task_name} ===')
        TASKS[task_name]()


def run_all():
    started_at = datetime.now()
    _run_task_sequence(ALL_TASK_NAMES, 'all')
    create_all_summary(started_at)


def run_contract_all():
    """一次跑所有合同任务,不含附件下载。"""
    _run_task_sequence(CONTRACT_ALL_TASKS, 'contract_all')


def run_contract_all_with_attachments():
    """一次跑所有合同任务,含附件下载(先出全部导入 Excel,再下载附件)。"""
    _run_task_sequence(CONTRACT_ALL_TASKS + CONTRACT_ATTACHMENT_TASKS, 'contract_all_with_attachments')


TASKS['all'] = run_all
TASKS['contract_all'] = run_contract_all
TASKS['contract_all_with_attachments'] = run_contract_all_with_attachments


def main():
    arg = sys.argv[1] if len(sys.argv) > 1 else 'ap_payment_opening_extra_db'
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
