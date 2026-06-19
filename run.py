# -*- coding: utf-8 -*-
"""任务入口。用法:
    python run.py            # 默认跑 ap_payment_opening
    python run.py ap_payment_opening # 跑指定任务
    python run.py --list     # 列出所有任务

新增任务:在 etl/tasks 下加一个 <任务名>.py,再在下面 TASKS 登记。
"""
import sys

from etl.tasks import (
    ap_payment_opening,
    ap_payment_opening_db,
    ap_payment_opening_extra_db,
    ap_prepayment_opening,
    ap_prepayment_opening_db,
    ar_invoice_opening,
    ar_invoice_opening_db,
    invoice_info_db,
)

# 登记任务:任务名 -> run 函数。新增任务在这里加一行。
TASKS = {
    'ap_payment_opening': ap_payment_opening.run,         # 应付期初 对公付款单
    'ap_payment_opening_db': ap_payment_opening_db.run,   # 应付期初 对公付款单(DB直连版)
    'ap_payment_opening_extra_db': ap_payment_opening_extra_db.run,  # 应付期初 批量费用流程/只转入外部成本(DB直连版)
    'ap_prepayment_opening': ap_prepayment_opening.run,   # 预付期初 供应商预付款单
    'ap_prepayment_opening_db': ap_prepayment_opening_db.run,  # 预付期初 供应商预付款单/零工预付款单(DB直连版)
    'ar_invoice_opening': ar_invoice_opening.run,         # 应收期初 应收报账单
    'ar_invoice_opening_db': ar_invoice_opening_db.run,   # 应收期初 应收报账单(DB直连版)
    'invoice_info_db': invoice_info_db.run,               # 发票信息(DB直连版)
}


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
