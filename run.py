# -*- coding: utf-8 -*-
"""任务入口。用法:
    python run.py            # 默认跑 ap_opening_payment
    python run.py ap_opening_payment # 跑指定任务
    python run.py --list     # 列出所有任务

新增任务:在 etl/tasks 下加一个 <任务名>.py,再在下面 TASKS 登记。
"""
import sys

from etl.tasks import ap_opening_payment

# 登记任务:任务名 -> run 函数。新增任务在这里加一行。
TASKS = {
    'ap_opening_payment': ap_opening_payment.run,   # 应付期初 对公付款单
}


def main():
    arg = sys.argv[1] if len(sys.argv) > 1 else 'ap_opening_payment'
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
