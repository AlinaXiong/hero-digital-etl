# -*- coding: utf-8 -*-
"""合同迁移 —— 智书一般流程 special(DB 直连版)。

复用 contract_general_db 主逻辑,但关闭源数据与后续过滤。
在「字段模板」sheet 标出原一般流程任务会导入的合同。

跑法:在项目根执行  python run.py contract_general_special_db
"""
from etl.util import common as c
from etl.contract import contract_general_db as general


TASK_NAME = 'contract_general_special_db'
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
OUTPUT_FILE = OUTPUT_DIR / f'智书合同字段_一般流程_special_{general.DATE_SUFFIX}.xlsx'
EXCEPTION_FILE = OUTPUT_DIR / f'未匹配清单_一般流程_special_{general.DATE_SUFFIX}.xlsx'


def run():
    return general.run(
        task_label='合同迁移-一般流程-special',
        output_file=OUTPUT_FILE,
        exception_file=EXCEPTION_FILE,
        disable_filters=True,
        marker_in_main_sheet=True,
        write_exception_file=False,
    )


if __name__ == '__main__':
    run()
