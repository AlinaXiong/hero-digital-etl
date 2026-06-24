# -*- coding: utf-8 -*-
"""从飞书(corehr)拉取全量员工信息并导出 Excel。

拉取范围(取决于应用已开通的 corehr 权限):
  员工基本信息 / 雇佣状态 / 人员类型 / 公司 / 部门 / 直属上级 / 工作地点 /
  法定姓名·常用名·花名 / 邮箱·手机号·证件号 / 银行账号 / 自定义字段(职级、发薪类型等)。

跑法:  python -m etl.export_feishu_employees
凭据:  .env 里的 FEISHU_APP_ID / FEISHU_APP_SECRET
输出:  output/feishu_employees/飞书员工信息_YYYYMMDD.xlsx
"""
import json
import sys
from pathlib import Path

import pandas as pd

if __package__ is None or __package__ == '':
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from etl import common as c
from etl import feishu

OUTPUT_DIR = c.OUT_DIR / 'feishu_employees'
OUTPUT_FILE = OUTPUT_DIR / f'飞书员工信息_{c.today_suffix()}.xlsx'


def _name_by_type(name_list, enum_name):
    for name in name_list or []:
        if (name.get('name_type') or {}).get('enum_name') == enum_name:
            return name.get('display_name_local_script') or name.get('display_name_local_and_western_script') or ''
    return ''


def _parse_custom_value(raw):
    """自定义字段 value 是 JSON 字符串: 标量直接取值, 选项取中文。"""
    text = (raw or '').strip()
    if not text:
        return ''
    try:
        parsed = json.loads(text)
    except (ValueError, TypeError):
        return text
    if isinstance(parsed, dict):
        return parsed.get('zh-CN') or parsed.get('zh_cn') or parsed.get('value') or ''
    if isinstance(parsed, list):
        return '、'.join(str(x) for x in parsed)
    return '' if parsed is None else str(parsed)


def _bank_join(bank_list, key):
    return '、'.join(str(b.get(key, '')) for b in (bank_list or []) if b.get(key))


def build_rows(employees, companies, departments, locations, emp_types):
    # 直属上级: direct_manager_id 是 employment_id, 用员工自身建名字映射
    manager_name = {}
    for emp in employees:
        eid = emp.get('employment_id_v2') or emp.get('employment_id')
        if eid:
            manager_name[eid] = (emp.get('person_info') or {}).get('legal_name', '')

    rows = []
    custom_labels = []
    seen_labels = set()
    for emp in employees:
        person = emp.get('person_info') or {}
        bank = person.get('bank_account_list') or []
        row = {
            'userId': emp.get('employment_id_v2') or emp.get('employment_id', ''),
            '工号': emp.get('employee_number', ''),
            '法定姓名': person.get('legal_name', ''),
            '常用姓名': _name_by_type(person.get('name_list'), 'preferred_name'),
            '花名': _name_by_type(person.get('name_list'), 'additional_name'),
            '雇佣状态': feishu.zh((emp.get('employment_status') or {}).get('display')),
            '雇佣类型': feishu.zh((emp.get('employment_type') or {}).get('display')),
            '人员类型': emp_types.get(emp.get('employee_type_id', ''), ''),
            '公司': companies.get(emp.get('company_id', ''), ''),
            '部门': departments.get(emp.get('department_id', ''), ''),
            '直属上级': manager_name.get(emp.get('direct_manager_id', ''), ''),
            '邮箱': emp.get('email_address', ''),
            '手机号': person.get('phone_number', ''),
            '身份证号': person.get('national_id_number', ''),
            '银行账号': _bank_join(bank, 'bank_account_number'),
            '银行名称': _bank_join(bank, 'bank_name'),
            '开户支行': _bank_join(bank, 'branch_name'),
            '开户人': _bank_join(bank, 'account_holder'),
            '工作地点': locations.get(emp.get('work_location_id', ''), ''),
        }
        for field in emp.get('custom_fields') or []:
            label = (field.get('name') or {}).get('zh_cn') or field.get('custom_api_name', '')
            if not label:
                continue
            if label not in seen_labels:
                seen_labels.add(label)
                custom_labels.append(label)
            row[label] = _parse_custom_value(field.get('value'))
        rows.append(row)

    base_cols = ['userId', '工号', '法定姓名', '常用姓名', '花名', '雇佣状态', '雇佣类型',
                 '人员类型', '公司', '部门', '直属上级', '邮箱', '手机号', '身份证号',
                 '银行账号', '银行名称', '开户支行', '开户人', '工作地点']
    return rows, base_cols + custom_labels


def run():
    print('[飞书员工导出] 拉取维度数据(公司/地点/人员类型)...')
    companies = feishu.fetch_company_name_map()
    locations = feishu.fetch_location_name_map()
    emp_types = feishu.fetch_employee_type_name_map()
    print(f'  公司 {len(companies)} / 地点 {len(locations)} / 人员类型 {len(emp_types)}')

    print('[飞书员工导出] 拉取全量员工...')
    employees = feishu.fetch_all_employees()
    print(f'  员工 {len(employees)} 人')

    dept_ids = [e.get('department_id') for e in employees if e.get('department_id')]
    departments = feishu.fetch_department_name_map(dept_ids)
    print(f'  部门 {len(departments)}')

    rows, columns = build_rows(employees, companies, departments, locations, emp_types)
    df = pd.DataFrame(rows).reindex(columns=columns)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='员工列表', index=False)
        worksheet = writer.sheets['员工列表']
        for idx, _ in enumerate(df.columns, start=1):
            worksheet.column_dimensions[worksheet.cell(row=1, column=idx).column_letter].width = 22
    print(f'[飞书员工导出] 完成: {len(df)} 行 -> {OUTPUT_FILE}')


if __name__ == '__main__':
    run()
