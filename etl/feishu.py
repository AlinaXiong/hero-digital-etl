# -*- coding: utf-8 -*-
"""飞书(corehr)公共类: 取员工/部门/公司/地点等信息。

凭据来自环境变量 FEISHU_APP_ID / FEISHU_APP_SECRET(可写入项目根 .env)。
- get_employee_status_maps(): 供合同"申请人状态(在职/离职)"判断;
- fetch_all_employees() + fetch_*_name_map(): 供全量员工信息导出。

飞书工号 employee_number 与泛微 hrmjobtitles.JOBTITLENAME / 汉得 employee_code
同口径(均为 V 编号)。
"""
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request

FEISHU_HOST = os.getenv('FEISHU_HOST', 'https://open.feishu.cn').rstrip('/')
_TOKEN_PATH = '/open-apis/auth/v3/tenant_access_token/internal'
_EMPLOYEE_SEARCH_PATH = '/open-apis/corehr/v2/employees/search'
_DEPARTMENT_BATCH_PATH = '/open-apis/corehr/v2/departments/batch_get'
_COMPANY_LIST_PATH = '/open-apis/corehr/v1/companies'
_LOCATION_LIST_PATH = '/open-apis/corehr/v1/locations'
_EMPLOYEE_TYPE_LIST_PATH = '/open-apis/corehr/v1/employee_types'
_MAX_RETRY = 4

_STATUS_MAPS_CACHE = None
_TOKEN_CACHE = None


def _request(method, path, params=None, payload=None, timeout=30, auth=True):
    """统一请求(GET/POST), 带瞬时网络/SSL 异常重试(代理环境偶发 SSL EOF)。"""
    url = FEISHU_HOST + path
    if params:
        url += '?' + urllib.parse.urlencode(params)
    data = json.dumps(payload).encode('utf-8') if payload is not None else None
    last_error = None
    for attempt in range(_MAX_RETRY):
        req = urllib.request.Request(url, data=data, method=method)
        if data is not None:
            req.add_header('Content-Type', 'application/json; charset=utf-8')
        if auth:
            req.add_header('Authorization', f'Bearer {_token()}')
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode('utf-8'))
        except (urllib.error.URLError, TimeoutError, ConnectionError) as error:
            last_error = error
            if attempt < _MAX_RETRY - 1:
                time.sleep(1.5 * (attempt + 1))
    raise last_error


def _post_json(path, payload, headers=None, timeout=30):
    """兼容旧调用: 直接 POST(headers 里通常已带 Authorization)。"""
    data = json.dumps(payload).encode('utf-8')
    last_error = None
    for attempt in range(_MAX_RETRY):
        req = urllib.request.Request(FEISHU_HOST + path, data=data, method='POST')
        req.add_header('Content-Type', 'application/json; charset=utf-8')
        for key, value in (headers or {}).items():
            req.add_header(key, value)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode('utf-8'))
        except (urllib.error.URLError, TimeoutError, ConnectionError) as error:
            last_error = error
            if attempt < _MAX_RETRY - 1:
                time.sleep(1.5 * (attempt + 1))
    raise last_error


def _get_tenant_access_token():
    app_id = os.environ.get('FEISHU_APP_ID', '').strip()
    app_secret = os.environ.get('FEISHU_APP_SECRET', '').strip()
    if not app_id or not app_secret:
        raise RuntimeError('缺少 FEISHU_APP_ID / FEISHU_APP_SECRET; 请在 .env 配置')
    result = _request('POST', _TOKEN_PATH, payload={'app_id': app_id, 'app_secret': app_secret}, auth=False)
    if result.get('code') != 0:
        raise RuntimeError(f"飞书获取 tenant_access_token 失败: {result.get('code')} {result.get('msg')}")
    return result['tenant_access_token']


def _token():
    global _TOKEN_CACHE
    if not _TOKEN_CACHE:
        _TOKEN_CACHE = _get_tenant_access_token()
    return _TOKEN_CACHE


def _check(result, tag):
    if result.get('code') != 0:
        raise RuntimeError(f"飞书 {tag} 失败: {result.get('code')} {result.get('msg')}")
    return result.get('data') or {}


def search_all(path, fields=None, extra_params=None):
    """分页 POST search, 逐条产出 items。"""
    page_token = ''
    while True:
        params = {'page_size': '100'}
        params.update(extra_params or {})
        if page_token:
            params['page_token'] = page_token
        payload = {'fields': fields} if fields is not None else {}
        data = _check(_request('POST', path, params=params, payload=payload), f'search {path}')
        for item in data.get('items') or []:
            yield item
        if not data.get('has_more'):
            break
        page_token = data.get('page_token') or ''
        if not page_token:
            break


def list_all(path, params=None):
    """分页 GET list, 逐条产出 items。"""
    page_token = ''
    while True:
        query = {'page_size': '100'}
        query.update(params or {})
        if page_token:
            query['page_token'] = page_token
        data = _check(_request('GET', path, params=query), f'list {path}')
        for item in data.get('items') or []:
            yield item
        if not data.get('has_more'):
            break
        page_token = data.get('page_token') or ''
        if not page_token:
            break


def zh(display_list):
    """从 [{lang,value}] 取中文(zh-CN), 兜底取第一个。"""
    items = display_list or []
    for item in items:
        if item.get('lang') == 'zh-CN' and item.get('value'):
            return item['value']
    return items[0].get('value', '') if items else ''


# ============================ 员工 ============================
EMPLOYEE_FIELDS = [
    'employee_number', 'employment_status', 'employment_type', 'employee_type_id',
    'email_address', 'company_id', 'department_id', 'direct_manager_id', 'work_location_id',
    'custom_fields',
    'person_info.legal_name', 'person_info.name_list', 'person_info.phone_number',
    'person_info.national_id_number', 'person_info.bank_account_list',
]


def fetch_all_employees(fields=None, user_id_type='user_id'):
    """拉全量员工记录(原始结构)。"""
    return list(search_all(_EMPLOYEE_SEARCH_PATH,
                           fields=fields if fields is not None else EMPLOYEE_FIELDS,
                           extra_params={'user_id_type': user_id_type,
                                         'department_id_type': 'open_department_id'}))


def fetch_company_name_map():
    return {item['id']: zh(((item.get('hiberarchy_common') or {}).get('name')))
            for item in list_all(_COMPANY_LIST_PATH) if item.get('id')}


def fetch_location_name_map():
    return {item['id']: zh(((item.get('hiberarchy_common') or {}).get('name')))
            for item in list_all(_LOCATION_LIST_PATH) if item.get('id')}


def fetch_employee_type_name_map():
    return {item['id']: zh(item.get('name')) for item in list_all(_EMPLOYEE_TYPE_LIST_PATH) if item.get('id')}


def fetch_department_name_map(dept_ids):
    """部门 open_department_id -> 中文名(batch_get, 每批<=100)。"""
    ids = [d for d in dict.fromkeys(dept_ids) if d]
    result = {}
    for start in range(0, len(ids), 100):
        batch = ids[start:start + 100]
        data = _check(_request('POST', _DEPARTMENT_BATCH_PATH,
                               params={'department_id_type': 'open_department_id'},
                               payload={'department_id_list': batch, 'fields': ['department_name']}),
                      'departments.batch_get')
        for item in data.get('items') or []:
            if item.get('id'):
                result[item['id']] = zh(item.get('department_name'))
    return result


# ============================ 在职/离职(合同申请人状态用) ============================
def _item_names(item):
    person = item.get('person_info') or {}
    names = set()
    legal = (person.get('legal_name') or '').strip()
    if legal:
        names.add(legal)
    for name in person.get('name_list') or []:
        local = (name.get('display_name_local_script') or '').strip()
        if local:
            names.add(local)
    return names


def build_employee_status_maps():
    """分页拉全量员工, 返回 (by_number, by_name_unique):

      by_number:      工号(employee_number) -> employment_status.enum_name(如 'hired');
      by_name_unique: 姓名 -> enum_name, 仅当该姓名在飞书唯一对应一名员工(重名则剔除)。
    """
    by_number = {}
    name_records = {}
    for item in search_all(_EMPLOYEE_SEARCH_PATH,
                           fields=['employee_number', 'employment_status',
                                   'person_info.legal_name', 'person_info.name_list'],
                           extra_params={'user_id_type': 'user_id'}):
        status = ((item.get('employment_status') or {}).get('enum_name') or '').strip()
        number = (item.get('employee_number') or '').strip()
        if number:
            by_number[number] = status
        for name in _item_names(item):
            name_records.setdefault(name, []).append(status)
    by_name_unique = {name: statuses[0] for name, statuses in name_records.items() if len(statuses) == 1}
    return by_number, by_name_unique


def get_employee_status_maps():
    """带进程内缓存的 (by_number, by_name_unique)(同一次运行只拉一次)。"""
    global _STATUS_MAPS_CACHE
    if _STATUS_MAPS_CACHE is None:
        _STATUS_MAPS_CACHE = build_employee_status_maps()
    return _STATUS_MAPS_CACHE
