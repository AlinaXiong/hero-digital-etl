# -*- coding: utf-8 -*-
"""飞书(corehr)公共类: 取员工/部门/公司/地点等信息。

凭据来自环境变量 FEISHU_APP_ID / FEISHU_APP_SECRET(可写入项目根 .env)。
- get_employee_status_maps(): 供合同"申请人状态(在职/离职)"判断;
- fetch_all_employees() + fetch_*_name_map(): 供全量员工信息导出。

飞书工号 employee_number 与泛微 hrmjobtitles.JOBTITLENAME / 汉得 employee_code
同口径(均为 V 编号)。
"""
import json
import mimetypes
import os
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path

FEISHU_HOST = os.getenv('FEISHU_HOST', 'https://open.feishu.cn').rstrip('/')
_TOKEN_PATH = '/open-apis/auth/v3/tenant_access_token/internal'
_EMPLOYEE_SEARCH_PATH = '/open-apis/corehr/v2/employees/search'
_DEPARTMENT_BATCH_PATH = '/open-apis/corehr/v2/departments/batch_get'
_COMPANY_LIST_PATH = '/open-apis/corehr/v1/companies'
_LOCATION_LIST_PATH = '/open-apis/corehr/v1/locations'
_EMPLOYEE_TYPE_LIST_PATH = '/open-apis/corehr/v1/employee_types'
_MESSAGE_CREATE_PATH = '/open-apis/im/v1/messages'
_FILE_UPLOAD_PATH = '/open-apis/im/v1/files'
_MAX_MESSAGE_FILE_BYTES = 30 * 1024 * 1024
_MAX_RETRY = 4

_STATUS_MAPS_CACHE = None
_TOKEN_CACHE = None
_MESSAGE_TOKEN_CACHE = None


def _request(method, path, params=None, payload=None, timeout=30, auth=True, token=None):
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
            req.add_header('Authorization', f'Bearer {token or _token()}')
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode('utf-8'))
        except urllib.error.HTTPError as error:
            # 4xx 多为权限、应用可用范围或收件人 ID 问题；保留飞书响应体便于定位。
            detail = error.read().decode('utf-8', errors='replace').strip()
            raise RuntimeError(f'飞书 API HTTP {error.code}: {detail or error.reason}') from error
        except (urllib.error.URLError, TimeoutError, ConnectionError) as error:
            last_error = error
            if attempt < _MAX_RETRY - 1:
                time.sleep(1.5 * (attempt + 1))
    raise last_error


def _request_multipart(path, fields, file_field, file_path, timeout=60, token=None):
    """上传飞书 IM 文件，保留与 JSON 请求相同的错误诊断与网络重试。"""
    target = Path(file_path)
    boundary = f'----HeroDigitalEtl{uuid.uuid4().hex}'
    chunks = []
    for name, value in fields.items():
        chunks.extend((
            f'--{boundary}\r\n'.encode('ascii'),
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode('utf-8'),
            str(value).encode('utf-8'),
            b'\r\n',
        ))
    content_type = mimetypes.guess_type(target.name)[0] or 'application/octet-stream'
    chunks.extend((
        f'--{boundary}\r\n'.encode('ascii'),
        (
            f'Content-Disposition: form-data; name="{file_field}"; '
            f'filename="{target.name}"\r\n'
        ).encode('utf-8'),
        f'Content-Type: {content_type}\r\n\r\n'.encode('ascii'),
        target.read_bytes(),
        b'\r\n',
        f'--{boundary}--\r\n'.encode('ascii'),
    ))
    data = b''.join(chunks)
    last_error = None
    for attempt in range(_MAX_RETRY):
        request = urllib.request.Request(FEISHU_HOST + path, data=data, method='POST')
        request.add_header('Authorization', f'Bearer {token or _token()}')
        request.add_header('Content-Type', f'multipart/form-data; boundary={boundary}')
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return json.loads(response.read().decode('utf-8'))
        except urllib.error.HTTPError as error:
            detail = error.read().decode('utf-8', errors='replace').strip()
            raise RuntimeError(f'飞书文件上传 HTTP {error.code}: {detail or error.reason}') from error
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


def _get_tenant_access_token_for(app_id, app_secret, credential_name):
    if not app_id or not app_secret:
        raise RuntimeError(f'缺少 {credential_name}; 请在 .env 配置')
    result = _request('POST', _TOKEN_PATH, payload={'app_id': app_id, 'app_secret': app_secret}, auth=False)
    if result.get('code') != 0:
        raise RuntimeError(f"飞书获取 tenant_access_token 失败: {result.get('code')} {result.get('msg')}")
    return result['tenant_access_token']


def _get_tenant_access_token():
    return _get_tenant_access_token_for(
        os.environ.get('FEISHU_APP_ID', '').strip(),
        os.environ.get('FEISHU_APP_SECRET', '').strip(),
        'FEISHU_APP_ID / FEISHU_APP_SECRET',
    )


def _token():
    global _TOKEN_CACHE
    if not _TOKEN_CACHE:
        _TOKEN_CACHE = _get_tenant_access_token()
    return _TOKEN_CACHE


def _message_token():
    """通知机器人可单独配置，避免影响用于 CoreHR 查询的既有应用。"""
    global _MESSAGE_TOKEN_CACHE
    app_id = os.environ.get('FEISHU_NOTIFY_APP_ID', '').strip()
    app_secret = os.environ.get('FEISHU_NOTIFY_APP_SECRET', '').strip()
    if not app_id and not app_secret:
        return _token()
    if not app_id or not app_secret:
        raise RuntimeError('FEISHU_NOTIFY_APP_ID / FEISHU_NOTIFY_APP_SECRET 必须同时配置')
    if not _MESSAGE_TOKEN_CACHE:
        _MESSAGE_TOKEN_CACHE = _get_tenant_access_token_for(
            app_id,
            app_secret,
            'FEISHU_NOTIFY_APP_ID / FEISHU_NOTIFY_APP_SECRET',
        )
    return _MESSAGE_TOKEN_CACHE


def _check(result, tag):
    if result.get('code') != 0:
        raise RuntimeError(f"飞书 {tag} 失败: {result.get('code')} {result.get('msg')}")
    return result.get('data') or {}


def send_text_message(open_id, text):
    """以应用机器人身份向指定飞书 open_id 发送文本消息。"""
    receiver = str(open_id or '').strip()
    if not receiver:
        raise ValueError('飞书 open_id 不能为空')
    content = str(text or '').strip()
    if not content:
        raise ValueError('飞书消息内容不能为空')
    result = _request(
        'POST',
        _MESSAGE_CREATE_PATH,
        params={'receive_id_type': 'open_id'},
        payload={
            'receive_id': receiver,
            'msg_type': 'text',
            'content': json.dumps({'text': content}, ensure_ascii=False),
        },
        token=_message_token(),
    )
    return _check(result, '发送消息')


def send_interactive_card(open_id, card):
    """以应用机器人身份向指定 open_id 发送飞书交互卡片。"""
    receiver = str(open_id or '').strip()
    if not receiver:
        raise ValueError('飞书 open_id 不能为空')
    if not isinstance(card, dict) or not card:
        raise ValueError('飞书卡片内容不能为空')
    result = _request(
        'POST',
        _MESSAGE_CREATE_PATH,
        params={'receive_id_type': 'open_id'},
        payload={
            'receive_id': receiver,
            'msg_type': 'interactive',
            'content': json.dumps(card, ensure_ascii=False),
        },
        token=_message_token(),
    )
    return _check(result, '发送卡片消息')


def send_file_message(open_id, file_path):
    """以应用机器人身份上传并发送一个不超过 30 MB 的本地文件。"""
    receiver = str(open_id or '').strip()
    if not receiver:
        raise ValueError('飞书 open_id 不能为空')
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f'待发送的飞书文件不存在: {path}')
    size = path.stat().st_size
    if size <= 0:
        raise ValueError(f'待发送的飞书文件为空: {path.name}')
    if size > _MAX_MESSAGE_FILE_BYTES:
        raise ValueError(
            f'飞书单文件最大支持 30 MB，当前 ZIP 为 {size / 1024 / 1024:.1f} MB: {path.name}'
        )

    token = _message_token()
    upload = _request_multipart(
        _FILE_UPLOAD_PATH,
        fields={'file_type': 'stream', 'file_name': path.name},
        file_field='file',
        file_path=path,
        token=token,
    )
    file_key = _check(upload, '上传文件').get('file_key')
    if not file_key:
        raise RuntimeError(f'飞书上传文件未返回 file_key: {upload}')
    result = _request(
        'POST',
        _MESSAGE_CREATE_PATH,
        params={'receive_id_type': 'open_id'},
        payload={
            'receive_id': receiver,
            'msg_type': 'file',
            'content': json.dumps({'file_key': file_key}, ensure_ascii=False),
        },
        token=token,
    )
    return _check(result, '发送文件消息')


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
