# -*- coding: utf-8 -*-
"""泛微客商附件同步到业财 UAT。

从泛微 ``uf_khgys`` 读取客商证件号/纳税人识别号及以下附件字段，
精确匹配业财 UAT 客户、供应商后，上传到其「营业执照」附件页：

* gsjjdzb：公司简介电子版
* yyzzsmj：营业执照扫描件
* khxksmj：开户许可扫描件
* ztsmxianzhang、jtgssmxz：主体/集团扫描鲜章
* yqlcxz：印签留存鲜章
* gsqdxz：公司清单鲜章

默认仅生成匹配日志（dry-run），必须显式传入 ``--execute`` 才会下载泛微
文件并上传 UAT。仅对「一个泛微客商 -> 一个 UAT 客商」的精确匹配执行上传；
无匹配、同一客商类型内重复匹配、缺少附件关联键都会写入结果清单而不会上传。
同一标识同时唯一匹配供应商和客户时，会分别上传到两者的营业执照附件页。

运行前配置（项目根目录 .env.local 优先）：

    HFBS_UAT_ACCESS_TOKEN=<UAT 登录 Bearer Token，必填>
    WEAVER_CONTRACT_ATTACHMENT_COOKIE=<泛微 Cookie，执行上传时必填>

可选配置：

    HFBS_UAT_BASE_URL=https://uat.link.heroesports.com/gtw/hfbs
    HFBS_UAT_ORGANIZATION_ID=0
    HFBS_UAT_PAGE_SIZE=200
    HFBS_UAT_TIMEOUT_SECONDS=90
    HFBS_UAT_UPLOAD_RETRIES=3

示例：

    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py
    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py --source-id 12345
    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py --execute

@author xiongyilin@heroesports.com
@since 2026-08-31
"""
import argparse
import csv
import json
import logging
import mimetypes
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

if __package__ is None or __package__ == '':
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from etl.util import common as c


TASK_NAME = 'sync_weaver_partner_attachments_to_hfbs_uat'
OUTPUT_DIR = c.OUT_DIR / TASK_NAME
DEFAULT_UAT_BASE_URL = 'https://uat.link.heroesports.com/gtw/hfbs'
DEFAULT_UAT_ORGANIZATION_ID = '0'
DEFAULT_PAGE_SIZE = 200
DEFAULT_TIMEOUT_SECONDS = 90
DEFAULT_UPLOAD_RETRIES = 3
DEFAULT_MAX_FILE_BYTES = 50 * 1024 * 1024
UAT_TOKEN_ENV = 'HFBS_UAT_ACCESS_TOKEN'
UAT_BASE_URL_ENV = 'HFBS_UAT_BASE_URL'
UAT_ORGANIZATION_ID_ENV = 'HFBS_UAT_ORGANIZATION_ID'
UAT_PAGE_SIZE_ENV = 'HFBS_UAT_PAGE_SIZE'
UAT_TIMEOUT_ENV = 'HFBS_UAT_TIMEOUT_SECONDS'
UAT_UPLOAD_RETRIES_ENV = 'HFBS_UAT_UPLOAD_RETRIES'
UAT_MAX_FILE_BYTES_ENV = 'HFBS_UAT_MAX_FILE_BYTES'

ATTACHMENT_FIELDS: Tuple[Tuple[str, str], ...] = (
    ('gsjjdzb', '公司简介电子版'),
    ('yyzzsmj', '营业执照扫描件'),
    ('khxksmj', '开户许可扫描件'),
    ('ztsmxianzhang', '主体扫描/鲜章'),
    ('jtgssmxz', '集团公司扫描/鲜章'),
    ('yqlcxz', '印签留存鲜章'),
    ('gsqdxz', '公司清单鲜章'),
)
ATTACHMENT_FIELD_LABELS = dict(ATTACHMENT_FIELDS)
IDENTIFIER_SOURCE_FIELDS: Tuple[Tuple[str, str], ...] = (
    ('sh', '证件号/税号'),
    ('khsh', '纳税人识别号'),
)
TARGET_CONFIG = {
    'customer': {
        'endpoint': 'system-customer',
        'attachment_table': 'CUSTOMER_APPLY_HEADER',
        'primary_key_field': 'customerId',
        'code_field': 'customerCode',
        'name_fields': ('description', 'taxpayerName'),
        'identifier_fields': ('taxpayerNumber', 'taxIdNumber', 'certificateNumber', 'certificateNo'),
    },
    'vender': {
        'endpoint': 'system-vender',
        'attachment_table': 'VENDER_APPLY_HEADER',
        'primary_key_field': 'venderId',
        'code_field': 'venderCode',
        'name_fields': ('description', 'taxpayerName'),
        'identifier_fields': ('taxIdNumber', 'taxpayerNumber', 'certificateNumber', 'certificateNo'),
    },
}
RESULT_COLUMNS = (
    'time',
    'status',
    'message',
    'source_id',
    'source_name',
    'source_identifier_type',
    'source_identifier_masked',
    'source_attachment_field',
    'source_attachment_label',
    'source_docid',
    'weaver_imagefileid',
    'attachment_name',
    'target_type',
    'target_code',
    'target_name',
    'target_primary_id',
    'target_unique_code',
    'target_unique_code_source',
    'matched_target_field',
    'http_status',
)
DOC_ID_RE = re.compile(r'\d+')
IDENTIFIER_NORMALIZE_RE = re.compile(r'[^0-9A-Za-z]')


@dataclass(frozen=True)
class SourceAttachment:
    source_id: str
    source_name: str
    identifier_values: Tuple[Tuple[str, str], ...]
    field_name: str
    field_label: str
    docid: int


@dataclass(frozen=True)
class AttachmentMeta:
    docid: int
    imagefileid: int
    attachment_name: str
    file_size: Optional[int]


@dataclass(frozen=True)
class UatTarget:
    target_type: str
    primary_id: str
    code: str
    name: str
    attachment_table: str
    unique_code: str
    unique_code_source: str
    identifier_fields: Tuple[Tuple[str, str], ...]


@dataclass(frozen=True)
class TargetMatch:
    """一条泛微附件与一个唯一 UAT 客商的匹配结果。"""

    target: UatTarget
    source_identifier_type: str
    target_identifier_field: str


@dataclass(frozen=True)
class MatchOutcome:
    """按客户/供应商分别判重后的匹配结果。"""

    matches: Tuple[TargetMatch, ...]
    reason: str = ''
    ambiguous_target_types: Tuple[str, ...] = ()


class UatRequestError(RuntimeError):
    """UAT HTTP 调用失败，保留可写入日志的状态码和错误摘要。"""

    def __init__(self, message: str, status_code: Optional[int] = None):
        super().__init__(message)
        self.status_code = status_code


def _text(value) -> str:
    if value is None:
        return ''
    text = str(value).strip()
    return '' if text.lower() in ('', 'nan', 'none', 'nat') else text


def _positive_int_env(name: str, default: int, minimum: int = 1, maximum: Optional[int] = None) -> int:
    raw = os.getenv(name, '').strip()
    try:
        value = int(raw) if raw else default
    except ValueError:
        value = default
    value = max(minimum, value)
    return min(value, maximum) if maximum is not None else value


def _normalize_identifier(value) -> str:
    text = unicodedata.normalize('NFKC', _text(value)).upper()
    return IDENTIFIER_NORMALIZE_RE.sub('', text)


def _mask_identifier(value: str) -> str:
    value = _text(value)
    if len(value) <= 8:
        return '*' * len(value)
    return f'{value[:4]}****{value[-4:]}'


def _extract_docids(value) -> List[int]:
    """附件字段是 DOCID 列表，兼容英文/中文逗号和泛微格式化文本。"""
    seen: Set[int] = set()
    docids: List[int] = []
    for raw_docid in DOC_ID_RE.findall(_text(value)):
        docid = int(raw_docid)
        if docid > 0 and docid not in seen:
            seen.add(docid)
            docids.append(docid)
    return docids


def _safe_filename(name: str, docid: int, imagefileid: int) -> str:
    name = Path(_text(name)).name.replace('\x00', '').strip()
    if not name:
        return f'weaver_doc_{docid}_file_{imagefileid}'
    return name


def _normalized_file_name(value: str) -> str:
    """与 UAT 文件列表保持一致：文件名可能经过 URL 编码。"""
    return urllib.parse.unquote(_text(value)).casefold()


def _collision_renamed_attachment_name(meta: AttachmentMeta) -> str:
    """同名时保留扩展名，并追加稳定的泛微文件标识以支持重复运行去重。"""
    filename = _safe_filename(meta.attachment_name, meta.docid, meta.imagefileid)
    suffix = Path(filename).suffix
    stem = filename[:-len(suffix)] if suffix else filename
    return f'{stem}__weaver_doc{meta.docid}_file{meta.imagefileid}{suffix}'


def _setup_logger(run_id: str) -> Tuple[logging.Logger, Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log_file = OUTPUT_DIR / f'同步日志_{run_id}.log'
    logger = logging.getLogger(f'{TASK_NAME}.{run_id}')
    logger.setLevel(logging.INFO)
    logger.propagate = False
    formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.handlers[:] = [console_handler, file_handler]
    return logger, log_file


def _validate_source_fields() -> None:
    expected_fields = {
        'khmc': '企业名称',
        'sh': '税号',
        'khsh': '纳税人识别号',
    }
    expected_fields.update(ATTACHMENT_FIELD_LABELS)
    c.validate_fw_fields('uf_khgys', {'': expected_fields})


def _load_source_attachments(source_ids: Sequence[str]) -> List[SourceAttachment]:
    attachment_selects = ', '.join(f'k.{field_name}' for field_name, _ in ATTACHMENT_FIELDS)
    attachment_conditions = ' OR '.join(
        f"COALESCE(TRIM(k.{field_name}), '') <> ''" for field_name, _ in ATTACHMENT_FIELDS
    )
    params: List[str] = []
    where_parts = [f'({attachment_conditions})']
    if source_ids:
        where_parts.append(f'k.id IN ({c.in_placeholders(source_ids)})')
        params.extend(source_ids)
    source_df = c.query_db(
        'FW',
        'vspn_xtyy',
        f'''
        SELECT k.id, k.khmc, k.sh, k.khsh, {attachment_selects}
        FROM uf_khgys k
        WHERE {' AND '.join(where_parts)}
        ORDER BY k.id
        ''',
        params,
    )

    attachments: List[SourceAttachment] = []
    for _, row in source_df.iterrows():
        identifiers = tuple(
            (field_name, _text(row.get(field_name)))
            for field_name, _ in IDENTIFIER_SOURCE_FIELDS
            if _normalize_identifier(row.get(field_name))
        )
        for field_name, field_label in ATTACHMENT_FIELDS:
            for docid in _extract_docids(row.get(field_name)):
                attachments.append(
                    SourceAttachment(
                        source_id=_text(row.get('id')),
                        source_name=_text(row.get('khmc')),
                        identifier_values=identifiers,
                        field_name=field_name,
                        field_label=field_label,
                        docid=docid,
                    )
                )
    return attachments


def _load_attachment_metadata(docids: Iterable[int]) -> Dict[int, List[AttachmentMeta]]:
    docids = sorted(set(docid for docid in docids if docid > 0))
    if not docids:
        return {}

    metadata: Dict[int, List[AttachmentMeta]] = defaultdict(list)
    batch_size = 800
    for start in range(0, len(docids), batch_size):
        batch = docids[start:start + batch_size]
        rows = c.query_db(
            'FW',
            'vspn_xtyy',
            f'''
            SELECT
                d.DOCID AS docid,
                d.IMAGEFILEID AS imagefileid,
                COALESCE(NULLIF(i.IMAGEFILENAME, ''), NULLIF(d.IMAGEFILENAME, '')) AS attachment_name,
                i.FILESIZE AS file_size
            FROM docimagefile d
            LEFT JOIN imagefile i ON i.IMAGEFILEID = d.IMAGEFILEID
            WHERE d.DOCID IN ({c.in_placeholders(batch)})
            ORDER BY d.DOCID, d.IMAGEFILEID
            ''',
            batch,
        )
        for _, row in rows.iterrows():
            docid = int(row['docid'])
            imagefileid = int(row['imagefileid'])
            raw_size = _text(row.get('file_size'))
            try:
                file_size = int(raw_size) if raw_size else None
            except ValueError:
                file_size = None
            metadata[docid].append(
                AttachmentMeta(
                    docid=docid,
                    imagefileid=imagefileid,
                    attachment_name=_safe_filename(row.get('attachment_name'), docid, imagefileid),
                    file_size=file_size,
                )
            )
    return metadata


def _response_text(response_body: bytes) -> str:
    return response_body.decode('utf-8', errors='replace').replace('\n', ' ')[:500]


def _authorization_headers(access_token: str) -> Dict[str, str]:
    return {
        'Accept': 'application/json, text/plain, */*',
        'Authorization': f'bearer {access_token}',
    }


def _request_json(url: str, access_token: str, timeout_seconds: int) -> Tuple[object, int]:
    request = urllib.request.Request(url, headers=_authorization_headers(access_token), method='GET')
    try:
        with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
            body = response.read()
            status_code = response.getcode()
    except urllib.error.HTTPError as exc:
        raise UatRequestError(f'UAT HTTP {exc.code}: {_response_text(exc.read())}', exc.code) from exc
    except urllib.error.URLError as exc:
        raise UatRequestError(f'UAT 网络错误: {exc.reason}') from exc
    try:
        return json.loads(body.decode('utf-8')), status_code
    except json.JSONDecodeError as exc:
        raise UatRequestError(f'UAT 返回非 JSON: {_response_text(body)}', status_code) from exc


def _extract_page_items(payload: object) -> Tuple[List[dict], Optional[int]]:
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)], len(payload)
    if not isinstance(payload, dict):
        raise UatRequestError(f'UAT 分页返回格式异常: {type(payload).__name__}')

    if payload.get('failed'):
        raise UatRequestError(_text(payload.get('message')) or _text(payload.get('detailsMessage')) or 'UAT 请求失败')

    for container in (payload, payload.get('data')):
        if not isinstance(container, dict):
            continue
        content = container.get('content')
        if isinstance(content, list):
            total = container.get('totalElements')
            try:
                total = int(total) if total is not None else None
            except (TypeError, ValueError):
                total = None
            return [item for item in content if isinstance(item, dict)], total
        if isinstance(container.get('rows'), list):
            total = container.get('total')
            try:
                total = int(total) if total is not None else None
            except (TypeError, ValueError):
                total = None
            return [item for item in container['rows'] if isinstance(item, dict)], total
    if isinstance(payload.get('data'), list):
        return [item for item in payload['data'] if isinstance(item, dict)], len(payload['data'])
    raise UatRequestError('UAT 分页返回中未找到 content/rows/data 列表')


def _fetch_all_uat_records(
    base_url: str,
    organization_id: str,
    endpoint: str,
    access_token: str,
    page_size: int,
    timeout_seconds: int,
    logger: logging.Logger,
) -> List[dict]:
    records: List[dict] = []
    page = 0
    while True:
        query = urllib.parse.urlencode({'page': page, 'size': page_size})
        url = f'{base_url.rstrip("/")}/v1/{organization_id}/{endpoint}?{query}'
        payload, _ = _request_json(url, access_token, timeout_seconds)
        items, total = _extract_page_items(payload)
        records.extend(items)
        logger.info('[UAT] %s 第 %s 页: %s 条，累计 %s 条', endpoint, page, len(items), len(records))
        if not items:
            break
        if total is not None and len(records) >= total:
            break
        if total is None and len(items) < page_size:
            break
        page += 1
    return records


def _target_name(record: dict, name_fields: Sequence[str]) -> str:
    return next((_text(record.get(field_name)) for field_name in name_fields if _text(record.get(field_name))), '')


def _build_uat_targets(
    target_type: str,
    records: Iterable[dict],
    allow_primary_key_unique_code: bool,
) -> List[UatTarget]:
    config = TARGET_CONFIG[target_type]
    targets: List[UatTarget] = []
    for record in records:
        primary_id = _text(record.get(config['primary_key_field']))
        if not primary_id:
            continue
        header_id = _text(record.get('headerId'))
        if header_id:
            unique_code, unique_code_source = header_id, 'headerId'
        elif allow_primary_key_unique_code:
            unique_code, unique_code_source = primary_id, config['primary_key_field']
        else:
            unique_code, unique_code_source = '', ''

        identifier_fields = tuple(
            (field_name, _text(record.get(field_name)))
            for field_name in config['identifier_fields']
            if _normalize_identifier(record.get(field_name))
        )
        if not identifier_fields:
            continue
        targets.append(
            UatTarget(
                target_type=target_type,
                primary_id=primary_id,
                code=_text(record.get(config['code_field'])),
                name=_target_name(record, config['name_fields']),
                attachment_table=config['attachment_table'],
                unique_code=unique_code,
                unique_code_source=unique_code_source,
                identifier_fields=identifier_fields,
            )
        )
    return targets


def _build_identifier_index(targets: Iterable[UatTarget]) -> Dict[str, List[Tuple[UatTarget, str]]]:
    index: Dict[str, List[Tuple[UatTarget, str]]] = defaultdict(list)
    for target in targets:
        for field_name, identifier in target.identifier_fields:
            index[_normalize_identifier(identifier)].append((target, field_name))
    return index


def _match_targets(
    source: SourceAttachment,
    identifier_index: Dict[str, List[Tuple[UatTarget, str]]],
) -> MatchOutcome:
    """按客商类型独立匹配，允许同一标识同时命中一个客户和一个供应商。

    同一类型内有多条候选仍视为歧义，避免把附件挂错；另一类型若唯一命中，
    则继续处理并额外写出 ``identifier_ambiguous_partial`` 日志。
    """
    matched: Dict[Tuple[str, str], Tuple[UatTarget, str, str]] = {}
    for source_field, source_identifier in source.identifier_values:
        normalized = _normalize_identifier(source_identifier)
        for target, target_field in identifier_index.get(normalized, []):
            matched[(target.target_type, target.primary_id)] = (target, source_field, target_field)

    if not matched:
        return MatchOutcome((), 'identifier_not_found')

    candidates_by_type: Dict[str, List[Tuple[UatTarget, str, str]]] = defaultdict(list)
    for target, source_field, target_field in matched.values():
        candidates_by_type[target.target_type].append((target, source_field, target_field))

    matches: List[TargetMatch] = []
    ambiguous_target_types: List[str] = []
    for target_type in TARGET_CONFIG:
        candidates = candidates_by_type.get(target_type, [])
        if len(candidates) == 1:
            target, source_field, target_field = candidates[0]
            matches.append(TargetMatch(target, source_field, target_field))
        elif len(candidates) > 1:
            ambiguous_target_types.append(target_type)

    if not matches:
        return MatchOutcome((), 'identifier_ambiguous', tuple(ambiguous_target_types))
    if ambiguous_target_types:
        return MatchOutcome(
            tuple(matches), 'identifier_ambiguous_partial', tuple(ambiguous_target_types),
        )
    return MatchOutcome(tuple(matches))


def _build_weaver_headers(cookie: str, meta: AttachmentMeta) -> Dict[str, str]:
    base_url = os.getenv(c.ATTACHMENT_BASE_URL_ENV, c.DEFAULT_ATTACHMENT_BASE_URL).rstrip('/')
    return {
        'Accept': '*/*',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Connection': 'keep-alive',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/146 Safari/537.36',
        'Cookie': cookie,
        'Referer': c.build_attachment_referer(base_url, meta.imagefileid, meta.docid),
    }


def _download_weaver_attachment(
    cookie: str,
    meta: AttachmentMeta,
    timeout_seconds: int,
    max_file_bytes: int,
) -> bytes:
    base_url = os.getenv(c.ATTACHMENT_BASE_URL_ENV, c.DEFAULT_ATTACHMENT_BASE_URL).rstrip('/')
    url = f'{base_url}/weaver/weaver.file.FileDownload?fileid={meta.imagefileid}'
    request = urllib.request.Request(url, headers=_build_weaver_headers(cookie, meta))
    try:
        with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
            final_url = response.geturl()
            content_type = response.headers.get('Content-Type', '')
            content_length = response.headers.get('Content-Length', '')
            try:
                expected_length = int(content_length) if content_length else None
            except ValueError:
                expected_length = None
            if expected_length is not None and expected_length > max_file_bytes:
                raise RuntimeError(f'泛微附件过大: {expected_length} bytes，超过限制 {max_file_bytes} bytes')
            data = response.read(max_file_bytes + 1)
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f'泛微下载 HTTP {exc.code}: {_response_text(exc.read())}') from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f'泛微下载网络错误: {exc.reason}') from exc

    if len(data) > max_file_bytes:
        raise RuntimeError(f'泛微附件超过限制 {max_file_bytes} bytes')
    if 'login' in final_url.lower():
        raise RuntimeError(f'泛微下载跳转登录页: {final_url}')
    if 'text/html' in content_type.lower() and not meta.attachment_name.lower().endswith(('.html', '.htm')):
        raise RuntimeError(f'泛微下载返回 HTML: {_response_text(data)}')
    if not data:
        raise RuntimeError('泛微下载返回空文件')
    return data


def _multipart_body(filename: str, data: bytes, order_number: str) -> Tuple[bytes, str]:
    boundary = f'----HeroAttachmentSync{time.time_ns()}'
    mime_type = mimetypes.guess_type(filename)[0] or 'application/octet-stream'
    chunks = [
        f'--{boundary}\r\n'.encode(),
        b'Content-Disposition: form-data; name="orderNumber"\r\n\r\n',
        order_number.encode(),
        b'\r\n',
        f'--{boundary}\r\n'.encode(),
        (
            'Content-Disposition: form-data; name="file"; '
            f'filename="{filename.replace(chr(34), "_")}"\r\n'
        ).encode('utf-8'),
        f'Content-Type: {mime_type}\r\n\r\n'.encode(),
        data,
        b'\r\n',
        f'--{boundary}--\r\n'.encode(),
    ]
    return b''.join(chunks), f'multipart/form-data; boundary={boundary}'


def _fetch_existing_attachment_names(
    base_url: str,
    organization_id: str,
    access_token: str,
    target: UatTarget,
    page_size: int,
    timeout_seconds: int,
) -> Set[str]:
    names: Set[str] = set()
    page = 0
    total_records = 0
    while True:
        query = urllib.parse.urlencode({
            'tableName': target.attachment_table,
            'uniqueCode': target.unique_code,
            'page': page,
            'size': page_size,
        })
        url = f'{base_url.rstrip("/")}/v1/{organization_id}/file/page?{query}'
        payload, _ = _request_json(url, access_token, timeout_seconds)
        items, total = _extract_page_items(payload)
        total_records += len(items)
        names.update(_normalized_file_name(item.get('fileName')) for item in items if _text(item.get('fileName')))
        if not items:
            break
        if total is not None and total_records >= total:
            break
        if total is None and len(items) < page_size:
            break
        page += 1
    return names


def _upload_to_uat(
    base_url: str,
    organization_id: str,
    access_token: str,
    target: UatTarget,
    meta: AttachmentMeta,
    data: bytes,
    timeout_seconds: int,
) -> int:
    query = urllib.parse.urlencode({
        'bucketName': 'hfins',
        'tableName': target.attachment_table,
        'uniqueCode': target.unique_code,
    })
    url = f'{base_url.rstrip("/")}/v1/{organization_id}/file?{query}'
    body, content_type = _multipart_body(meta.attachment_name, data, str(int(time.time() * 1000)))
    headers = _authorization_headers(access_token)
    headers['Content-Type'] = content_type
    headers['Content-Length'] = str(len(body))
    request = urllib.request.Request(url, data=body, headers=headers, method='POST')
    try:
        with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
            response_body = response.read()
            status_code = response.getcode()
    except urllib.error.HTTPError as exc:
        raise UatRequestError(f'UAT 上传 HTTP {exc.code}: {_response_text(exc.read())}', exc.code) from exc
    except urllib.error.URLError as exc:
        raise UatRequestError(f'UAT 上传网络错误: {exc.reason}') from exc

    if response_body:
        try:
            payload = json.loads(response_body.decode('utf-8'))
        except json.JSONDecodeError:
            payload = None
        if isinstance(payload, dict) and payload.get('failed'):
            raise UatRequestError(
                _text(payload.get('message')) or _text(payload.get('detailsMessage')) or 'UAT 上传失败',
                status_code,
            )
    return status_code


def _result_row(
    source: SourceAttachment,
    status: str,
    message: str,
    target: Optional[UatTarget] = None,
    source_identifier_type: str = '',
    target_identifier_field: str = '',
    meta: Optional[AttachmentMeta] = None,
    http_status: Optional[int] = None,
) -> dict:
    identifiers = dict(source.identifier_values)
    identifier = identifiers.get(source_identifier_type, '')
    return {
        'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'status': status,
        'message': message,
        'source_id': source.source_id,
        'source_name': source.source_name,
        'source_identifier_type': source_identifier_type,
        'source_identifier_masked': _mask_identifier(identifier),
        'source_attachment_field': source.field_name,
        'source_attachment_label': source.field_label,
        'source_docid': source.docid,
        'weaver_imagefileid': meta.imagefileid if meta else '',
        'attachment_name': meta.attachment_name if meta else '',
        'target_type': target.target_type if target else '',
        'target_code': target.code if target else '',
        'target_name': target.name if target else '',
        'target_primary_id': target.primary_id if target else '',
        'target_unique_code': target.unique_code if target else '',
        'target_unique_code_source': target.unique_code_source if target else '',
        'matched_target_field': target_identifier_field,
        'http_status': http_status or '',
    }


def _write_result(writer: csv.DictWriter, logger: logging.Logger, row: dict) -> None:
    writer.writerow(row)
    logger.info(
        '[%s] source=%s doc=%s target=%s/%s file=%s %s',
        row['status'],
        row['source_id'],
        row['source_docid'],
        row['target_type'] or '-',
        row['target_code'] or '-',
        row['attachment_name'] or '-',
        row['message'],
    )


def _retry_upload(
    base_url: str,
    organization_id: str,
    access_token: str,
    target: UatTarget,
    meta: AttachmentMeta,
    data: bytes,
    timeout_seconds: int,
    retries: int,
) -> int:
    last_error: Optional[Exception] = None
    for attempt in range(1, retries + 1):
        try:
            return _upload_to_uat(
                base_url, organization_id, access_token, target, meta, data, timeout_seconds,
            )
        except UatRequestError as exc:
            last_error = exc
            retryable = exc.status_code in (429, 500, 502, 503, 504) or exc.status_code is None
            if not retryable or attempt == retries:
                raise
            time.sleep(min(attempt * 2, 10))
    raise last_error or RuntimeError('未知上传失败')


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description='泛微客商附件同步到业财 UAT 营业执照页')
    parser.add_argument(
        '--source-id', action='append', default=[],
        help='仅处理指定泛微 uf_khgys.id；可重复传入，例如 --source-id 1 --source-id 2',
    )
    parser.add_argument('--limit', type=int, default=0, help='最多处理多少个泛微客商 ID（0 表示全部）')
    parser.add_argument('--execute', action='store_true', help='真正下载泛微附件并上传 UAT；默认仅 dry-run')
    parser.add_argument(
        '--allow-primary-key-unique-code', action='store_true',
        help='UAT 列表未返回 headerId 时，允许使用 customerId/venderId 作为附件 uniqueCode（默认禁止）',
    )
    parser.add_argument(
        '--no-skip-existing', action='store_true',
        help='同名时仍使用原文件名上传（默认改名为 __weaver_doc{DOCID}_file{IMAGEFILEID} 后上传）',
    )
    return parser


def run(argv: Optional[Sequence[str]] = None) -> Path:
    args = build_parser().parse_args(argv)
    run_id = datetime.now().strftime('%Y%m%d_%H%M%S')
    logger, log_file = _setup_logger(run_id)
    result_file = OUTPUT_DIR / f'同步结果_{run_id}.csv'
    access_token = os.getenv(UAT_TOKEN_ENV, '').strip()
    if not access_token:
        raise RuntimeError(f'缺少 {UAT_TOKEN_ENV}；请配置 UAT Bearer Token 后重试。')
    cookie = os.getenv(c.ATTACHMENT_COOKIE_ENV, '').strip()
    if args.execute and not cookie:
        raise RuntimeError(f'缺少 {c.ATTACHMENT_COOKIE_ENV}；执行上传前必须配置有效泛微 Cookie。')

    base_url = os.getenv(UAT_BASE_URL_ENV, DEFAULT_UAT_BASE_URL).strip().rstrip('/')
    organization_id = os.getenv(UAT_ORGANIZATION_ID_ENV, DEFAULT_UAT_ORGANIZATION_ID).strip()
    page_size = _positive_int_env(UAT_PAGE_SIZE_ENV, DEFAULT_PAGE_SIZE, maximum=500)
    timeout_seconds = _positive_int_env(UAT_TIMEOUT_ENV, DEFAULT_TIMEOUT_SECONDS, maximum=300)
    upload_retries = _positive_int_env(UAT_UPLOAD_RETRIES_ENV, DEFAULT_UPLOAD_RETRIES, maximum=10)
    max_file_bytes = _positive_int_env(UAT_MAX_FILE_BYTES_ENV, DEFAULT_MAX_FILE_BYTES, maximum=500 * 1024 * 1024)
    logger.info(
        '开始同步: mode=%s base_url=%s org=%s page_size=%s source_ids=%s',
        'execute' if args.execute else 'dry-run', base_url, organization_id, page_size,
        ','.join(args.source_id) if args.source_id else 'all',
    )

    _validate_source_fields()
    source_attachments = _load_source_attachments(args.source_id)
    source_ids_in_order = list(dict.fromkeys(item.source_id for item in source_attachments))
    if args.limit > 0:
        allowed_source_ids = set(source_ids_in_order[:args.limit])
        source_attachments = [item for item in source_attachments if item.source_id in allowed_source_ids]
    logger.info('泛微待处理: %s 个客商，%s 条附件 DOCID 关联', len({item.source_id for item in source_attachments}), len(source_attachments))

    all_targets: List[UatTarget] = []
    for target_type, config in TARGET_CONFIG.items():
        records = _fetch_all_uat_records(
            base_url, organization_id, config['endpoint'], access_token, page_size, timeout_seconds, logger,
        )
        targets = _build_uat_targets(target_type, records, args.allow_primary_key_unique_code)
        all_targets.extend(targets)
        logger.info('UAT %s: %s 条接口记录，%s 条含可匹配标识', target_type, len(records), len(targets))
    identifier_index = _build_identifier_index(all_targets)

    matched_sources: List[SourceAttachment] = []
    match_results: Dict[SourceAttachment, MatchOutcome] = {}
    for source in source_attachments:
        outcome = _match_targets(source, identifier_index)
        match_results[source] = outcome
        if outcome.matches:
            matched_sources.append(source)
    attachment_metadata = _load_attachment_metadata(item.docid for item in matched_sources)

    existing_names_cache: Dict[Tuple[str, str, str], Set[str]] = {}
    status_counts: Counter = Counter()
    with open(result_file, 'w', encoding='utf-8-sig', newline='') as result_handle:
        writer = csv.DictWriter(result_handle, fieldnames=RESULT_COLUMNS)
        writer.writeheader()
        for source in source_attachments:
            outcome = match_results[source]
            if not outcome.matches:
                message = (
                    '未找到 UAT 客商，未上传'
                    if outcome.reason == 'identifier_not_found'
                    else '同一客商类型匹配到多条 UAT 客商，未上传'
                )
                row = _result_row(source, outcome.reason, message)
                _write_result(writer, logger, row)
                status_counts[row['status']] += 1
                continue

            if outcome.reason == 'identifier_ambiguous_partial':
                ambiguous_labels = '、'.join(
                    '供应商' if target_type == 'vender' else '客户'
                    for target_type in outcome.ambiguous_target_types
                )
                row = _result_row(
                    source,
                    outcome.reason,
                    f'{ambiguous_labels}匹配到多条 UAT 客商，已跳过该类型；其余唯一类型继续处理',
                )
                _write_result(writer, logger, row)
                status_counts[row['status']] += 1

            metas = attachment_metadata.get(source.docid, [])
            for match in outcome.matches:
                target = match.target
                source_identifier_type = match.source_identifier_type
                target_identifier_field = match.target_identifier_field
                if not target.unique_code:
                    row = _result_row(
                        source, 'target_attachment_key_missing',
                        'UAT 列表未返回 headerId；为避免附件挂错，未上传。可核实后使用 --allow-primary-key-unique-code。',
                        target, source_identifier_type, target_identifier_field,
                    )
                    _write_result(writer, logger, row)
                    status_counts[row['status']] += 1
                    continue

                if not metas:
                    row = _result_row(
                        source, 'missing_docimagefile', '泛微 docimagefile/imagefile 未找到附件元数据',
                        target, source_identifier_type, target_identifier_field,
                    )
                    _write_result(writer, logger, row)
                    status_counts[row['status']] += 1
                    continue

                cache_key = (target.attachment_table, target.unique_code, target.target_type)
                if cache_key not in existing_names_cache:
                    try:
                        existing_names_cache[cache_key] = _fetch_existing_attachment_names(
                            base_url, organization_id, access_token, target, page_size, timeout_seconds,
                        )
                    except UatRequestError as exc:
                        for meta in metas:
                            row = _result_row(
                                source, 'target_attachment_query_failed', str(exc), target,
                                source_identifier_type, target_identifier_field, meta, exc.status_code,
                            )
                            _write_result(writer, logger, row)
                            status_counts[row['status']] += 1
                        continue

                for meta in metas:
                    existing_names = existing_names_cache[cache_key]
                    upload_meta = meta
                    original_name = _normalized_file_name(meta.attachment_name)
                    if not args.no_skip_existing and original_name in existing_names:
                        renamed_name = _collision_renamed_attachment_name(meta)
                        renamed_normalized = _normalized_file_name(renamed_name)
                        if renamed_normalized in existing_names:
                            row = _result_row(
                                source,
                                'skipped_existing_renamed_name',
                                f'目标营业执照页已有同名附件，改名版本 {renamed_name} 也已存在',
                                target, source_identifier_type, target_identifier_field,
                                replace(meta, attachment_name=renamed_name),
                            )
                            _write_result(writer, logger, row)
                            status_counts[row['status']] += 1
                            continue
                        upload_meta = replace(meta, attachment_name=renamed_name)
                    if not args.execute:
                        row = _result_row(
                            source,
                            'matched_dry_run',
                            (
                                f'目标页已有同名附件，dry-run 将改名为 {upload_meta.attachment_name} 后上传'
                                if upload_meta != meta else '已唯一匹配，dry-run 未上传'
                            ),
                            target, source_identifier_type, target_identifier_field, upload_meta,
                        )
                        _write_result(writer, logger, row)
                        status_counts[row['status']] += 1
                        continue

                    try:
                        file_data = _download_weaver_attachment(cookie, meta, timeout_seconds, max_file_bytes)
                        status_code = _retry_upload(
                            base_url, organization_id, access_token, target, upload_meta, file_data,
                            timeout_seconds, upload_retries,
                        )
                        existing_names.add(_normalized_file_name(upload_meta.attachment_name))
                        row = _result_row(
                            source,
                            'uploaded_renamed' if upload_meta != meta else 'uploaded',
                            (
                                f'同名附件已改名为 {upload_meta.attachment_name} 后上传成功，{len(file_data)} bytes'
                                if upload_meta != meta else f'上传成功，{len(file_data)} bytes'
                            ),
                            target, source_identifier_type, target_identifier_field, upload_meta, status_code,
                        )
                    except UatRequestError as exc:
                        row = _result_row(
                            source, 'upload_failed', str(exc), target, source_identifier_type,
                            target_identifier_field, upload_meta, exc.status_code,
                        )
                    except Exception as exc:  # 单个附件失败不可中断全量迁移。
                        row = _result_row(
                            source, 'download_or_upload_failed', str(exc), target,
                            source_identifier_type, target_identifier_field, upload_meta,
                        )
                    _write_result(writer, logger, row)
                    status_counts[row['status']] += 1

    logger.info('同步完成: %s', dict(sorted(status_counts.items())))
    logger.info('结果清单: %s', result_file)
    logger.info('运行日志: %s', log_file)
    return result_file


if __name__ == '__main__':
    run()
