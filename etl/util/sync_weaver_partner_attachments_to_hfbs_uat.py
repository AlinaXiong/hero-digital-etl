# -*- coding: utf-8 -*-
"""泛微客商附件同步到业财 UAT 或生产环境。

从泛微 ``uf_khgys`` 读取客商证件号/纳税人识别号及以下附件字段，
精确匹配业财 UAT 客户、供应商后，上传到其「营业执照」附件页：

* gsjjdzb：公司简介电子版
* yyzzsmj：营业执照扫描件
* khxksmj：开户许可扫描件
* ztsmxianzhang、jtgssmxz：主体/集团扫描鲜章
* yqlcxz：印签留存鲜章
* gsqdxz：公司清单鲜章

默认目标为 UAT 且仅生成匹配日志（dry-run），必须显式传入 ``--execute`` 才会
下载泛微文件并上传。生产环境必须额外显式传入 ``--target-environment prod``，
并通过 SSH 只读查询 ``hfins_base`` 客商主数据；附件绑定使用主数据表的
``header_id``，不会误用 ``customer_id`` / ``vender_id``。仅对证件号/税号
精确匹配且明确处于启用状态的目标客商上传；
同一标识命中多个已启用客户或供应商时会全部上传，禁用记录、无匹配记录、
缺少附件关联键的记录会写入结果清单而不会上传。

传入 ``--ocr-only`` 时不上传附件，而是下载泛微 ``yyzzsmj`` 附件并调用生产
汉得营业执照 OCR；每个泛微客商的匹配、附件存在性和 OCR 结果会写入 SQLite。
``--execute`` 使用豆包及人工复核后的本地通过版本上传，不再重复调用生产 OCR。

传入 ``--download-business-licenses`` / ``--download-identity-cards`` 时，把
``yyzzsmj`` 营业执照 / ``sfzfyj`` 身份证复印件下载到本地，并写入 SQLite
豆包识别队列；企业豆包识别后可用
``--import-doubao-results <JSONL 文件>`` 将逐条 JSON 结果回写该队列。

运行前配置（项目根目录 .env.local 优先）：

    HFBS_UAT_ACCESS_TOKEN=<UAT 登录 Bearer Token，必填>
    HFBS_PROD_ACCESS_TOKEN=<生产登录 Bearer Token，生产目标必填>
    WEAVER_CONTRACT_ATTACHMENT_COOKIE=<泛微 Cookie，执行上传时必填>
    HFINS_PROD_OCR_ACCESS_TOKEN=<生产汉得 OCR Token，OCR 时必填>

可选配置：

    HFBS_UAT_BASE_URL=https://uat.link.heroesports.com/gtw/hfbs
    HFBS_PROD_BASE_URL=http://api.link.heroesports.com/hfbs
    HFBS_UAT_ORGANIZATION_ID=0
    HFBS_UAT_PAGE_SIZE=200
    HFBS_UAT_TIMEOUT_SECONDS=90
    HFBS_UAT_UPLOAD_RETRIES=3
    HFINS_PROD_OCR_URL=http://api.link.heroesports.com/hfins/v2/0/ocr/recognize/business-file

示例：

    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py
    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py --source-id 12345
    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py --execute
    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py --target-environment prod
    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py --target-environment prod --execute
    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py --ocr-only
    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py --download-business-licenses
    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py --download-identity-cards
    python etl/util/sync_weaver_partner_attachments_to_hfbs_uat.py --import-doubao-results doubao_result.jsonl

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
import sqlite3
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
DEFAULT_PROD_BASE_URL = 'http://api.link.heroesports.com/hfbs'
DEFAULT_UAT_ORGANIZATION_ID = '0'
DEFAULT_PAGE_SIZE = 200
DEFAULT_TIMEOUT_SECONDS = 90
DEFAULT_UPLOAD_RETRIES = 3
UAT_PDF_COMPRESSION_TRIGGER_BYTES = 50 * 1024 * 1024
UAT_PDF_COMPRESSION_TARGET_BYTES = 45 * 1024 * 1024
UAT_PDF_SAFE_UPLOAD_BYTES = 49 * 1024 * 1024
DEFAULT_PROD_OCR_URL = 'http://api.link.heroesports.com/hfins/v2/0/ocr/recognize/business-file'
UAT_TOKEN_ENV = 'HFBS_UAT_ACCESS_TOKEN'
UAT_BASE_URL_ENV = 'HFBS_UAT_BASE_URL'
PROD_TOKEN_ENV = 'HFBS_PROD_ACCESS_TOKEN'
PROD_BASE_URL_ENV = 'HFBS_PROD_BASE_URL'
UAT_ORGANIZATION_ID_ENV = 'HFBS_UAT_ORGANIZATION_ID'
PROD_ORGANIZATION_ID_ENV = 'HFBS_PROD_ORGANIZATION_ID'
UAT_PAGE_SIZE_ENV = 'HFBS_UAT_PAGE_SIZE'
UAT_TIMEOUT_ENV = 'HFBS_UAT_TIMEOUT_SECONDS'
UAT_UPLOAD_RETRIES_ENV = 'HFBS_UAT_UPLOAD_RETRIES'
PROD_OCR_TOKEN_ENV = 'HFINS_PROD_OCR_ACCESS_TOKEN'
PROD_OCR_URL_ENV = 'HFINS_PROD_OCR_URL'
OCR_AUDIT_DB_NAME = '泛微客商营业执照OCR结果.sqlite'
BUSINESS_LICENSE_FIELD = 'yyzzsmj'
BUSINESS_LICENSE_DOWNLOAD_DIR_NAME = '营业执照附件'
IDENTITY_CARD_FIELD = 'sfzfyj'
IDENTITY_CARD_DOWNLOAD_DIR_NAME = '身份证附件'
DOUBAO_QUEUE_TABLE = 'weaver_partner_business_license_doubao_queue'
OCR_CACHEABLE_STATUSES = frozenset({
    'business_license_ok',
    'business_license_problem',
    'business_license_identifier_mismatch',
})

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
IDENTITY_CARD_ATTACHMENT_FIELDS: Tuple[Tuple[str, str], ...] = (
    (IDENTITY_CARD_FIELD, '身份证复印件'),
)
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
WINDOWS_FILENAME_INVALID_RE = re.compile(r'[<>:"/\\|?*\x00-\x1F]')


@dataclass(frozen=True)
class SourceAttachment:
    source_id: str
    source_name: str
    identifier_values: Tuple[Tuple[str, str], ...]
    field_name: str
    field_label: str
    docid: int
    legal_representative: str = ''
    identity_number: str = ''


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
    """一条泛微附件与一个唯一目标环境客商的匹配结果。"""

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
    """目标环境 HTTP 调用失败，保留可写入日志的状态码和错误摘要。"""

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


def _detected_file_extension(data: bytes) -> str:
    """根据文件头补全泛微中缺失的扩展名，避免 UAT 拒绝无后缀文件名。"""
    if data.startswith(b'%PDF-'):
        return '.pdf'
    if data.startswith(b'\x89PNG\r\n\x1a\n'):
        return '.png'
    if data.startswith(b'\xff\xd8\xff'):
        return '.jpg'
    if data.startswith((b'GIF87a', b'GIF89a')):
        return '.gif'
    if data.startswith((b'II*\x00', b'MM\x00*')):
        return '.tif'
    if data.startswith(b'BM'):
        return '.bmp'
    if len(data) >= 12 and data.startswith(b'RIFF') and data[8:12] == b'WEBP':
        return '.webp'
    return ''


def _with_detected_file_extension(meta: AttachmentMeta, data: bytes) -> AttachmentMeta:
    """UAT 强制文件名带后缀；仅在原名无后缀且文件头可确认时补全。"""
    if Path(meta.attachment_name).suffix:
        return meta
    extension = _detected_file_extension(data)
    if not extension:
        return meta
    return replace(meta, attachment_name=f'{meta.attachment_name}{extension}')


def _compressed_pdf_name(filename: str) -> str:
    """为压缩副本生成稳定文件名，不覆盖 UAT 中可能存在的原文件。"""
    path = Path(filename)
    suffix = path.suffix or '.pdf'
    stem = path.stem if path.suffix else path.name
    if stem.endswith('_压缩版'):
        return f'{stem}{suffix}'
    return f'{stem}_压缩版{suffix}'


def _compress_pdf_for_uat(data: bytes) -> Tuple[bytes, str]:
    """把大 PDF 压到 UAT 50MB 单文件限制内；优先无损，必要时逐级栅格压缩。"""
    try:
        import pymupdf as fitz  # 按需加载，普通附件不增加启动成本。
    except ImportError as exc:
        raise RuntimeError('压缩大 PDF 需要 PyMuPDF，请先执行 pip install -r requirements.txt') from exc

    source_document = fitz.open(stream=data, filetype='pdf')
    try:
        if source_document.needs_pass:
            raise RuntimeError('PDF 已加密，无法自动压缩')
        if source_document.page_count <= 0:
            raise RuntimeError('PDF 没有可压缩页面')

        best_data = data
        best_method = '原文件'
        try:
            optimized = source_document.tobytes(
                garbage=4,
                clean=True,
                deflate=True,
                deflate_images=True,
                deflate_fonts=True,
            )
            if len(optimized) < len(best_data):
                best_data = optimized
                best_method = '无损整理'
            if len(best_data) <= UAT_PDF_COMPRESSION_TARGET_BYTES:
                return best_data, best_method
        except (RuntimeError, ValueError):
            # 部分结构异常的 PDF 无法无损重写，仍可尝试逐页栅格化生成副本。
            pass

        profiles = (
            (144, 82),
            (120, 78),
            (96, 72),
            (72, 65),
            (60, 55),
            (48, 45),
        )
        for dpi, quality in profiles:
            compressed_document = fitz.open()
            try:
                scale = dpi / 72.0
                matrix = fitz.Matrix(scale, scale)
                for page_number in range(source_document.page_count):
                    source_page = source_document.load_page(page_number)
                    pixmap = source_page.get_pixmap(
                        matrix=matrix,
                        colorspace=fitz.csRGB,
                        alpha=False,
                        annots=True,
                    )
                    jpeg_data = pixmap.tobytes('jpeg', jpg_quality=quality)
                    output_page = compressed_document.new_page(
                        width=source_page.rect.width,
                        height=source_page.rect.height,
                    )
                    output_page.insert_image(output_page.rect, stream=jpeg_data)
                candidate = compressed_document.tobytes(garbage=4, deflate=True)
            finally:
                compressed_document.close()

            if len(candidate) < len(best_data):
                best_data = candidate
                best_method = f'{dpi}dpi/JPEG{quality}'
            if len(best_data) <= UAT_PDF_COMPRESSION_TARGET_BYTES:
                return best_data, best_method

        if len(best_data) <= UAT_PDF_SAFE_UPLOAD_BYTES:
            return best_data, best_method
        raise RuntimeError(
            f'PDF 自动压缩后仍有 {len(best_data)} bytes，无法降到 UAT 50MB 限制以内'
        )
    finally:
        source_document.close()


def _normalized_file_name(value: str) -> str:
    """与 UAT 文件列表保持一致：文件名可能经过 URL 编码。"""
    return urllib.parse.unquote(_text(value)).casefold()


def _setup_logger(run_id: str) -> Tuple[logging.Logger, Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log_file = OUTPUT_DIR / f'同步日志_{run_id}.log'
    logger = logging.getLogger(f'{TASK_NAME}.{run_id}')
    logger.setLevel(logging.INFO)
    logger.propagate = False
    formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')

    # Windows 终端通常使用 GBK；国外附件名中的韩文、土耳其文、不间断空格等
    # 字符可能无法编码，导致每条进度日志都报 UnicodeEncodeError。保留终端原编码，
    # 仅对无法表示的字符做可读转义；UTF-8 日志文件仍保留完整原文件名。
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(errors='backslashreplace')
        except (AttributeError, ValueError):
            pass

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.handlers[:] = [console_handler, file_handler]
    return logger, log_file


def _load_successful_uploads(
    output_dir: Path,
    target_environment: str = 'uat',
) -> Dict[Tuple[str, str, str, int, int], Set[str]]:
    """从历史结果清单恢复成功上传记录，重复运行时避免再次上传同一目标附件。

    UAT 的 primaryId/headerId 是与当前登录 Token 关联的加密值，重新登录后会变化，
    不能作为跨次运行的去重键。这里使用稳定的客商类型 + 客商编码。
    """
    successful: Dict[Tuple[str, str, str, int, int], Set[str]] = defaultdict(set)
    for result_path in sorted(output_dir.glob('同步结果_*.csv')):
        is_prod_result = result_path.name.startswith('同步结果_prod_')
        if (target_environment == 'prod') != is_prod_result:
            continue
        try:
            with result_path.open(encoding='utf-8-sig', newline='') as result_handle:
                for row in csv.DictReader(result_handle):
                    if row.get('status') not in ('uploaded', 'uploaded_renamed'):
                        continue
                    try:
                        key = (
                            _text(row.get('target_type')),
                            _text(row.get('target_code')).casefold(),
                            _text(row.get('source_id')),
                            int(_text(row.get('source_docid'))),
                            int(_text(row.get('weaver_imagefileid'))),
                        )
                    except (TypeError, ValueError):
                        continue
                    if all(key[:3]) and key[3] > 0 and key[4] > 0:
                        uploaded_name = _text(row.get('attachment_name'))
                        if uploaded_name:
                            successful[key].add(uploaded_name)
        except (OSError, UnicodeError, csv.Error):
            continue
    return successful


def _validate_source_fields(include_identity_cards: bool = False) -> None:
    expected_fields = {
        'khmc': '企业名称',
        'sh': '税号',
        'khsh': '纳税人识别号',
    }
    expected_fields.update(ATTACHMENT_FIELD_LABELS)
    if include_identity_cards:
        expected_fields.update({
            IDENTITY_CARD_FIELD: '身份证复印件',
            'sfzh': '身份证号',
            'frdb': '法人代表',
        })
    c.validate_fw_fields('uf_khgys', {'': expected_fields})


def _load_source_attachments(
    source_ids: Sequence[str],
    include_identity_cards: bool = False,
) -> List[SourceAttachment]:
    source_attachment_fields = ATTACHMENT_FIELDS + (
        IDENTITY_CARD_ATTACHMENT_FIELDS if include_identity_cards else ()
    )
    attachment_selects = ', '.join(f'k.{field_name}' for field_name, _ in source_attachment_fields)
    attachment_conditions = ' OR '.join(
        f"COALESCE(TRIM(k.{field_name}), '') <> ''" for field_name, _ in source_attachment_fields
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
            SELECT k.id, k.khmc, k.sh, k.khsh, k.frdb, k.sfzh, {attachment_selects}
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
        for field_name, field_label in source_attachment_fields:
            for docid in _extract_docids(row.get(field_name)):
                attachments.append(
                    SourceAttachment(
                        source_id=_text(row.get('id')),
                        source_name=_text(row.get('khmc')),
                        identifier_values=identifiers,
                        field_name=field_name,
                        field_label=field_label,
                        docid=docid,
                        legal_representative=_text(row.get('frdb')),
                        identity_number=_text(row.get('sfzh')),
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
        raise UatRequestError(f'目标 API HTTP {exc.code}: {_response_text(exc.read())}', exc.code) from exc
    except urllib.error.URLError as exc:
        raise UatRequestError(f'目标 API 网络错误: {exc.reason}') from exc
    try:
        return json.loads(body.decode('utf-8')), status_code
    except json.JSONDecodeError as exc:
        raise UatRequestError(f'目标 API 返回非 JSON: {_response_text(body)}', status_code) from exc


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


def _uat_record_enabled(record: dict) -> bool:
    """只允许 UAT 明确标记为启用的客户/供应商进入附件匹配。"""
    value = record.get('enabledFlag')
    if isinstance(value, bool):
        return value
    return _text(value).casefold() in ('1', 'true', 'yes', 'y', '是', '启用', 'enabled')


def _build_uat_targets(
    target_type: str,
    records: Iterable[dict],
    allow_primary_key_unique_code: bool,
) -> List[UatTarget]:
    config = TARGET_CONFIG[target_type]
    targets: List[UatTarget] = []
    for record in records:
        if not _uat_record_enabled(record):
            continue
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


def _load_prod_targets_from_db(
    target_type: str,
    organization_id: str,
    logger: logging.Logger,
) -> List[UatTarget]:
    """通过 SSH 读取业财生产客户/供应商及页面实际使用的附件 ``header_id``。"""
    try:
        tenant_id = int(organization_id)
    except ValueError as exc:
        raise RuntimeError(f'生产环境组织 ID 必须是整数: {organization_id}') from exc

    if target_type == 'customer':
        sql = '''
            SELECT
                CAST(customer_id AS CHAR) AS customerId,
                CAST(header_id AS CHAR) AS headerId,
                customer_code AS customerCode,
                description,
                taxpayer_name AS taxpayerName,
                taxpayer_number AS taxpayerNumber,
                enabled_flag AS enabledFlag
            FROM hfbs_system_customer
            WHERE tenant_id = %s AND enabled_flag = 1
            ORDER BY customer_id
        '''
    elif target_type == 'vender':
        sql = '''
            SELECT
                CAST(vender_id AS CHAR) AS venderId,
                CAST(header_id AS CHAR) AS headerId,
                vender_code AS venderCode,
                description,
                taxpayer_name AS taxpayerName,
                tax_id_number AS taxIdNumber,
                taxpayer_number AS taxpayerNumber,
                enabled_flag AS enabledFlag
            FROM hfbs_system_vender
            WHERE tenant_id = %s AND enabled_flag = 1
            ORDER BY vender_id
        '''
    else:
        raise RuntimeError(f'不支持的生产客商类型: {target_type}')

    records = c.query_db('HAND', 'hfins_base', sql, [tenant_id]).to_dict('records')
    # 生产「营业执照」页绑定的是主数据表 header_id。customer_id/vender_id 与
    # header_id 并不相同，缺失时宁可跳过，也不能把附件挂到错误业务键上。
    targets = _build_uat_targets(target_type, records, allow_primary_key_unique_code=False)

    # 同一 tableName + uniqueCode 无法区分两个客商。生产数据若出现重复 header_id，
    # 禁止自动上传，避免一个客商的证件同时出现在另一个客商页面。
    unique_code_counts = Counter(target.unique_code for target in targets if target.unique_code)
    duplicate_unique_codes = {
        unique_code for unique_code, count in unique_code_counts.items() if count > 1
    }
    if duplicate_unique_codes:
        logger.error(
            '[PROD-DB] %s 存在 %s 个重复 header_id，相关客商将跳过上传: %s',
            target_type, len(duplicate_unique_codes), ','.join(sorted(duplicate_unique_codes)),
        )
        targets = [
            replace(target, unique_code='', unique_code_source='duplicate_headerId')
            if target.unique_code in duplicate_unique_codes else target
            for target in targets
        ]

    attachment_ready = sum(bool(target.unique_code) for target in targets)
    missing_header = sum(not target.unique_code and not target.unique_code_source for target in targets)
    duplicate_header = sum(target.unique_code_source == 'duplicate_headerId' for target in targets)
    logger.info(
        '[PROD-DB] %s: %s 条启用记录，%s 条含可匹配标识，%s 条附件键可用，'
        '%s 条缺少 header_id，%s 条命中重复 header_id 保护',
        target_type, len(records), len(targets), attachment_ready,
        missing_header, duplicate_header,
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
    """按标识精确匹配全部已启用客户/供应商；命中几个目标就处理几个。"""
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
    for target_type in TARGET_CONFIG:
        candidates = candidates_by_type.get(target_type, [])
        for target, source_field, target_field in candidates:
            matches.append(TargetMatch(target, source_field, target_field))

    if not matches:
        return MatchOutcome((), 'identifier_not_found')
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
) -> bytes:
    base_url = os.getenv(c.ATTACHMENT_BASE_URL_ENV, c.DEFAULT_ATTACHMENT_BASE_URL).rstrip('/')
    url = f'{base_url}/weaver/weaver.file.FileDownload?fileid={meta.imagefileid}'
    request = urllib.request.Request(url, headers=_build_weaver_headers(cookie, meta))
    try:
        with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
            final_url = response.geturl()
            content_type = response.headers.get('Content-Type', '')
            data = response.read()
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f'泛微下载 HTTP {exc.code}: {_response_text(exc.read())}') from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f'泛微下载网络错误: {exc.reason}') from exc

    if 'login' in final_url.lower():
        raise RuntimeError(f'泛微下载跳转登录页: {final_url}')
    if 'text/html' in content_type.lower() and not meta.attachment_name.lower().endswith(('.html', '.htm')):
        raise RuntimeError(f'泛微下载返回 HTML: {_response_text(data)}')
    if not data:
        raise RuntimeError('泛微下载返回空文件')
    return data


def _safe_local_path_component(value: str, fallback: str) -> str:
    value = WINDOWS_FILENAME_INVALID_RE.sub('_', _text(value)).strip().rstrip('.')
    return value or fallback


def _document_type(source: SourceAttachment) -> str:
    return 'identity_card' if source.field_name == IDENTITY_CARD_FIELD else 'business_license'


def _document_expected_name(source: SourceAttachment) -> str:
    if _document_type(source) == 'identity_card':
        # 泛微 uf_khgys.frdb 在现有身份证附件数据中为空，个人客商名称存于 khmc。
        return source.legal_representative or source.source_name
    return source.source_name


def _document_expected_identifier(source: SourceAttachment) -> str:
    return source.identity_number if _document_type(source) == 'identity_card' else _source_certificate_number(source)


def _is_mainland_business_license(expected_name: str, expected_identifier: str) -> bool:
    """与人工复核口径一致：中文主体且号码为18位统一社会信用代码。"""
    return bool(re.search(r'[\u4e00-\u9fff]', expected_name or '')) and bool(
        re.fullmatch(r'[0-9A-Z]{18}', (expected_identifier or '').strip().upper())
    )


def _document_review_passed(row: sqlite3.Row, result: dict) -> bool:
    """计算豆包及人工复核后的最终通过状态。"""
    if 'review_overall_match' in result:
        return bool(result.get('review_overall_match'))
    if not bool(result.get('is_target_document')):
        return False

    name_match = bool(result.get('name_match'))
    identifier_match = bool(result.get('identifier_match'))
    expected_name = _text(row['expected_name'])
    expected_identifier = _text(row['expected_identifier'])
    if row['document_type'] == 'business_license':
        if _is_mainland_business_license(expected_name, expected_identifier):
            return name_match and identifier_match
        # 境外官方营业、税务或公司登记材料，名称或登记号码任一可靠确认即可。
        return name_match or identifier_match
    # 个人客商在泛微中经常使用昵称、项目名称、个体户名称或公司名称，不能要求
    # source_name 与身份证法定姓名一致。身份证号码是唯一标识：存在预期号码时，
    # 号码精确核验通过即可确认归属；只有缺少号码时才退回使用姓名核验。
    if expected_identifier:
        return identifier_match
    if expected_name:
        return name_match
    return False


def _load_approved_document_files(db_path: Path) -> Dict[Tuple[str, int, int, str], Path]:
    """加载最终复核通过的营业执照/身份证，并固定到已人工查看的本地文件。"""
    if not db_path.is_file():
        raise RuntimeError(f'豆包复核 SQLite 不存在: {db_path}')

    approved: Dict[Tuple[str, int, int, str], Path] = {}
    missing_files: List[str] = []
    with sqlite3.connect(db_path) as connection:
        connection.row_factory = sqlite3.Row
        rows = connection.execute(
            f'''
            SELECT source_id, source_docid, weaver_imagefileid, document_type,
                   expected_name, expected_identifier, local_file_path, doubao_result_json
            FROM {DOUBAO_QUEUE_TABLE}
            WHERE doubao_status = 'completed'
              AND document_type IN ('business_license', 'identity_card')
            ORDER BY document_type, CAST(source_id AS INTEGER), source_docid, weaver_imagefileid
            '''
        ).fetchall()
        for row in rows:
            try:
                result = json.loads(_text(row['doubao_result_json']) or '{}')
            except json.JSONDecodeError as exc:
                raise RuntimeError(
                    f"豆包结果 JSON 无效: source={row['source_id']} doc={row['source_docid']} "
                    f"file={row['weaver_imagefileid']}"
                ) from exc
            if not isinstance(result, dict) or not _document_review_passed(row, result):
                continue
            local_path = Path(_text(row['local_file_path']))
            if not local_path.is_file() or local_path.stat().st_size <= 0:
                missing_files.append(
                    f"{row['document_type']}:{row['source_id']}:{row['source_docid']}:"
                    f"{row['weaver_imagefileid']}:{local_path}"
                )
                continue
            key = (
                _text(row['source_id']), int(row['source_docid']),
                int(row['weaver_imagefileid']), _text(row['document_type']),
            )
            approved[key] = local_path.resolve()

    if missing_files:
        preview = '\n'.join(missing_files[:10])
        raise RuntimeError(
            f'有 {len(missing_files)} 个审批通过附件缺少本地原文件，已中止上传。前10项:\n{preview}'
        )
    return approved


def _approved_document_key(
    source: SourceAttachment,
    meta: AttachmentMeta,
) -> Tuple[str, int, int, str]:
    return source.source_id, meta.docid, meta.imagefileid, _document_type(source)


def _read_approved_document(path: Path) -> bytes:
    file_size = path.stat().st_size
    if file_size <= 0:
        raise RuntimeError(f'审批通过附件为空文件: {path}')
    return path.read_bytes()


def _document_local_path(source: SourceAttachment, meta: AttachmentMeta) -> Path:
    source_dir = _safe_local_path_component(source.source_id, 'unknown_source')
    file_name = _safe_local_path_component(
        meta.attachment_name,
        f'weaver_doc_{meta.docid}_file_{meta.imagefileid}',
    )
    return (
        OUTPUT_DIR
        / (
            IDENTITY_CARD_DOWNLOAD_DIR_NAME
            if _document_type(source) == 'identity_card' else BUSINESS_LICENSE_DOWNLOAD_DIR_NAME
        )
        / source_dir
        / f'{meta.docid}_{meta.imagefileid}_{file_name}'
    )


def _download_document_to_local(
    source: SourceAttachment,
    meta: AttachmentMeta,
    cookie: str,
    timeout_seconds: int,
) -> Tuple[Path, str, int]:
    """幂等下载营业执照或身份证；文件地址固定，避免重复下载同一个 IMAGEFILEID。"""
    target_path = _document_local_path(source, meta)
    if target_path.is_file() and target_path.stat().st_size > 0:
        return target_path.resolve(), 'existing', target_path.stat().st_size

    data = _download_weaver_attachment(cookie, meta, timeout_seconds)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = target_path.with_name(f'{target_path.name}.part-{os.getpid()}')
    try:
        temp_path.write_bytes(data)
        temp_path.replace(target_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return target_path.resolve(), 'downloaded', len(data)


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


def _file_multipart_body(filename: str, data: bytes) -> Tuple[bytes, str]:
    """构造仅含 ``file`` 字段的 multipart 请求，用于生产营业执照 OCR。"""
    boundary = f'----HeroBusinessLicenseOcr{time.time_ns()}'
    mime_type = mimetypes.guess_type(filename)[0] or 'application/octet-stream'
    chunks = [
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


def _with_access_token_query(url: str, access_token: str) -> str:
    """生产 OCR 前端以 query access_token 鉴权；去重后再追加当前 Token。"""
    parsed = urllib.parse.urlsplit(url)
    query_items = [
        (key, value)
        for key, value in urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
        if key != 'access_token'
    ]
    query_items.append(('access_token', access_token))
    return urllib.parse.urlunsplit(parsed._replace(query=urllib.parse.urlencode(query_items)))


def _request_production_business_license_ocr(
    ocr_url: str,
    access_token: str,
    meta: AttachmentMeta,
    data: bytes,
    timeout_seconds: int,
) -> Tuple[object, int]:
    """调用生产汉得营业执照 OCR，不将 Token 或文件内容写入日志。"""
    request_url = _with_access_token_query(ocr_url, access_token)
    body, content_type = _file_multipart_body(meta.attachment_name, data)
    headers = _authorization_headers(access_token)
    headers['Content-Type'] = content_type
    headers['Content-Length'] = str(len(body))
    request = urllib.request.Request(request_url, data=body, headers=headers, method='POST')
    try:
        with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
            response_body = response.read()
            status_code = response.getcode()
    except urllib.error.HTTPError as exc:
        raise UatRequestError(f'生产 OCR HTTP {exc.code}: {_response_text(exc.read())}', exc.code) from exc
    except urllib.error.URLError as exc:
        raise UatRequestError(f'生产 OCR 网络错误: {exc.reason}') from exc

    try:
        payload = json.loads(response_body.decode('utf-8'))
    except json.JSONDecodeError as exc:
        raise UatRequestError(f'生产 OCR 返回非 JSON: {_response_text(response_body)}', status_code) from exc
    if isinstance(payload, dict) and payload.get('failed'):
        raise UatRequestError(
            _text(payload.get('message')) or _text(payload.get('detailsMessage')) or '生产 OCR 调用失败',
            status_code,
        )
    return payload, status_code


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


def _source_certificate_number(source: SourceAttachment) -> str:
    """保留源表证件号/税号原值，仅写入本机 SQLite 审计库，不写入运行日志。"""
    return ' | '.join(dict.fromkeys(value for _, value in source.identifier_values if _text(value)))


def _match_audit_values(outcome: MatchOutcome) -> Tuple[str, str, str]:
    match_result = outcome.reason or ('matched' if outcome.matches else 'identifier_not_found')
    target_summaries = []
    for match in outcome.matches:
        target = match.target
        target_summaries.append(
            f'{target.target_type}:{target.code or target.primary_id}:{target.name}'
        )
    if outcome.ambiguous_target_types:
        ambiguous = '、'.join(
            '供应商' if target_type == 'vender' else '客户'
            for target_type in outcome.ambiguous_target_types
        )
        detail = f'{ambiguous}存在多条匹配；唯一匹配目标：{"; ".join(target_summaries) or "无"}'
    elif target_summaries:
        detail = f'唯一匹配目标：{"; ".join(target_summaries)}'
    else:
        detail = (
            '未找到目标环境客商'
            if match_result == 'identifier_not_found'
            else '同一客商类型匹配到多条目标环境客商'
        )
    return match_result, detail, '; '.join(target_summaries)


def _ocr_items(payload: object) -> List[dict]:
    """兼容 HZero 网关直返列表及 ``data`` 包装两种 OCR 响应。"""
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    if isinstance(payload, dict):
        data = payload.get('data')
        if isinstance(data, list):
            return [item for item in data if isinstance(item, dict)]
        if isinstance(data, dict):
            return [data]
    return []


def _evaluate_business_license_ocr(payload: object, source: SourceAttachment) -> dict:
    """把生产 OCR 返回归类为正常、营业执照有问题或响应异常。"""
    items = _ocr_items(payload)
    item = next((candidate for candidate in items if _text(candidate.get('ocrType')) == 'BUSINESS_LICENSE'), None)
    if item is None and items:
        item = items[0]
    if item is None:
        return {
            'ocr_status': 'business_license_problem',
            'ocr_is_business_license': False,
            'ocr_identifier_match': 'not_available',
            'ocr_type': '',
            'ocr_message': '生产 OCR 未返回识别结果',
            'ocr_result_json': json.dumps(payload, ensure_ascii=False, default=str),
        }

    ocr_type = _text(item.get('ocrType'))
    result_info = item.get('resultInfo') if isinstance(item.get('resultInfo'), dict) else {}
    ocr_message = _text(item.get('message'))
    has_result_values = any(_text(value) for value in result_info.values())
    ocr_uscc = _text(result_info.get('uscc'))
    source_identifiers = {
        _normalize_identifier(value) for _, value in source.identifier_values if _normalize_identifier(value)
    }
    if ocr_uscc:
        ocr_identifier_match = (
            'matched_source_identifier'
            if _normalize_identifier(ocr_uscc) in source_identifiers
            else 'mismatched_source_identifier'
        )
    elif source_identifiers:
        ocr_identifier_match = 'ocr_identifier_not_returned'
    else:
        ocr_identifier_match = 'source_identifier_not_available'

    if ocr_type != 'BUSINESS_LICENSE' or not has_result_values:
        ocr_status = 'business_license_problem'
    elif ocr_identifier_match == 'mismatched_source_identifier':
        ocr_status = 'business_license_identifier_mismatch'
    else:
        ocr_status = 'business_license_ok'
    if not ocr_message and ocr_status != 'business_license_ok':
        ocr_message = f'ocrType={ocr_type or "未返回"}，未取得有效营业执照识别字段'
    return {
        'ocr_status': ocr_status,
        'ocr_is_business_license': ocr_type == 'BUSINESS_LICENSE' and has_result_values,
        'ocr_identifier_match': ocr_identifier_match,
        'ocr_type': ocr_type,
        'ocr_message': ocr_message,
        'ocr_result_json': json.dumps(payload, ensure_ascii=False, default=str),
    }


def _connect_sqlite_autocommit(db_path: Path) -> sqlite3.Connection:
    """打开自动提交连接，保证长时间下载过程中每条状态都已落盘。"""
    connection = sqlite3.connect(db_path, isolation_level=None)
    # 避免有人同时查看/导入结果时，短暂的 SQLite 锁直接导致写入失败。
    connection.execute('PRAGMA busy_timeout = 30000')
    return connection


def _init_ocr_audit_db(connection: sqlite3.Connection) -> None:
    connection.execute(
        '''
        CREATE TABLE IF NOT EXISTS weaver_partner_business_license_ocr_result (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id TEXT NOT NULL,
            called_at TEXT NOT NULL,
            source_id TEXT NOT NULL,
            source_name TEXT,
            certificate_number TEXT,
            match_result TEXT NOT NULL,
            match_detail TEXT,
            matched_targets TEXT,
            has_business_license INTEGER NOT NULL,
            source_docid INTEGER NOT NULL DEFAULT 0,
            weaver_imagefileid INTEGER NOT NULL DEFAULT 0,
            attachment_name TEXT,
            ocr_called INTEGER NOT NULL,
            ocr_cache_hit INTEGER NOT NULL DEFAULT 0,
            ocr_status TEXT NOT NULL,
            ocr_is_business_license INTEGER,
            ocr_identifier_match TEXT,
            ocr_type TEXT,
            ocr_message TEXT,
            ocr_result_json TEXT,
            ocr_http_status INTEGER
        )
        '''
    )
    columns = {
        _text(row[1])
        for row in connection.execute('PRAGMA table_info(weaver_partner_business_license_ocr_result)')
    }
    if 'ocr_cache_hit' not in columns:
        connection.execute(
            'ALTER TABLE weaver_partner_business_license_ocr_result '
            'ADD COLUMN ocr_cache_hit INTEGER NOT NULL DEFAULT 0'
        )
    connection.execute(
        '''
        CREATE INDEX IF NOT EXISTS idx_weaver_partner_license_ocr_source
        ON weaver_partner_business_license_ocr_result (source_id, called_at)
        '''
    )


def _init_doubao_queue_db(connection: sqlite3.Connection) -> None:
    connection.execute(
        f'''
        CREATE TABLE IF NOT EXISTS {DOUBAO_QUEUE_TABLE} (
            source_id TEXT NOT NULL,
            source_docid INTEGER NOT NULL,
            weaver_imagefileid INTEGER NOT NULL,
            document_type TEXT NOT NULL DEFAULT 'business_license',
            source_name TEXT,
            certificate_number TEXT,
            expected_name TEXT,
            expected_identifier TEXT,
            match_result TEXT NOT NULL,
            match_detail TEXT,
            matched_targets TEXT,
            attachment_name TEXT,
            local_file_path TEXT,
            download_status TEXT NOT NULL,
            file_size INTEGER,
            doubao_status TEXT NOT NULL DEFAULT 'pending',
            doubao_is_business_license INTEGER,
            doubao_is_identity_card INTEGER,
            doubao_message TEXT,
            doubao_result_json TEXT,
            doubao_updated_at TEXT,
            PRIMARY KEY (source_id, source_docid, weaver_imagefileid)
        )
        '''
    )
    columns = {
        _text(row[1])
        for row in connection.execute(f'PRAGMA table_info({DOUBAO_QUEUE_TABLE})')
    }
    migrations = {
        'document_type': "TEXT NOT NULL DEFAULT 'business_license'",
        'expected_name': 'TEXT',
        'expected_identifier': 'TEXT',
        'local_file_path': 'TEXT',
        'doubao_is_identity_card': 'INTEGER',
    }
    for column_name, column_definition in migrations.items():
        if column_name not in columns:
            connection.execute(
                f'ALTER TABLE {DOUBAO_QUEUE_TABLE} ADD COLUMN {column_name} {column_definition}'
            )
    if 'business_license_local_path' in columns:
        connection.execute(
            f'''
            UPDATE {DOUBAO_QUEUE_TABLE}
            SET local_file_path = business_license_local_path
            WHERE COALESCE(local_file_path, '') = ''
              AND COALESCE(business_license_local_path, '') <> ''
            '''
        )
    connection.execute(
        f'''
        CREATE INDEX IF NOT EXISTS idx_weaver_partner_doubao_queue_status
        ON {DOUBAO_QUEUE_TABLE} (doubao_status, source_id)
        '''
    )


def _upsert_doubao_queue_row(
    connection: sqlite3.Connection,
    source: SourceAttachment,
    outcome: MatchOutcome,
    meta: Optional[AttachmentMeta],
    download_status: str,
    local_path: Optional[Path] = None,
    file_size: Optional[int] = None,
) -> None:
    match_result, match_detail, matched_targets = _match_audit_values(outcome)
    source_docid = meta.docid if meta else source.docid
    imagefileid = meta.imagefileid if meta else 0
    attachment_name = meta.attachment_name if meta else ''
    connection.execute(
        f'''
        INSERT INTO {DOUBAO_QUEUE_TABLE} (
            source_id, source_docid, weaver_imagefileid, document_type,
            source_name, certificate_number, expected_name, expected_identifier,
            match_result, match_detail, matched_targets, attachment_name,
            local_file_path, download_status, file_size, doubao_status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(source_id, source_docid, weaver_imagefileid) DO UPDATE SET
            document_type = excluded.document_type,
            source_name = excluded.source_name,
            certificate_number = excluded.certificate_number,
            expected_name = excluded.expected_name,
            expected_identifier = excluded.expected_identifier,
            match_result = excluded.match_result,
            match_detail = excluded.match_detail,
            matched_targets = excluded.matched_targets,
            attachment_name = excluded.attachment_name,
            local_file_path = excluded.local_file_path,
            download_status = excluded.download_status,
            file_size = excluded.file_size,
            doubao_status = CASE
                WHEN excluded.download_status = 'download_failed' THEN 'download_failed'
                WHEN {DOUBAO_QUEUE_TABLE}.doubao_status = 'completed' THEN 'completed'
                ELSE 'pending'
            END
        ''',
        (
            source.source_id,
            source_docid,
            imagefileid,
            _document_type(source),
            source.source_name,
            _source_certificate_number(source),
            _document_expected_name(source),
            _document_expected_identifier(source),
            match_result,
            match_detail,
            matched_targets,
            attachment_name,
            str(local_path) if local_path else '',
            download_status,
            file_size,
            'download_failed' if download_status == 'download_failed' else 'pending',
        ),
    )


def _optional_bool(value) -> Optional[bool]:
    if value is None or _text(value) == '':
        return None
    if isinstance(value, bool):
        return value
    normalized = _text(value).lower()
    if normalized in ('1', 'true', 'yes', 'y', '是'):
        return True
    if normalized in ('0', 'false', 'no', 'n', '否'):
        return False
    raise ValueError(f'无法识别布尔值: {value}')


def _import_doubao_results(result_file: Path, logger: logging.Logger) -> Path:
    """导入企业豆包逐行 JSON 识别结果，按泛微附件唯一键更新本地 SQLite 队列。"""
    if not result_file.is_file():
        raise RuntimeError(f'豆包结果文件不存在: {result_file}')
    db_path = OUTPUT_DIR / OCR_AUDIT_DB_NAME
    imported = 0
    skipped = 0
    with _connect_sqlite_autocommit(db_path) as connection:
        _init_doubao_queue_db(connection)
        for line_number, line in enumerate(result_file.read_text(encoding='utf-8-sig').splitlines(), start=1):
            if not line.strip():
                continue
            try:
                payload = json.loads(line)
                if not isinstance(payload, dict):
                    raise ValueError('每行必须是 JSON 对象')
                source_id = _text(payload.get('source_id'))
                source_docid = int(payload.get('source_docid'))
                imagefileid = int(payload.get('weaver_imagefileid'))
                if not source_id or source_docid <= 0 or imagefileid <= 0:
                    raise ValueError('缺少 source_id/source_docid/weaver_imagefileid')
                status = _text(payload.get('doubao_status')) or 'completed'
                is_business_license = _optional_bool(payload.get('is_business_license'))
                is_identity_card = _optional_bool(payload.get('is_identity_card'))
                message = _text(payload.get('message'))
                result = payload.get('result', payload.get('ocr_result'))
                result_json = json.dumps(result, ensure_ascii=False, default=str) if result is not None else ''
                cursor = connection.execute(
                    f'''
                    UPDATE {DOUBAO_QUEUE_TABLE}
                    SET doubao_status = ?, doubao_is_business_license = ?, doubao_is_identity_card = ?, doubao_message = ?,
                        doubao_result_json = ?, doubao_updated_at = ?
                    WHERE source_id = ? AND source_docid = ? AND weaver_imagefileid = ?
                    ''',
                    (
                        status,
                        None if is_business_license is None else int(is_business_license),
                        None if is_identity_card is None else int(is_identity_card),
                        message,
                        result_json,
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        source_id,
                        source_docid,
                        imagefileid,
                    ),
                )
                if cursor.rowcount != 1:
                    raise ValueError('SQLite 队列中未找到对应的附件')
                imported += 1
            except Exception as exc:
                skipped += 1
                logger.warning('[doubao_import_skipped] line=%s %s', line_number, exc)
    logger.info('豆包结果导入完成: imported=%s skipped=%s sqlite=%s', imported, skipped, db_path)
    return db_path


def _write_ocr_audit_row(
    connection: sqlite3.Connection,
    run_id: str,
    source: SourceAttachment,
    outcome: MatchOutcome,
    has_business_license: bool,
    ocr_record: dict,
    meta: Optional[AttachmentMeta] = None,
    source_docid: int = 0,
    http_status: Optional[int] = None,
) -> None:
    match_result, match_detail, matched_targets = _match_audit_values(outcome)
    connection.execute(
        '''
        INSERT INTO weaver_partner_business_license_ocr_result (
            run_id, called_at, source_id, source_name, certificate_number,
            match_result, match_detail, matched_targets, has_business_license,
            source_docid, weaver_imagefileid, attachment_name, ocr_called,
            ocr_cache_hit, ocr_status, ocr_is_business_license, ocr_identifier_match, ocr_type,
            ocr_message, ocr_result_json, ocr_http_status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''',
        (
            run_id,
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            source.source_id,
            source.source_name,
            _source_certificate_number(source),
            match_result,
            match_detail,
            matched_targets,
            int(has_business_license),
            meta.docid if meta else source_docid,
            meta.imagefileid if meta else 0,
            meta.attachment_name if meta else '',
            int(bool(ocr_record.get('ocr_called'))),
            int(bool(ocr_record.get('ocr_cache_hit'))),
            _text(ocr_record.get('ocr_status')),
            (
                int(bool(ocr_record['ocr_is_business_license']))
                if ocr_record.get('ocr_is_business_license') is not None else None
            ),
            _text(ocr_record.get('ocr_identifier_match')),
            _text(ocr_record.get('ocr_type')),
            _text(ocr_record.get('ocr_message')),
            _text(ocr_record.get('ocr_result_json')),
            http_status,
        ),
    )


def _load_cached_business_license_ocr(
    connection: sqlite3.Connection,
    source: SourceAttachment,
    meta: AttachmentMeta,
) -> Optional[dict]:
    """同一泛微 IMAGEFILEID 已有有效结论时复用，避免重复产生生产 OCR 费用。"""
    placeholders = ', '.join('?' for _ in OCR_CACHEABLE_STATUSES)
    row = connection.execute(
        f'''
        SELECT ocr_status, ocr_is_business_license, ocr_identifier_match,
               ocr_type, ocr_message, ocr_result_json
        FROM weaver_partner_business_license_ocr_result
        WHERE source_id = ?
          AND source_docid = ?
          AND weaver_imagefileid = ?
          AND ocr_status IN ({placeholders})
        ORDER BY id DESC
        LIMIT 1
        ''',
        (source.source_id, meta.docid, meta.imagefileid, *sorted(OCR_CACHEABLE_STATUSES)),
    ).fetchone()
    if row is None:
        return None
    return {
        'ocr_called': False,
        'ocr_cache_hit': True,
        'ocr_status': _text(row[0]),
        'ocr_is_business_license': None if row[1] is None else bool(row[1]),
        'ocr_identifier_match': _text(row[2]),
        'ocr_type': _text(row[3]),
        'ocr_message': _text(row[4]),
        'ocr_result_json': _text(row[5]),
    }


def _get_business_license_ocr_result(
    connection: sqlite3.Connection,
    source: SourceAttachment,
    meta: AttachmentMeta,
    ocr_enabled: bool,
    ocr_url: str,
    ocr_access_token: str,
    cookie: str,
    timeout_seconds: int,
) -> Tuple[dict, Optional[int], Optional[bytes]]:
    """取缓存或调用 OCR；新下载的内存 bytes 返回给同轮 UAT 上传复用。"""
    if not ocr_enabled:
        return (
            {
                'ocr_called': False,
                'ocr_status': 'ocr_not_called',
                'ocr_message': 'dry-run 未调用生产 OCR；请使用 --ocr-only 或 --execute',
            },
            None,
            None,
        )

    cached = _load_cached_business_license_ocr(connection, source, meta)
    if cached is not None:
        return cached, None, None

    try:
        file_data = _download_weaver_attachment(cookie, meta, timeout_seconds)
    except Exception as exc:  # 单个附件异常不可中断整批 OCR。
        return (
            {
                'ocr_called': True,
                'ocr_status': 'ocr_download_or_call_failed',
                'ocr_message': str(exc),
            },
            None,
            None,
        )
    try:
        ocr_payload, http_status = _request_production_business_license_ocr(
            ocr_url, ocr_access_token, meta, file_data, timeout_seconds,
        )
        ocr_record = _evaluate_business_license_ocr(ocr_payload, source)
        ocr_record['ocr_called'] = True
        return ocr_record, http_status, file_data
    except UatRequestError as exc:
        return (
            {
                'ocr_called': True,
                'ocr_status': 'ocr_call_failed',
                'ocr_message': str(exc),
            },
            exc.status_code,
            file_data,
        )
    except Exception as exc:  # 单个附件异常不可中断整批 OCR。
        return (
            {
                'ocr_called': True,
                'ocr_status': 'ocr_download_or_call_failed',
                'ocr_message': str(exc),
            },
            None,
            file_data,
        )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description='泛微客商附件同步到业财 UAT/PROD 营业执照页')
    parser.add_argument(
        '--target-environment', choices=('uat', 'prod'), default='uat',
        help='目标环境；默认 uat。生产上传必须显式指定 prod',
    )
    parser.add_argument(
        '--source-id', action='append', default=[],
        help='仅处理指定泛微 uf_khgys.id；可重复传入，例如 --source-id 1 --source-id 2',
    )
    parser.add_argument('--limit', type=int, default=0, help='最多处理多少个泛微客商 ID（0 表示全部）')
    parser.add_argument(
        '--target-type', action='append', choices=tuple(TARGET_CONFIG), default=[],
        help='仅处理指定目标客商类型；可重复传入 customer/vender，默认两种都处理',
    )
    parser.add_argument('--execute', action='store_true', help='上传审批通过的证件及其他泛微附件；默认仅 dry-run')
    parser.add_argument(
        '--ocr-only', action='store_true',
        help='仅调用生产汉得 OCR 校验 yyzzsmj 营业执照附件，不上传 UAT',
    )
    parser.add_argument(
        '--download-business-licenses', action='store_true',
        help='仅下载 yyzzsmj 营业执照到本地，并写入 SQLite 豆包识别队列，不上传 UAT、不调用汉得 OCR',
    )
    parser.add_argument(
        '--download-identity-cards', action='store_true',
        help='仅下载 sfzfyj 身份证复印件到本地，并写入 SQLite 豆包识别队列，不上传 UAT、不调用汉得 OCR',
    )
    parser.add_argument(
        '--import-doubao-results', type=Path,
        help='导入企业豆包逐行 JSON（JSONL）识别结果并回写 SQLite 队列',
    )
    parser.add_argument(
        '--allow-primary-key-unique-code', action='store_true',
        help='仅限 UAT：列表未返回 headerId 时允许使用 customerId/venderId 作为附件 uniqueCode（生产始终禁止）',
    )
    parser.add_argument(
        '--no-skip-existing', action='store_true',
        help='目标页已有同名附件时仍使用原文件名上传（默认直接跳过）',
    )
    return parser


def run(argv: Optional[Sequence[str]] = None) -> Path:
    args = build_parser().parse_args(argv)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    run_id = f'prod_{timestamp}' if args.target_environment == 'prod' else timestamp
    logger, log_file = _setup_logger(run_id)
    if args.import_doubao_results:
        if (
            args.execute or args.ocr_only or args.download_business_licenses
            or args.download_identity_cards or args.source_id or args.limit or args.target_type
        ):
            raise RuntimeError('--import-doubao-results 不能与下载、OCR、上传或筛选参数同时使用。')
        return _import_doubao_results(args.import_doubao_results, logger)
    if args.execute and args.ocr_only:
        raise RuntimeError('--execute 与 --ocr-only 不能同时使用。')
    if (args.download_business_licenses or args.download_identity_cards) and (args.execute or args.ocr_only):
        raise RuntimeError(
            '--download-business-licenses / --download-identity-cards 仅用于生成豆包离线识别队列，'
            '不能与 --execute 或 --ocr-only 同时使用。'
        )
    result_file = OUTPUT_DIR / f'同步结果_{run_id}.csv'
    ocr_audit_db = OUTPUT_DIR / OCR_AUDIT_DB_NAME
    target_environment = args.target_environment
    access_token_env = PROD_TOKEN_ENV if target_environment == 'prod' else UAT_TOKEN_ENV
    access_token = os.getenv(access_token_env, '').strip()
    # dry-run 也会查询目标页现有附件，以准确判断是否跳过，因此始终需要对应环境 Token。
    if not access_token:
        raise RuntimeError(f'缺少 {access_token_env}；请配置目标环境 Bearer Token 后重试。')
    cookie = os.getenv(c.ATTACHMENT_COOKIE_ENV, '').strip()
    ocr_enabled = args.ocr_only
    business_license_download_enabled = args.download_business_licenses
    identity_card_download_enabled = args.download_identity_cards
    document_download_enabled = business_license_download_enabled or identity_card_download_enabled
    upload_planning_enabled = not args.ocr_only and not document_download_enabled
    include_identity_cards = identity_card_download_enabled or upload_planning_enabled
    if (args.execute or ocr_enabled or document_download_enabled) and not cookie:
        raise RuntimeError(f'缺少 {c.ATTACHMENT_COOKIE_ENV}；下载附件或 OCR 前必须配置有效泛微 Cookie。')
    ocr_access_token = os.getenv(PROD_OCR_TOKEN_ENV, '').strip()
    if ocr_enabled and not ocr_access_token:
        raise RuntimeError(f'缺少 {PROD_OCR_TOKEN_ENV}；生产 OCR 校验前必须配置有效 Token。')

    if target_environment == 'prod':
        base_url = os.getenv(PROD_BASE_URL_ENV, DEFAULT_PROD_BASE_URL).strip().rstrip('/')
        organization_id = os.getenv(
            PROD_ORGANIZATION_ID_ENV, DEFAULT_UAT_ORGANIZATION_ID,
        ).strip()
    else:
        base_url = os.getenv(UAT_BASE_URL_ENV, DEFAULT_UAT_BASE_URL).strip().rstrip('/')
        organization_id = os.getenv(UAT_ORGANIZATION_ID_ENV, DEFAULT_UAT_ORGANIZATION_ID).strip()
    page_size = _positive_int_env(UAT_PAGE_SIZE_ENV, DEFAULT_PAGE_SIZE, maximum=500)
    timeout_seconds = _positive_int_env(UAT_TIMEOUT_ENV, DEFAULT_TIMEOUT_SECONDS, maximum=300)
    upload_retries = _positive_int_env(UAT_UPLOAD_RETRIES_ENV, DEFAULT_UPLOAD_RETRIES, maximum=10)
    ocr_url = os.getenv(PROD_OCR_URL_ENV, DEFAULT_PROD_OCR_URL).strip() or DEFAULT_PROD_OCR_URL
    logger.info(
        '开始同步: environment=%s mode=%s ocr=%s base_url=%s org=%s page_size=%s source_ids=%s',
        target_environment,
        (
            'execute' if args.execute else (
                'ocr-only' if args.ocr_only else (
                    (
                        'download-business-licenses+identity-cards'
                        if business_license_download_enabled and identity_card_download_enabled
                        else (
                            'download-business-licenses'
                            if business_license_download_enabled
                            else ('download-identity-cards' if identity_card_download_enabled else 'dry-run')
                        )
                    )
                )
            )
        ),
        'enabled' if ocr_enabled else 'not-called', base_url, organization_id, page_size,
        ','.join(args.source_id) if args.source_id else 'all',
    )

    approved_document_files: Dict[Tuple[str, int, int, str], Path] = {}
    if upload_planning_enabled:
        approved_document_files = _load_approved_document_files(ocr_audit_db)
        approved_counts = Counter(key[3] for key in approved_document_files)
        logger.info(
            '审批通过附件: 营业执照=%s 身份证=%s；其他企业附件不做内容校验、直接进入匹配上传流程',
            approved_counts['business_license'], approved_counts['identity_card'],
        )

    _validate_source_fields(include_identity_cards=include_identity_cards)
    source_attachments = _load_source_attachments(
        args.source_id,
        include_identity_cards=include_identity_cards,
    )
    source_ids_in_order = list(dict.fromkeys(item.source_id for item in source_attachments))
    if args.limit > 0:
        allowed_source_ids = set(source_ids_in_order[:args.limit])
        source_attachments = [item for item in source_attachments if item.source_id in allowed_source_ids]
    logger.info('泛微待处理: %s 个客商，%s 条附件 DOCID 关联', len({item.source_id for item in source_attachments}), len(source_attachments))

    all_targets: List[UatTarget] = []
    selected_target_types = set(args.target_type) if args.target_type else set(TARGET_CONFIG)
    for target_type, config in TARGET_CONFIG.items():
        if target_type not in selected_target_types:
            continue
        if target_environment == 'prod':
            targets = _load_prod_targets_from_db(target_type, organization_id, logger)
        else:
            records = _fetch_all_uat_records(
                base_url, organization_id, config['endpoint'], access_token,
                page_size, timeout_seconds, logger,
            )
            targets = _build_uat_targets(
                target_type, records, args.allow_primary_key_unique_code,
            )
            enabled_records = sum(_uat_record_enabled(record) for record in records)
            logger.info(
                'UAT %s: %s 条接口记录，%s 条启用，%s 条启用且含可匹配标识',
                target_type, len(records), enabled_records, len(targets),
            )
        all_targets.extend(targets)
    identifier_index = _build_identifier_index(all_targets)

    matched_sources: List[SourceAttachment] = []
    match_results: Dict[SourceAttachment, MatchOutcome] = {}
    sources_by_id: Dict[str, List[SourceAttachment]] = defaultdict(list)
    business_license_sources_by_id: Dict[str, List[SourceAttachment]] = defaultdict(list)
    identity_card_sources_by_id: Dict[str, List[SourceAttachment]] = defaultdict(list)
    for source in source_attachments:
        outcome = _match_targets(source, identifier_index)
        match_results[source] = outcome
        sources_by_id[source.source_id].append(source)
        if source.field_name == BUSINESS_LICENSE_FIELD:
            business_license_sources_by_id[source.source_id].append(source)
        if source.field_name == IDENTITY_CARD_FIELD:
            identity_card_sources_by_id[source.source_id].append(source)
        if outcome.matches:
            matched_sources.append(source)
    metadata_docids = {item.docid for item in matched_sources}
    metadata_docids.update(
        item.docid
        for item in source_attachments
        if item.field_name in (BUSINESS_LICENSE_FIELD, IDENTITY_CARD_FIELD)
    )
    attachment_metadata = _load_attachment_metadata(metadata_docids)

    planned_upload_actions = 0
    if upload_planning_enabled:
        for source in source_attachments:
            planned_metas = attachment_metadata.get(source.docid, [])
            if source.field_name in (BUSINESS_LICENSE_FIELD, IDENTITY_CARD_FIELD):
                planned_metas = [
                    meta for meta in planned_metas
                    if _approved_document_key(source, meta) in approved_document_files
                ]
            uploadable_targets = sum(
                bool(match.target.unique_code) for match in match_results[source].matches
            )
            planned_upload_actions += len(planned_metas) * uploadable_targets
        logger.info(
            '附件处理进度总数=%s（已启用目标客商 × 可上传附件；包含随后可能因已存在而跳过的项目）',
            planned_upload_actions,
        )

    ocr_status_counts: Counter = Counter()
    with _connect_sqlite_autocommit(ocr_audit_db) as ocr_connection:
        _init_ocr_audit_db(ocr_connection)
        _init_doubao_queue_db(ocr_connection)
        for source_id, source_items in sources_by_id.items():
            source = source_items[0]
            outcome = match_results[source]
            business_license_sources = business_license_sources_by_id.get(source_id, [])
            identity_card_sources = identity_card_sources_by_id.get(source_id, [])
            if not business_license_sources:
                _write_ocr_audit_row(
                    ocr_connection,
                    run_id,
                    source,
                    outcome,
                    has_business_license=False,
                    ocr_record={
                        'ocr_called': False,
                        'ocr_status': 'no_business_license',
                        'ocr_message': '泛微 yyzzsmj 字段未找到营业执照附件',
                    },
                )
                ocr_status_counts['no_business_license'] += 1
                if not identity_card_download_enabled or not identity_card_sources:
                    continue

            # --execute 时在后续逐个附件上传前 OCR，才能复用同一份内存 bytes。
            if args.execute:
                continue

            for business_license_source in business_license_sources:
                license_metas = attachment_metadata.get(business_license_source.docid, [])
                if not license_metas:
                    if business_license_download_enabled:
                        _upsert_doubao_queue_row(
                            ocr_connection,
                            business_license_source,
                            outcome,
                            meta=None,
                            download_status='missing_docimagefile',
                        )
                    _write_ocr_audit_row(
                        ocr_connection,
                        run_id,
                        business_license_source,
                        outcome,
                        has_business_license=True,
                        source_docid=business_license_source.docid,
                        ocr_record={
                            'ocr_called': False,
                            'ocr_status': 'missing_docimagefile',
                            'ocr_message': '泛微 docimagefile/imagefile 未找到营业执照附件元数据',
                        },
                    )
                    ocr_status_counts['missing_docimagefile'] += 1
                    continue

                for meta in license_metas:
                    if business_license_download_enabled:
                        _upsert_doubao_queue_row(
                            ocr_connection,
                            business_license_source,
                            outcome,
                            meta,
                            download_status='downloading',
                        )
                        try:
                            local_path, download_status, file_size = _download_document_to_local(
                                business_license_source,
                                meta,
                                cookie,
                                timeout_seconds,
                            )
                            logger.info(
                                '[business_license_download:%s] source=%s doc=%s file=%s path=%s',
                                download_status,
                                business_license_source.source_id,
                                meta.docid,
                                meta.attachment_name,
                                local_path,
                            )
                        except Exception as exc:  # 单个下载失败不影响其余营业执照。
                            local_path = None
                            download_status = 'download_failed'
                            file_size = None
                            logger.warning(
                                '[business_license_download_failed] source=%s doc=%s file=%s %s',
                                business_license_source.source_id,
                                meta.docid,
                                meta.attachment_name,
                                exc,
                            )
                        _upsert_doubao_queue_row(
                            ocr_connection,
                            business_license_source,
                            outcome,
                            meta,
                            download_status,
                            local_path,
                            file_size,
                        )
                    ocr_record, http_status, _ = _get_business_license_ocr_result(
                        ocr_connection,
                        business_license_source,
                        meta,
                        ocr_enabled,
                        ocr_url,
                        ocr_access_token,
                        cookie,
                        timeout_seconds,
                    )
                    _write_ocr_audit_row(
                        ocr_connection,
                        run_id,
                        business_license_source,
                        outcome,
                        has_business_license=True,
                        meta=meta,
                        ocr_record=ocr_record,
                        http_status=http_status,
                    )
                    ocr_status_counts[_text(ocr_record.get('ocr_status'))] += 1
                    logger.info(
                        '[OCR:%s] source=%s doc=%s file=%s',
                        ocr_record.get('ocr_status'),
                        business_license_source.source_id,
                        meta.docid,
                        meta.attachment_name,
                    )

            if not identity_card_download_enabled:
                continue

            for identity_card_source in identity_card_sources:
                identity_metas = attachment_metadata.get(identity_card_source.docid, [])
                if not identity_metas:
                    _upsert_doubao_queue_row(
                        ocr_connection,
                        identity_card_source,
                        outcome,
                        meta=None,
                        download_status='missing_docimagefile',
                    )
                    logger.warning(
                        '[identity_card_metadata_missing] source=%s doc=%s 泛微 docimagefile/imagefile 未找到身份证附件元数据',
                        identity_card_source.source_id,
                        identity_card_source.docid,
                    )
                    continue

                for meta in identity_metas:
                    _upsert_doubao_queue_row(
                        ocr_connection,
                        identity_card_source,
                        outcome,
                        meta,
                        download_status='downloading',
                    )
                    try:
                        local_path, download_status, file_size = _download_document_to_local(
                            identity_card_source,
                            meta,
                            cookie,
                            timeout_seconds,
                        )
                        logger.info(
                            '[identity_card_download:%s] source=%s doc=%s file=%s path=%s',
                            download_status,
                            identity_card_source.source_id,
                            meta.docid,
                            meta.attachment_name,
                            local_path,
                        )
                    except Exception as exc:  # 单个下载失败不影响其余身份证附件。
                        local_path = None
                        download_status = 'download_failed'
                        file_size = None
                        logger.warning(
                            '[identity_card_download_failed] source=%s doc=%s file=%s %s',
                            identity_card_source.source_id,
                            meta.docid,
                            meta.attachment_name,
                            exc,
                        )
                    _upsert_doubao_queue_row(
                        ocr_connection,
                        identity_card_source,
                        outcome,
                        meta,
                        download_status,
                        local_path,
                        file_size,
                    )
    if not args.execute:
        logger.info('营业执照 OCR 审计完成: %s', dict(sorted(ocr_status_counts.items())))
        logger.info('营业执照 OCR SQLite: %s', ocr_audit_db)

    existing_names_cache: Dict[Tuple[str, str, str], Set[str]] = {}
    status_counts: Counter = Counter()
    execute_ocr_connection: Optional[sqlite3.Connection] = None
    if args.execute and ocr_enabled:
        execute_ocr_connection = _connect_sqlite_autocommit(ocr_audit_db)
        _init_ocr_audit_db(execute_ocr_connection)
    successful_uploads = _load_successful_uploads(OUTPUT_DIR, target_environment)
    logger.info(
        '%s 历史成功上传记录=%s 条；重复运行将先与目标环境现有文件交叉确认',
        target_environment.upper(), len(successful_uploads),
    )
    # 每条营业执照只在其 OCR 到 UAT 上传之间短暂保留，避免整批附件占满内存。
    ocr_upload_memory: Dict[Tuple[str, int, int], bytes] = {}
    upload_action_progress = 0
    # 结果清单逐行落盘。即使下载或网络请求使进程中断，重跑时也能读取已经成功的记录。
    with open(result_file, 'w', encoding='utf-8-sig', newline='', buffering=1) as result_handle:
        writer = csv.DictWriter(result_handle, fieldnames=RESULT_COLUMNS)
        writer.writeheader()
        for source in source_attachments:
            prepared_upload_cache: Dict[
                Tuple[str, int, int], Tuple[AttachmentMeta, bytes, str]
            ] = {}
            outcome = match_results[source]
            all_metas = attachment_metadata.get(source.docid, [])
            approval_required = upload_planning_enabled and source.field_name in (
                BUSINESS_LICENSE_FIELD, IDENTITY_CARD_FIELD,
            )
            if approval_required:
                metas = [
                    meta for meta in all_metas
                    if _approved_document_key(source, meta) in approved_document_files
                ]
                for rejected_meta in (
                    meta for meta in all_metas
                    if _approved_document_key(source, meta) not in approved_document_files
                ):
                    row = _result_row(
                        source,
                        'skipped_review_not_approved',
                        '营业执照/身份证未通过最终复核，未上传',
                        meta=rejected_meta,
                    )
                    _write_result(writer, logger, row)
                    status_counts[row['status']] += 1
                if all_metas and not metas:
                    continue
            else:
                # 公司简介、开户许可、鲜章等无法按证件字段校验，直接进入匹配上传流程。
                metas = all_metas

            if args.execute and ocr_enabled and source.field_name == BUSINESS_LICENSE_FIELD:
                if execute_ocr_connection is None:
                    raise RuntimeError('生产 OCR SQLite 连接未初始化')
                license_metas = metas
                if not license_metas:
                    ocr_record = {
                        'ocr_called': False,
                        'ocr_status': 'missing_docimagefile',
                        'ocr_message': '泛微 docimagefile/imagefile 未找到营业执照附件元数据',
                    }
                    _write_ocr_audit_row(
                        execute_ocr_connection,
                        run_id,
                        source,
                        outcome,
                        has_business_license=True,
                        source_docid=source.docid,
                        ocr_record=ocr_record,
                    )
                    ocr_status_counts['missing_docimagefile'] += 1
                for ocr_meta in license_metas:
                    ocr_record, ocr_http_status, file_data = _get_business_license_ocr_result(
                        execute_ocr_connection,
                        source,
                        ocr_meta,
                        True,
                        ocr_url,
                        ocr_access_token,
                        cookie,
                        timeout_seconds,
                    )
                    _write_ocr_audit_row(
                        execute_ocr_connection,
                        run_id,
                        source,
                        outcome,
                        has_business_license=True,
                        meta=ocr_meta,
                        ocr_record=ocr_record,
                        http_status=ocr_http_status,
                    )
                    ocr_status_counts[_text(ocr_record.get('ocr_status'))] += 1
                    if file_data is not None:
                        ocr_upload_memory[(source.source_id, ocr_meta.docid, ocr_meta.imagefileid)] = file_data
                    logger.info(
                        '[OCR:%s%s] source=%s doc=%s file=%s',
                        ocr_record.get('ocr_status'),
                        ':cache' if ocr_record.get('ocr_cache_hit') else '',
                        source.source_id,
                        ocr_meta.docid,
                        ocr_meta.attachment_name,
                    )
            if not outcome.matches:
                message = (
                    '未找到目标环境客商，未上传'
                    if outcome.reason == 'identifier_not_found'
                    else '同一客商类型匹配到多条目标环境客商，未上传'
                )
                row = _result_row(source, outcome.reason, message)
                _write_result(writer, logger, row)
                status_counts[row['status']] += 1
                if source.field_name == BUSINESS_LICENSE_FIELD:
                    for meta in attachment_metadata.get(source.docid, []):
                        ocr_upload_memory.pop((source.source_id, meta.docid, meta.imagefileid), None)
                continue

            if outcome.reason == 'identifier_ambiguous_partial':
                ambiguous_labels = '、'.join(
                    '供应商' if target_type == 'vender' else '客户'
                    for target_type in outcome.ambiguous_target_types
                )
                row = _result_row(
                    source,
                    outcome.reason,
                    f'{ambiguous_labels}匹配到多条目标环境客商，已跳过该类型；其余唯一类型继续处理',
                )
                _write_result(writer, logger, row)
                status_counts[row['status']] += 1

            for match in outcome.matches:
                target = match.target
                source_identifier_type = match.source_identifier_type
                target_identifier_field = match.target_identifier_field
                if not target.unique_code:
                    if target.unique_code_source == 'duplicate_headerId':
                        missing_key_message = (
                            '生产客商 header_id 与其他启用客商重复；为避免附件串到错误客商，未上传。'
                        )
                    elif target_environment == 'prod':
                        missing_key_message = (
                            '生产客商主数据缺少 header_id；页面没有可安全绑定的附件业务键，未上传。'
                        )
                    else:
                        missing_key_message = (
                            'UAT 列表未返回 headerId；为避免附件挂错，未上传。'
                            '可核实后使用 --allow-primary-key-unique-code。'
                        )
                    row = _result_row(
                        source, 'target_attachment_key_missing',
                        missing_key_message,
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
                            upload_action_progress += 1
                            row = _result_row(
                                source, 'target_attachment_query_failed',
                                f'[进度 {upload_action_progress}/{planned_upload_actions}] {exc}', target,
                                source_identifier_type, target_identifier_field, meta, exc.status_code,
                            )
                            _write_result(writer, logger, row)
                            status_counts[row['status']] += 1
                        continue

                for meta in metas:
                    upload_action_progress += 1
                    progress_text = f'[进度 {upload_action_progress}/{planned_upload_actions}]'
                    existing_names = existing_names_cache[cache_key]
                    upload_meta = meta
                    upload_history_key = (
                        target.target_type, _text(target.code).casefold(), source.source_id,
                        meta.docid, meta.imagefileid,
                    )
                    prior_uploaded_names = successful_uploads.get(upload_history_key, set())
                    prior_uploaded_name = next(
                        (
                            uploaded_name
                            for uploaded_name in sorted(prior_uploaded_names)
                            if _normalized_file_name(uploaded_name) in existing_names
                        ),
                        '',
                    )
                    if prior_uploaded_name:
                        row = _result_row(
                            source,
                            'skipped_prior_success',
                            f'{progress_text} 历史清单已记录上传成功，且目标环境仍存在文件 '
                            f'{prior_uploaded_name}，本次跳过',
                            target, source_identifier_type, target_identifier_field,
                            replace(meta, attachment_name=prior_uploaded_name),
                        )
                        _write_result(writer, logger, row)
                        status_counts[row['status']] += 1
                        continue
                    if not args.execute:
                        original_name = _normalized_file_name(meta.attachment_name)
                        if not args.no_skip_existing and original_name in existing_names:
                            row = _result_row(
                                source,
                                'skipped_existing_name',
                                f'{progress_text} 目标营业执照页已有同名附件，本次跳过',
                                target, source_identifier_type, target_identifier_field, meta,
                            )
                            _write_result(writer, logger, row)
                            status_counts[row['status']] += 1
                            continue
                        row = _result_row(
                            source,
                            'matched_dry_run',
                            f'{progress_text} 已匹配，dry-run 未上传',
                            target, source_identifier_type, target_identifier_field, upload_meta,
                        )
                        _write_result(writer, logger, row)
                        status_counts[row['status']] += 1
                        continue

                    try:
                        memory_key = (source.source_id, meta.docid, meta.imagefileid)
                        prepared_upload = prepared_upload_cache.get(memory_key)
                        if prepared_upload is None:
                            if approval_required:
                                approved_path = approved_document_files[_approved_document_key(source, meta)]
                                file_data = _read_approved_document(approved_path)
                            else:
                                file_data = ocr_upload_memory.get(memory_key)
                            if file_data is None:
                                file_data = _download_weaver_attachment(
                                    cookie, meta, timeout_seconds,
                                )

                            prepared_meta = _with_detected_file_extension(meta, file_data)
                            if prepared_meta != meta:
                                logger.info(
                                    '[filename_extension_added] source=%s doc=%s old=%s new=%s',
                                    source.source_id, source.docid, meta.attachment_name,
                                    prepared_meta.attachment_name,
                                )

                            compression_note = ''
                            is_pdf = (
                                prepared_meta.attachment_name.casefold().endswith('.pdf')
                                or file_data[:1024].lstrip().startswith(b'%PDF-')
                            )
                            if len(file_data) > UAT_PDF_COMPRESSION_TRIGGER_BYTES and is_pdf:
                                original_size = len(file_data)
                                logger.info(
                                    '[pdf_compress_start] source=%s doc=%s file=%s original_size=%s',
                                    source.source_id, source.docid,
                                    prepared_meta.attachment_name, original_size,
                                )
                                file_data, compression_method = _compress_pdf_for_uat(file_data)
                                prepared_meta = replace(
                                    prepared_meta,
                                    attachment_name=_compressed_pdf_name(prepared_meta.attachment_name),
                                )
                                compression_note = (
                                    f'PDF压缩 {original_size} -> {len(file_data)} bytes'
                                    f'（{compression_method}）'
                                )
                                logger.info(
                                    '[pdf_compress_done] source=%s doc=%s file=%s '
                                    'original_size=%s compressed_size=%s method=%s',
                                    source.source_id, source.docid,
                                    prepared_meta.attachment_name, original_size,
                                    len(file_data), compression_method,
                                )
                            prepared_upload = (prepared_meta, file_data, compression_note)
                            prepared_upload_cache[memory_key] = prepared_upload

                        upload_meta, file_data, compression_note = prepared_upload
                        prepared_name = _normalized_file_name(upload_meta.attachment_name)
                        if not args.no_skip_existing and prepared_name in existing_names:
                            row = _result_row(
                                source,
                                'skipped_existing_name',
                                f'{progress_text} 目标营业执照页已有同名附件，本次跳过',
                                target, source_identifier_type, target_identifier_field,
                                upload_meta,
                            )
                            _write_result(writer, logger, row)
                            logger.info(
                                '[upload_done] progress=%s/%s status=%s source=%s doc=%s '
                                'target=%s/%s file=%s',
                                upload_action_progress, planned_upload_actions, row['status'],
                                source.source_id, source.docid, target.target_type, target.code,
                                upload_meta.attachment_name,
                            )
                            status_counts[row['status']] += 1
                            continue
                        logger.info(
                            '[upload_start] progress=%s/%s source=%s doc=%s target=%s/%s '
                            'file=%s size=%s',
                            upload_action_progress, planned_upload_actions,
                            source.source_id, source.docid, target.target_type, target.code,
                            upload_meta.attachment_name, len(file_data),
                        )
                        status_code = _retry_upload(
                            base_url, organization_id, access_token, target, upload_meta, file_data,
                            timeout_seconds, upload_retries,
                        )
                        existing_names.add(_normalized_file_name(upload_meta.attachment_name))
                        successful_uploads.setdefault(upload_history_key, set()).add(
                            upload_meta.attachment_name
                        )
                        row = _result_row(
                            source,
                            'uploaded',
                            f'{progress_text} '
                            + (f'{compression_note}；' if compression_note else '')
                            + '上传成功，'
                            + f'{len(file_data)} bytes',
                            target, source_identifier_type, target_identifier_field, upload_meta, status_code,
                        )
                    except UatRequestError as exc:
                        row = _result_row(
                            source, 'upload_failed', f'{progress_text} {exc}', target, source_identifier_type,
                            target_identifier_field, upload_meta, exc.status_code,
                        )
                    except Exception as exc:  # 单个附件失败不可中断全量迁移。
                        row = _result_row(
                            source, 'download_or_upload_failed', f'{progress_text} {exc}', target,
                            source_identifier_type, target_identifier_field, upload_meta,
                        )
                    _write_result(writer, logger, row)
                    logger.info(
                        '[upload_done] progress=%s/%s status=%s source=%s doc=%s target=%s/%s file=%s',
                        upload_action_progress, planned_upload_actions, row['status'],
                        source.source_id, source.docid, target.target_type, target.code,
                        upload_meta.attachment_name,
                    )
                    status_counts[row['status']] += 1

            if source.field_name == BUSINESS_LICENSE_FIELD:
                for meta in attachment_metadata.get(source.docid, []):
                    ocr_upload_memory.pop((source.source_id, meta.docid, meta.imagefileid), None)

    if execute_ocr_connection is not None:
        execute_ocr_connection.close()
        logger.info('营业执照 OCR 审计完成: %s', dict(sorted(ocr_status_counts.items())))
        logger.info('营业执照 OCR SQLite: %s', ocr_audit_db)
    logger.info('同步完成: %s', dict(sorted(status_counts.items())))
    logger.info('结果清单: %s', result_file)
    logger.info('运行日志: %s', log_file)
    return result_file


if __name__ == '__main__':
    run()
