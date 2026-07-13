# -*- coding: utf-8 -*-
"""ETL 的 HTTP API。

所有 ETL 任务会写入共享的输出目录，也会复用模块级配置。因此服务一次只执行
一个任务。混合合同增补支持直接上传 Excel；任务完成后可下载本次生成的 Excel、
JSON、审计文件和附件 ZIP。
"""
from __future__ import annotations

import io
import os
import re
import sys
import threading
import traceback
import uuid
import zipfile
from contextlib import contextmanager, redirect_stderr, redirect_stdout
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Literal

from fastapi import FastAPI, File, Form, HTTPException, Query, Request, UploadFile, status
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from etl.util import common as c
from run import TASKS


API_RUN_ROOT = c.OUT_DIR / "api_runs"
MAX_UPLOAD_BYTES = int(os.getenv("API_MAX_UPLOAD_MB", "50")) * 1024 * 1024
UPLOAD_CHUNK_BYTES = 1024 * 1024
EXCEL_EXTENSIONS = frozenset({".xlsx", ".xlsm"})
FEISHU_OPEN_ID_PATTERN = re.compile(r"^ou_[a-z0-9]{16,64}$")
EMAIL_PATTERN = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
FEISHU_FILE_MAX_BYTES = 30 * 1024 * 1024
RunStatus = Literal["queued", "running", "succeeded", "failed"]
NotificationStatus = Literal["not_requested", "pending", "sent", "failed"]


class RunInfo(BaseModel):
    """一次 ETL 运行的状态。"""

    run_id: str
    task_name: str
    status: RunStatus
    submitted_at: datetime
    started_at: datetime | None = None
    finished_at: datetime | None = None
    error: str | None = None
    log_tail: str = ""
    output_files: list[str] = Field(default_factory=list)
    download_url: str | None = None
    notification_status: NotificationStatus = "not_requested"
    notification_error: str | None = None


class RunLog(BaseModel):
    run_id: str
    log_tail: str
    is_truncated: bool


class HealthInfo(BaseModel):
    status: Literal["ok"] = "ok"
    active_run_id: str | None = None


class TaskInfo(BaseModel):
    name: str
    description: str = ""


class _LiveRunLog(io.TextIOBase):
    """将任务输出同时写入终端与内存，供日志接口在任务进行中读取。"""

    def __init__(self, run_id: str, terminal: io.TextIOBase) -> None:
        self.run_id = run_id
        self.terminal = terminal

    def write(self, value: str) -> int:
        text = str(value)
        if not text:
            return 0
        with _state_lock:
            content = _run_logs.get(self.run_id, "") + text
            _run_logs[self.run_id] = content
            run = _runs.get(self.run_id)
            if run is not None:
                run.log_tail = _tail(content, 4_000)[0]
        self.terminal.write(text)
        self.terminal.flush()
        return len(text)

    def flush(self) -> None:
        self.terminal.flush()

    def getvalue(self) -> str:
        with _state_lock:
            return _run_logs.get(self.run_id, "")


app = FastAPI(
    title="Hero Digital ETL API",
    version="1.0.0",
    description=(
        "提交 ETL 任务、查看运行日志，并下载混合合同增补结果。\n\n"
        "**混合合同增补使用顺序：**先下载业务清单模板，填写后提交；"
        "任务结束会向指定飞书用户发送通知。"
    ),
)

_state_lock = threading.RLock()
_runs: dict[str, RunInfo] = {}
_run_logs: dict[str, str] = {}
_active_run_id: str | None = None


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _tail(value: str, chars: int) -> tuple[str, bool]:
    if len(value) <= chars:
        return value, False
    return value[-chars:], True


def _copy_run(run: RunInfo) -> RunInfo:
    return run.model_copy(deep=True)


def _get_run_or_404(run_id: str) -> RunInfo:
    with _state_lock:
        run = _runs.get(run_id)
        if run is None:
            raise HTTPException(status_code=404, detail=f"运行记录不存在: {run_id}")
        return run


def _task_description(task: Callable[[], Any]) -> str:
    return (task.__doc__ or "").strip().split("\n", 1)[0]


def health() -> HealthInfo:
    """返回服务状态，供健康检查和自动化调用。"""
    with _state_lock:
        return HealthInfo(active_run_id=_active_run_id)


@app.get("/health", response_model=HealthInfo, tags=["系统"])
def health_endpoint() -> HealthInfo:
    return health()


@app.get("/tasks", response_model=list[TaskInfo], tags=["任务"])
def list_tasks() -> list[TaskInfo]:
    return [
        TaskInfo(name=name, description=_task_description(task))
        for name, task in sorted(TASKS.items())
    ]


def _save_log_file(run_id: str, content: str) -> None:
    log_path = API_RUN_ROOT / run_id / "run.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text(content, encoding="utf-8")


def _download_url(run_id: str, request_base_url: str | None = None) -> str:
    path = f"/runs/{run_id}/download"
    base_url = (
        os.getenv("API_PUBLIC_BASE_URL", "").strip().rstrip("/")
        or str(request_base_url or "").strip().rstrip("/")
    )
    return f"{base_url}{path}" if base_url else path


def _result_zip_path(run: RunInfo) -> Path | None:
    """解析运行记录中实际存在的结果 ZIP 路径。"""
    for item in run.output_files:
        path = Path(item)
        candidate = path if path.is_absolute() else c.ROOT / path
        if candidate.suffix.lower() == ".zip" and candidate.is_file():
            return candidate
    return None


def _run_download_url(run: RunInfo) -> str:
    # 已配置对外/局域网地址时，不能回退到 Swagger 请求中可能出现的 127.0.0.1。
    if os.getenv("API_PUBLIC_BASE_URL", "").strip():
        return _download_url(run.run_id)
    return run.download_url or _download_url(run.run_id)


def _notification_card(
    run: RunInfo,
    *,
    title: str,
    template: str,
    delivery: str,
    package: Path | None = None,
    error: str | None = None,
    show_download: bool = True,
) -> dict[str, Any]:
    """构造兼容飞书消息接口的 1.0 交互卡片。"""
    fields = [{
        "is_short": False,
        "text": {"tag": "lark_md", "content": f"**任务编号**\n`{run.run_id}`"},
    }, {
        "is_short": False,
        "text": {"tag": "lark_md", "content": f"**处理结果**\n{delivery}"},
    }]
    if package is not None:
        fields.append({
            "is_short": True,
            "text": {
                "tag": "lark_md",
                "content": f"**结果包大小**\n{package.stat().st_size / 1024 / 1024:.1f} MB",
            },
        })
    if error:
        fields.append({
            "is_short": False,
            "text": {
                "tag": "lark_md",
                "content": f"**说明**\n{str(error).strip()[:500]}",
            },
        })

    elements: list[dict[str, Any]] = [{"tag": "div", "fields": fields}]
    if show_download:
        elements.extend((
            {"tag": "hr"},
            {
                "tag": "action",
                "actions": [{
                    "tag": "button",
                    "text": {"tag": "plain_text", "content": "下载结果"},
                    "type": "primary",
                    "url": _run_download_url(run),
                }],
            },
        ))
    return {
        "config": {"wide_screen_mode": True, "enable_forward": True},
        "header": {
            "template": template,
            "title": {"tag": "plain_text", "content": title},
        },
        "elements": elements,
    }


def _send_notification_card(
    feishu: Any,
    notify_open_id: str,
    run: RunInfo,
    **card_kwargs: Any,
) -> None:
    """优先发交互卡片；飞书卡片配置异常时降级为文本通知。"""
    card = _notification_card(run, **card_kwargs)
    try:
        feishu.send_interactive_card(notify_open_id, card)
    except Exception:
        fallback = (
            f"{card_kwargs['title']}\n"
            f"任务编号：{run.run_id}\n"
            f"处理结果：{card_kwargs['delivery']}\n"
            f"下载结果：{_run_download_url(run)}"
        )
        feishu.send_text_message(notify_open_id, fallback)


def _send_notification(
    run: RunInfo,
    notify_open_id: str,
    notify_email: str | None = None,
) -> None:
    from etl.lark import feishu

    if run.status != "succeeded":
        _send_notification_card(
            feishu,
            notify_open_id,
            run,
            title="混合合同增补执行失败",
            template="red",
            delivery="任务执行失败，请查看错误说明后修正业务清单。",
            error=run.error or "请通过接口日志查看详情",
            show_download=False,
        )
        return

    package = _result_zip_path(run)
    if package is None:
        raise FileNotFoundError(f"任务结果 ZIP 不存在: {run.run_id}")
    if package.stat().st_size > FEISHU_FILE_MAX_BYTES:
        if notify_email:
            from etl.util.mailer import send_result_zip_email

            try:
                send_result_zip_email(
                    notify_email,
                    package,
                    _run_download_url(run),
                    run.run_id,
                )
            except Exception as exc:
                _send_notification_card(
                    feishu,
                    notify_open_id,
                    run,
                    title="混合合同增补已完成",
                    template="orange",
                    delivery="ZIP 超过飞书 30 MB 限制，邮件发送失败，请通过接口下载。",
                    package=package,
                    error=f"{type(exc).__name__}: {exc}",
                )
                return
            _send_notification_card(
                feishu,
                notify_open_id,
                run,
                title="混合合同增补已完成",
                template="green",
                delivery="ZIP 超过飞书 30 MB 限制，已发送至您填写的邮箱。",
                package=package,
            )
            return
        _send_notification_card(
            feishu,
            notify_open_id,
            run,
            title="混合合同增补已完成",
            template="orange",
            delivery="ZIP 超过飞书 30 MB 限制；未填写邮箱，请通过接口下载。",
            package=package,
        )
        return
    try:
        feishu.send_file_message(notify_open_id, package)
    except Exception as exc:
        _send_notification_card(
            feishu,
            notify_open_id,
            run,
            title="混合合同增补已完成",
            template="orange",
            delivery="ZIP 飞书附件发送失败，请通过接口下载。",
            package=package,
            error=f"{type(exc).__name__}: {exc}",
        )
        return
    _send_notification_card(
        feishu,
        notify_open_id,
        run,
        title="混合合同增补已完成",
        template="green",
        delivery="结果 ZIP 已作为飞书附件发送。",
        package=package,
    )


def _run_in_background(
    run_id: str,
    runner: Callable[[], Any],
    notify_open_id: str | None = None,
    notify_email: str | None = None,
) -> None:
    global _active_run_id

    stdout_log = _LiveRunLog(run_id, sys.__stdout__)
    stderr_log = _LiveRunLog(run_id, sys.__stderr__)
    result: Any = None
    error: str | None = None
    succeeded = False
    with _state_lock:
        run = _runs[run_id]
        run.status = "running"
        run.started_at = _now()

    try:
        # 运行期间禁止并发提交，因而重定向标准输出不会串入另一个 ETL 任务日志。
        with redirect_stdout(stdout_log), redirect_stderr(stderr_log):
            result = runner()
        succeeded = True
    except Exception as exc:  # 保留任务原始异常及回溯，供调用方排查。
        error = f"{type(exc).__name__}: {exc}"
        traceback.print_exc(file=stderr_log)

    content = stdout_log.getvalue()
    completion_status: RunStatus = "succeeded" if succeeded else "failed"
    notification_run: RunInfo | None = None
    with _state_lock:
        run = _runs[run_id]
        run.finished_at = _now()
        run.error = error
        run.log_tail = _tail(content, 4_000)[0]
        if succeeded:
            run.output_files = _normalise_output_files(result)
            if run.task_name == "contract_mixed_add_all" and any(
                item.endswith(".zip") for item in run.output_files
            ):
                run.download_url = run.download_url or _download_url(run_id)
        if notify_open_id:
            run.notification_status = "pending"
            notification_run = _copy_run(run)
            notification_run.status = completion_status

    notification_error: str | None = None
    notification_sent = False
    if notify_open_id:
        try:
            _send_notification(
                notification_run or _copy_run(run),
                notify_open_id,
                notify_email,
            )
            notification_sent = True
            print("[API] 已发送飞书完成通知", file=stdout_log)
        except Exception as exc:
            notification_error = f"{type(exc).__name__}: {exc}"
            print(f"[API] 飞书通知发送失败: {notification_error}", file=stderr_log)

    content = stdout_log.getvalue()
    _save_log_file(run_id, content)
    with _state_lock:
        run = _runs[run_id]
        run.status = completion_status
        run.log_tail = _tail(content, 4_000)[0]
        if notify_open_id:
            run.notification_status = "sent" if notification_sent else "failed"
            run.notification_error = notification_error
        _active_run_id = None


def _normalise_output_files(value: Any) -> list[str]:
    """将任务返回的路径统一为用于状态展示的相对路径。"""
    if value is None:
        return []
    values = value if isinstance(value, (list, tuple, set)) else [value]
    files: list[str] = []
    for item in values:
        if not item:
            continue
        path = Path(item)
        try:
            files.append(str(path.resolve().relative_to(c.ROOT.resolve())))
        except ValueError:
            files.append(str(path))
    return files


def create_run(task_name: str, runner: Callable[[], Any] | None = None) -> RunInfo:
    """提交一个后台任务。测试或内部调用可传入 ``runner`` 覆盖已登记任务。"""
    global _active_run_id

    if runner is None:
        runner = TASKS.get(task_name)
    if runner is None:
        raise HTTPException(status_code=404, detail=f"未知任务: {task_name}")

    with _state_lock:
        if _active_run_id is not None:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f"已有任务正在运行: {_active_run_id}",
            )
        run_id = uuid.uuid4().hex
        run = RunInfo(
            run_id=run_id,
            task_name=task_name,
            status="queued",
            submitted_at=_now(),
        )
        _runs[run_id] = run
        _run_logs[run_id] = ""
        _active_run_id = run_id

    thread = threading.Thread(
        target=_run_in_background,
        args=(run_id, runner),
        name=f"etl-{run_id[:8]}",
        daemon=True,
    )
    thread.start()
    return _copy_run(run)


@app.post(
    "/tasks/{task_name}/runs",
    response_model=RunInfo,
    status_code=status.HTTP_202_ACCEPTED,
    tags=["任务"],
)
def create_task_run(task_name: str) -> RunInfo:
    return create_run(task_name)


@app.get("/runs", response_model=list[RunInfo], tags=["任务"])
def list_runs(limit: int = Query(default=20, ge=1, le=100)) -> list[RunInfo]:
    with _state_lock:
        runs = sorted(_runs.values(), key=lambda item: item.submitted_at, reverse=True)
        return [_copy_run(run) for run in runs[:limit]]


def get_run(run_id: str) -> RunInfo:
    return _copy_run(_get_run_or_404(run_id))


@app.get("/runs/{run_id}", response_model=RunInfo, tags=["任务"])
def get_run_endpoint(run_id: str) -> RunInfo:
    return get_run(run_id)


def get_run_logs(run_id: str, tail_chars: int = 4_000) -> RunLog:
    if not 1 <= tail_chars <= 100_000:
        raise HTTPException(status_code=422, detail="tail_chars 必须在 1 到 100000 之间")
    _get_run_or_404(run_id)
    with _state_lock:
        content = _run_logs.get(run_id, "")
    log_tail, is_truncated = _tail(content, tail_chars)
    return RunLog(run_id=run_id, log_tail=log_tail, is_truncated=is_truncated)


@app.get("/runs/{run_id}/logs", response_model=RunLog, tags=["任务"])
def get_run_logs_endpoint(
    run_id: str,
    tail_chars: int = Query(default=4_000, ge=1, le=100_000),
) -> RunLog:
    return get_run_logs(run_id, tail_chars)


def _safe_upload_name(filename: str | None) -> str:
    name = Path(filename or "contract_mixed_add.xlsx").name
    if not name or name in {".", ".."}:
        name = "contract_mixed_add.xlsx"
    suffix = Path(name).suffix.lower()
    if suffix not in EXCEL_EXTENSIONS:
        raise HTTPException(
            status_code=415,
            detail="只支持 .xlsx 或 .xlsm 格式的 Excel 文件",
        )
    return name


async def _store_upload(run_id: str, upload: UploadFile) -> Path:
    name = _safe_upload_name(upload.filename)
    path = API_RUN_ROOT / run_id / "input" / name
    path.parent.mkdir(parents=True, exist_ok=True)
    size = 0
    try:
        with path.open("wb") as target:
            while chunk := await upload.read(UPLOAD_CHUNK_BYTES):
                size += len(chunk)
                if size > MAX_UPLOAD_BYTES:
                    raise HTTPException(
                        status_code=413,
                        detail=(
                            "上传文件超过限制: "
                            f"{MAX_UPLOAD_BYTES // 1024 // 1024} MB"
                        ),
                    )
                target.write(chunk)
    except Exception:
        path.unlink(missing_ok=True)
        raise
    finally:
        await upload.close()
    if not size:
        path.unlink(missing_ok=True)
        raise HTTPException(status_code=422, detail="上传的 Excel 为空")
    return path


def _add_file_to_zip(archive: zipfile.ZipFile, file_path: Path, arcname: Path) -> None:
    if file_path.is_file():
        archive.write(file_path, arcname.as_posix())


def _package_contract_mixed_add_result(run_id: str, generated: list[Path]) -> Path:
    """打包该次任务的 Excel、请求 JSON 和实际下载附件。"""
    from etl.contract import contract_mixed_add as mixed

    package = API_RUN_ROOT / run_id / "result" / "contract_mixed_add_result.zip"
    package.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(package, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in generated:
            if path.is_file():
                _add_file_to_zip(archive, path, Path("输出文件") / path.name)

        attachment_root = mixed.MIXED_ATTACHMENT_ROOT
        if attachment_root.exists():
            for attachment in attachment_root.rglob("*"):
                if attachment.is_file():
                    _add_file_to_zip(
                        archive,
                        attachment,
                        Path("附件") / attachment.relative_to(attachment_root),
                    )
    return package


@contextmanager
def _mixed_add_run_context(input_file: Path, output_dir: Path):
    """将历史脚本的模块级输入/输出路径临时切换到单次 API 运行目录。"""
    from etl.contract import contract_mixed_add as mixed
    from etl.contract import contract_mixed_add_attachments_db as attachments

    output_dir.mkdir(parents=True, exist_ok=True)
    mixed_attributes = (
        "INPUT_FILE",
        "OUTPUT_DIR",
        "GENERAL_OUTPUT_FILE",
        "ANCHOR_OUTPUT_FILE",
        "MIXED_OUTPUT_FILE",
        "MIXED_ATTACHMENT_ROOT",
        "ARCHIVED_REQUEST_FILE",
        "OTHER_REQUEST_FILE",
        "APPROVE_TO_NODE_REQUEST_FILE",
        "YECAI_SYNC_REQUEST_FILE",
        "AUDIT_FILE",
    )
    previous_mixed = {name: getattr(mixed, name) for name in mixed_attributes}
    previous_attachment_output_dir = attachments.OUTPUT_DIR
    previous_manifest_file = attachments.MANIFEST_FILE

    mixed.INPUT_FILE = input_file
    mixed.OUTPUT_DIR = output_dir
    mixed.GENERAL_OUTPUT_FILE = output_dir / previous_mixed["GENERAL_OUTPUT_FILE"].name
    mixed.ANCHOR_OUTPUT_FILE = output_dir / previous_mixed["ANCHOR_OUTPUT_FILE"].name
    mixed.MIXED_OUTPUT_FILE = output_dir / previous_mixed["MIXED_OUTPUT_FILE"].name
    mixed.MIXED_ATTACHMENT_ROOT = output_dir / previous_mixed["MIXED_ATTACHMENT_ROOT"].name
    mixed.ARCHIVED_REQUEST_FILE = output_dir / previous_mixed["ARCHIVED_REQUEST_FILE"].name
    mixed.OTHER_REQUEST_FILE = output_dir / previous_mixed["OTHER_REQUEST_FILE"].name
    mixed.APPROVE_TO_NODE_REQUEST_FILE = output_dir / previous_mixed["APPROVE_TO_NODE_REQUEST_FILE"].name
    mixed.YECAI_SYNC_REQUEST_FILE = output_dir / previous_mixed["YECAI_SYNC_REQUEST_FILE"].name
    mixed.AUDIT_FILE = output_dir / previous_mixed["AUDIT_FILE"].name
    attachments.OUTPUT_DIR = output_dir / "附件下载清单"
    attachments.MANIFEST_FILE = attachments.OUTPUT_DIR / previous_manifest_file.name

    try:
        yield mixed, attachments
    finally:
        for name, value in previous_mixed.items():
            setattr(mixed, name, value)
        attachments.OUTPUT_DIR = previous_attachment_output_dir
        attachments.MANIFEST_FILE = previous_manifest_file


def _run_contract_mixed_add_all(input_file: Path, run_id: str) -> list[Path]:
    """在当前任务线程内用上传文件运行原有的一键流程。"""
    generated_dir = API_RUN_ROOT / run_id / "generated"
    with _mixed_add_run_context(input_file, generated_dir) as (mixed, attachments):
        print("=== 运行 contract_mixed_add_all: contract_mixed_add ===")
        generated = [Path(path) for path in mixed.run(suppress_audit=True)]
        print("=== 运行 contract_mixed_add_all: contract_mixed_add_attachments_db ===")
        attachment_manifest = attachments.run(suppress_manifest=True)
        if attachment_manifest:
            generated.append(Path(attachment_manifest))
        package = _package_contract_mixed_add_result(run_id, generated)
        print(f"[API] 已打包本次结果: {package}")
        return [*generated, package]


@app.get(
    "/contract-mixed-add/template",
    response_class=FileResponse,
    include_in_schema=False,
)
def download_contract_mixed_add_template() -> FileResponse:
    """下载混合流程接口上传所需的业务清单模板。"""
    from etl.contract import contract_mixed_add as mixed

    template = mixed.INPUT_TEMPLATE_FILE
    if not template.is_file():
        raise HTTPException(status_code=503, detail="混合合同增补模板尚未部署，请联系管理员")
    return FileResponse(
        template,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="混合合同增补业务清单模板.xlsx",
    )


@app.post(
    "/contract-mixed-add/runs",
    response_model=RunInfo,
    status_code=status.HTTP_202_ACCEPTED,
    tags=["混合合同增补"],
    summary="上传 Excel 并运行 contract_mixed_add_all",
    description="""
上传已填写的混合合同增补业务清单，系统会按合同类型分流到一般流程或主播流程，生成可导入智书的 13 Sheet Excel、同步 JSON 和附件结果 ZIP。为减少结果包体积，接口调用不会生成“处理清单”和“附件下载清单”两个审计 Excel。

**操作步骤：**

1. 点击 [下载业务清单模板](/contract-mixed-add/template)；填写前三列后上传。结果 ZIP 中会生成 13 Sheet 智书导入 Excel，不需要将其作为输入上传。
2. 填写业务清单的前三列：合同编号、关联业财订单、智书合同类型；上传 `.xlsx` 或 `.xlsm` 文件。
3. 填入接收飞书通知的 `notify_open_id`；`notify_email` 可选，仅在 ZIP 超过 30 MB 时用于接收邮件附件。响应中的 `run_id` 可用于查询状态、日志和下载 ZIP。
4. 机器人会以飞书卡片通知任务结果，并提供“下载结果”按钮。成功后会发送结果 ZIP 附件（飞书单文件上限 30 MB）；若超限，填写了邮箱则发送邮件附件，否则请点击卡片按钮或调用 `GET /runs/{run_id}/download` 下载结果。失败时请通过 `GET /runs/{run_id}/logs` 查看完整日志。

**如何获取 Open ID：**在飞书中 `@系统咨询小助手`，询问“我的 Open ID 是多少”。只可使用该通知机器人应用对应的 `ou_...`；其他应用或其他机器人的 Open ID 会因应用隔离而无法收取消息。
""",
)
async def create_contract_mixed_add_run(
    request: Request,
    file: UploadFile = File(
        ...,
        description="使用上方模板填写后的业务清单，仅接受 .xlsx 或 .xlsm",
    ),
    notify_open_id: str = Form(
        ...,
        description=(
            "必填。任务结束后接收飞书通知的 Open ID，格式：ou_ + 16 至 64 位小写字母/数字。"
            "请在飞书中 @系统咨询小助手 查询当前通知机器人应用对应的 Open ID。"
        ),
        min_length=19,
        max_length=67,
        pattern=FEISHU_OPEN_ID_PATTERN.pattern,
    ),
    notify_email: str | None = Form(
        default=None,
        description=(
            "可选。仅当结果 ZIP 超过飞书 30 MB 限制时，用于接收邮件附件；"
            "留空则通过飞书消息中的接口地址自行下载。"
        ),
        max_length=254,
    ),
) -> RunInfo:
    """上传业务清单；处理完成后从 ``/runs/{run_id}/download`` 下载 ZIP。"""
    global _active_run_id

    if not FEISHU_OPEN_ID_PATTERN.fullmatch(notify_open_id):
        raise HTTPException(
            status_code=422,
            detail="notify_open_id 格式错误：必须为 ou_ + 16 至 64 位小写字母或数字",
        )
    notify_email = (notify_email or "").strip().lower() or None
    if notify_email and not EMAIL_PATTERN.fullmatch(notify_email):
        raise HTTPException(status_code=422, detail="notify_email 格式错误：请输入有效邮箱地址")

    # 先取得全局运行资格，避免无效上传与正在运行任务争抢同一输出目录。
    with _state_lock:
        if _active_run_id is not None:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f"已有任务正在运行: {_active_run_id}",
            )
        run_id = uuid.uuid4().hex
        placeholder = RunInfo(
            run_id=run_id,
            task_name="contract_mixed_add_all",
            status="queued",
            submitted_at=_now(),
            download_url=_download_url(run_id, str(request.base_url)),
            notification_status="pending",
        )
        _runs[run_id] = placeholder
        _run_logs[run_id] = ""
        # 上传时也占用任务槽，防止保存结束前其他请求插入运行。
        _active_run_id = run_id

    try:
        input_file = await _store_upload(run_id, file)
    except Exception:
        with _state_lock:
            _runs.pop(run_id, None)
            _active_run_id = None
        raise

    thread = threading.Thread(
        target=_run_in_background,
        args=(
            run_id,
            lambda: _run_contract_mixed_add_all(input_file, run_id),
            notify_open_id,
            notify_email,
        ),
        name=f"etl-{run_id[:8]}",
        daemon=True,
    )
    thread.start()
    return _copy_run(placeholder)


@app.get(
    "/runs/{run_id}/download",
    response_class=FileResponse,
    responses={
        200: {
            "description": "混合合同增补结果 ZIP",
            "content": {"application/zip": {"schema": {"type": "string", "format": "binary"}}},
        },
    },
    tags=["混合合同增补"],
)
def download_run_result(run_id: str) -> FileResponse:
    run = _get_run_or_404(run_id)
    if run.status in {"queued", "running"}:
        raise HTTPException(status_code=409, detail="任务尚未完成")
    if run.status == "failed":
        raise HTTPException(status_code=409, detail="任务执行失败，无法下载结果")

    package = _result_zip_path(run)
    if package is None:
        raise HTTPException(status_code=404, detail="该任务没有可下载的结果包")
    return FileResponse(
        package,
        media_type="application/zip",
        filename=f"contract_mixed_add_{run_id[:8]}.zip",
    )
