# -*- coding: utf-8 -*-
"""SMTP 邮件发送工具，用于向任务提交人发送超大结果 ZIP。"""
from __future__ import annotations

import os
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path


def _enabled(value: str) -> bool:
    return value.strip().lower() in {'1', 'true', 'yes', 'y', 'on'}


def _smtp_config() -> tuple[str, int, str, str, str, bool, bool]:
    host = os.getenv('SMTP_HOST', '').strip()
    sender = os.getenv('SMTP_FROM', '').strip()
    username = os.getenv('SMTP_USERNAME', '').strip()
    password = os.getenv('SMTP_PASSWORD', '')
    if not host or not sender:
        raise RuntimeError('缺少 SMTP_HOST 或 SMTP_FROM，无法发送结果邮件')
    if username and not password:
        raise RuntimeError('已配置 SMTP_USERNAME，但缺少 SMTP_PASSWORD')
    try:
        port = int(os.getenv('SMTP_PORT', '465').strip())
    except ValueError as exc:
        raise RuntimeError('SMTP_PORT 必须为数字') from exc
    use_ssl = _enabled(os.getenv('SMTP_USE_SSL', '1'))
    starttls = _enabled(os.getenv('SMTP_STARTTLS', '0'))
    if use_ssl and starttls:
        raise RuntimeError('SMTP_USE_SSL 与 SMTP_STARTTLS 不能同时启用')
    return host, port, sender, username, password, use_ssl, starttls


def send_result_zip_email(recipient: str, package: Path, download_url: str, run_id: str) -> None:
    """将结果 ZIP 作为附件发送；SMTP 服务自身的大小限制由其错误反馈。"""
    destination = str(recipient or '').strip()
    path = Path(package)
    if not destination:
        raise ValueError('收件邮箱不能为空')
    if not path.is_file() or path.stat().st_size <= 0:
        raise FileNotFoundError(f'邮件附件不存在或为空: {path}')

    host, port, sender, username, password, use_ssl, starttls = _smtp_config()
    message = EmailMessage()
    message['Subject'] = f'智书合同导入清单结果 - {run_id[:8]}'
    message['From'] = sender
    message['To'] = destination
    message.set_content(
        '智书合同导入清单任务已完成。\n\n'
        f'任务编号：{run_id}\n'
        '结果 ZIP 已作为邮件附件发送。\n'
        f'备用下载地址：{download_url}\n'
    )
    message.add_attachment(
        path.read_bytes(),
        maintype='application',
        subtype='zip',
        filename=path.name,
    )

    if use_ssl:
        with smtplib.SMTP_SSL(host, port, context=ssl.create_default_context(), timeout=60) as client:
            if username:
                client.login(username, password)
            client.send_message(message)
        return

    with smtplib.SMTP(host, port, timeout=60) as client:
        if starttls:
            client.starttls(context=ssl.create_default_context())
        if username:
            client.login(username, password)
        client.send_message(message)
