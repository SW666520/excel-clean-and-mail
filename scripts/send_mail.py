#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""通过 SMTP 发送清洗结果邮件（唯一外部服务：网易邮箱）。

本脚本只依赖 Python 标准库（smtplib / email），不需要额外安装第三方包。
SMTP 账号凭据从配置文件读取，禁止把邮箱密码/授权码写死在本脚本里。

用法:
    python send_mail.py --to 收件人@example.com --subject "清洗结果" \
        --body "清洗完成，见附件。" --attach 清洗后.xlsx [--attach 汇总.xlsx]
    python send_mail.py --config config/smtp_config.json --to ... --dry-run

--dry-run 只校验配置与参数，不真正连接服务器，用于上线前自检。
"""
from __future__ import annotations

import argparse
import json
import os
import smtplib
import sys
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

CONFIG_FIELDS = ("smtp_host", "smtp_port", "use_ssl",
                 "sender_email", "auth_code")


def load_config(path: str) -> dict:
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(
            f"配置文件不存在: {path}\n"
            f"请把 config/smtp_config.example.json 复制为 config/smtp_config.json "
            f"并按《配置指引.md》填写后重试。")
    with open(p, encoding="utf-8") as f:
        cfg = json.load(f)
    missing = [k for k in CONFIG_FIELDS
               if not str(cfg.get(k, "")).strip()]
    if missing:
        raise ValueError(
            f"配置文件 {path} 缺少必填字段: {missing}\n"
            f"请参照《配置指引.md》中的「需要手动填写的位置」补全。")
    cfg["smtp_port"] = int(cfg["smtp_port"])
    return cfg


def build_message(subject: str, body: str,
                  sender_email: str, to_addrs: list[str],
                  attach_paths: list[str]) -> MIMEMultipart:
    msg = MIMEMultipart()
    msg["From"] = f"{Header('表格清洗官', 'utf-8')} <{sender_email}>"
    msg["To"] = ", ".join(to_addrs)
    msg["Subject"] = Header(subject, "utf-8")
    msg["Date"] = formatdate(localtime=True)

    msg.attach(MIMEText(body, "plain", "utf-8"))

    for ap in attach_paths:
        p = Path(ap)
        if not p.is_file():
            raise FileNotFoundError(f"附件不存在: {ap}")
        # 附件统一按二进制读入，避免中文名与编码问题
        with open(p, "rb") as f:
            payload = f.read()
        import email.mime.base
        import email.encoders
        mime_part = email.mime.base.MIMEBase(
            "application", "octet-stream")
        mime_part.set_payload(payload)
        email.encoders.encode_base64(mime_part)
        mime_part.add_header(
            "Content-Disposition", "attachment",
            filename=("utf-8", "", str(p.name)))
        msg.attach(mime_part)
    return msg


def send(cfg: dict, to_addrs: list[str], subject: str, body: str,
         attach_paths: list[str], dry_run: bool = False) -> dict:
    msg = build_message(subject, body, cfg["sender_email"],
                        to_addrs, attach_paths)

    summary = {
        "sender": cfg["sender_email"],
        "to": to_addrs,
        "subject": subject,
        "attachments": [os.path.basename(a) for a in attach_paths],
        "server": f"{cfg['smtp_host']}:{cfg['smtp_port']}"
                  f"({'SSL' if cfg['use_ssl'] else 'STARTTLS'})",
    }
    if dry_run:
        summary["dry_run"] = True
        summary["ok"] = True
        summary["message"] = "配置与参数校验通过（未实际发送）"
        return summary

    if cfg["use_ssl"]:
        client = smtplib.SMTP_SSL(cfg["smtp_host"], cfg["smtp_port"],
                                  timeout=30)
    else:
        client = smtplib.SMTP(cfg["smtp_host"], cfg["smtp_port"], timeout=30)
        client.starttls()
    try:
        client.login(cfg["sender_email"], cfg["auth_code"])
        client.sendmail(cfg["sender_email"], to_addrs, msg.as_string())
    finally:
        try:
            client.quit()
        except Exception:
            pass
    summary["ok"] = True
    summary["message"] = "邮件已发送"
    return summary


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--config",
                    default=str(Path(__file__).resolve().parent.parent
                                / "config" / "smtp_config.json"))
    ap.add_argument("--to", action="append", required=True,
                    help="收件人邮箱，可重复传入多个")
    ap.add_argument("--subject", required=True, help="邮件主题")
    ap.add_argument("--body", default="", help="邮件正文")
    ap.add_argument("--attach", action="append", default=[],
                    help="附件文件路径，可重复传入多个")
    ap.add_argument("--dry-run", action="store_true",
                    help="只校验配置与参数，不真正发送")
    args = ap.parse_args()

    try:
        cfg = load_config(args.config)
        to_addrs = [a.strip() for a in args.to if a.strip()]
        if not to_addrs:
            raise ValueError("--to 不能为空")
        result = send(cfg, to_addrs, args.subject, args.body,
                      args.attach, dry_run=args.dry_run)
    except Exception as e:
        print(json.dumps({"ok": False, "error": str(e)},
                         ensure_ascii=False, indent=2))
        return 1

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())