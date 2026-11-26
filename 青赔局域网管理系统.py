# -*- coding: utf-8 -*-
"""青赔资料局域网试用版：使用浏览器访问，数据保存在本机青赔资料库。"""

from __future__ import annotations

import argparse
import base64
import hashlib
import hmac
import json
import mimetypes
import os
import secrets
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from html import escape
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse


def relaunch_with_bundled_python() -> None:
    """双击脚本时自动使用已配好的运行环境。"""
    runtime = (
        Path(os.environ.get("USERPROFILE", ""))
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "python"
        / "python.exe"
    )
    current = Path(sys.executable).resolve()
    if runtime.exists() and current != runtime.resolve():
        completed = subprocess.run([str(runtime), str(Path(__file__).resolve()), *sys.argv[1:]])
        raise SystemExit(completed.returncode)


relaunch_with_bundled_python()

import qingpei_db


ACCOUNT_PATH = qingpei_db.DATABASE_DIR / "局域网账号.json"
SESSION_SECONDS = 8 * 60 * 60
PORT = 8088
OFFICE_USER = {"username": "office", "display_name": "办公室资料库", "role": "admin"}
LOCAL_CSRF = secrets.token_urlsafe(24)
ROLES = {
    "admin": "管理员",
    "data": "资料员",
    "review": "审核员",
    "finance": "财务",
}
SESSIONS: dict[str, dict[str, object]] = {}
SESSION_LOCK = threading.Lock()


def password_hash(password: str, salt: bytes | None = None) -> tuple[str, str]:
    salt = salt or os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
    return base64.b64encode(salt).decode("ascii"), base64.b64encode(digest).decode("ascii")


def default_accounts() -> dict[str, dict[str, str]]:
    salt, digest = password_hash("qingpei123")
    return {
        "admin": {
            "display_name": "系统管理员",
            "role": "admin",
            "salt": salt,
            "password_hash": digest,
        }
    }


def load_accounts() -> dict[str, dict[str, str]]:
    qingpei_db.DATABASE_DIR.mkdir(parents=True, exist_ok=True)
    if not ACCOUNT_PATH.exists():
        accounts = default_accounts()
        ACCOUNT_PATH.write_text(json.dumps(accounts, ensure_ascii=False, indent=2), encoding="utf-8")
        return accounts
    try:
        accounts = json.loads(ACCOUNT_PATH.read_text(encoding="utf-8"))
        return accounts if isinstance(accounts, dict) else default_accounts()
    except (OSError, json.JSONDecodeError):
        return default_accounts()


def save_accounts(accounts: dict[str, dict[str, str]]) -> None:
    temp_path = ACCOUNT_PATH.with_suffix(".tmp")
    temp_path.write_text(json.dumps(accounts, ensure_ascii=False, indent=2), encoding="utf-8")
    temp_path.replace(ACCOUNT_PATH)


def check_password(username: str, password: str) -> dict[str, str] | None:
    account = load_accounts().get(username.lower())
    if not account:
        return None
    try:
        salt = base64.b64decode(account["salt"])
        _, digest = password_hash(password, salt)
    except (KeyError, ValueError):
        return None
    return account if hmac.compare_digest(digest, account.get("password_hash", "")) else None


def set_password(username: str, password: str) -> None:
    accounts = load_accounts()
    account = accounts[username]
    salt, digest = password_hash(password)
    account["salt"] = salt
    account["password_hash"] = digest
    save_accounts(accounts)


def local_ip() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except OSError:
        return socket.gethostbyname(socket.gethostname())


def clean_number(value: str) -> float:
    text = value.strip().replace("，", ".").replace(",", "")
    if not text:
        return 0.0
    amount = float(text)
    if amount < 0:
        raise ValueError("金额不能小于 0")
    return round(amount, 2)


def page(title: str, body: str, user: dict[str, str] | None = None, csrf: str = "") -> str:
    navigation = ""
    if user:
        role = user.get("role", "")
        items = ['<a href="/dashboard">资料库</a>']
        if role in {"admin", "data"}:
            items.append('<a href="/new">新增档案</a>')
        navigation = f"""
        <header><div class="brand">青赔资料库 <span>办公室局域网</span></div>
        <nav>{''.join(items)}</nav></header>"""
    return f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>{escape(title)}</title><style>
    *{{box-sizing:border-box}} body{{margin:0;background:#eef3f5;color:#1d2a32;font-family:"Microsoft YaHei UI","Microsoft YaHei",Arial,sans-serif;font-size:14px}}
    header{{min-height:64px;background:#13303a;color:#f6fbfc;display:flex;align-items:center;padding:0 28px;gap:26px;border-bottom:3px solid #18a58d}}
    .brand{{font-size:19px;font-weight:700;white-space:nowrap}} .brand span{{font-size:12px;color:#a7d9d0;margin-left:8px;font-weight:400}}
    nav{{display:flex;gap:5px;flex:1}} nav a{{color:#dcebea;text-decoration:none;padding:9px 12px;border:1px solid transparent}} nav a:hover{{border-color:#5f9e98;background:#1d4651}}
    .user{{display:flex;align-items:center;gap:8px;color:#c6e0dc;font-size:13px;white-space:nowrap}} .user form{{display:inline}}
    main{{max-width:1420px;margin:0 auto;padding:24px 28px 36px}} h1{{font-size:23px;margin:0 0 5px;color:#12313b}} h2{{font-size:16px;margin:0 0 14px;color:#12313b}}
    .hint{{color:#627781;margin:0 0 20px}} .notice{{padding:10px 12px;background:#e1f3ee;border-left:4px solid #18a58d;margin:14px 0;color:#174c42}}
    .warning{{padding:10px 12px;background:#fff4d7;border-left:4px solid #db9e24;margin:14px 0;color:#6a4a06}}
    .toolbar,.form-row{{display:flex;gap:10px;align-items:center;flex-wrap:wrap}} .toolbar{{margin:14px 0}} .push{{margin-left:auto}}
    input,select,textarea{{border:1px solid #9bb3b9;background:#fff;padding:8px 9px;color:#18282f;font:inherit;border-radius:3px}} input:focus,select:focus,textarea:focus{{outline:2px solid #70c9ba;border-color:#168e7b}}
    input[readonly],select:disabled{{background:#edf2f3;color:#5a6b70}} button,.button{{border:1px solid #137a6c;background:#168e7b;color:#fff;padding:8px 13px;border-radius:3px;font:inherit;cursor:pointer;text-decoration:none;display:inline-block}}
    button:hover,.button:hover{{background:#106f62}} button.secondary,.button.secondary{{background:#fff;color:#24535b;border-color:#88a7ae}} button.secondary:hover,.button.secondary:hover{{background:#e8f0f1}} button.link{{background:none;border:0;padding:0;color:#d9f3ef;text-decoration:underline}}
    .summary{{display:grid;grid-template-columns:repeat(4,minmax(130px,1fr));gap:10px;margin:14px 0 20px}} .summary div{{background:#fff;border:1px solid #c7d7da;padding:13px}} .summary b{{display:block;font-size:22px;color:#126f62;margin-top:4px}}
    .panel{{background:#fff;border:1px solid #c7d7da;padding:18px;margin-top:14px}} table{{width:100%;border-collapse:collapse;background:#fff}} th{{background:#dcebed;color:#173e46;text-align:left;font-weight:700;white-space:nowrap}} th,td{{padding:10px 9px;border-bottom:1px solid #d7e2e4;vertical-align:middle}} tr:hover td{{background:#f4faf9}} .table-wrap{{overflow:auto;border:1px solid #c7d7da}}
    .status{{display:inline-block;padding:3px 7px;background:#edf5f4;border:1px solid #b7d5d0;color:#275d55;white-space:nowrap}} .money{{font-variant-numeric:tabular-nums;text-align:right;white-space:nowrap}}
    .form-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:13px 18px;max-width:780px}} label{{display:flex;flex-direction:column;gap:6px;color:#38535b;font-weight:600}} label.full{{grid-column:1/-1}} .actions{{margin-top:18px;display:flex;gap:10px}}
    .login{{max-width:410px;margin:80px auto;background:#fff;border:1px solid #c7d7da;padding:28px}} .login h1{{margin-bottom:8px}} .login label{{margin-top:14px}} .login button{{width:100%;margin-top:20px}}
    footer{{margin-top:22px;color:#6a7f84;font-size:12px}} @media(max-width:820px){{header{{padding:14px;align-items:flex-start;flex-wrap:wrap;gap:12px}}nav{{order:3;flex-basis:100%;overflow:auto}}main{{padding:18px 12px}}.summary{{grid-template-columns:repeat(2,1fr)}}.form-grid{{grid-template-columns:1fr}}.user{{margin-left:auto}}}}
    </style></head><body>{navigation}<main>{body}<footer>资料仅在办公室局域网内使用。请定期备份 `_数据库` 文件夹。</footer></main></body></html>"""


def entry_page() -> str:
    return """<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>青赔资料库</title><style>
    *{box-sizing:border-box} body{margin:0;background:#07090b;color:#f4f5ef;font-family:"Microsoft YaHei UI","Microsoft YaHei",Arial,sans-serif}
    .entry{min-height:100vh;position:relative;display:grid;place-items:center;overflow:hidden;padding:36px;background:#07090b}
    .frame{position:absolute;inset:24px;border:1px solid #34352f;pointer-events:none}.frame::before,.frame::after{content:"";position:absolute;width:64px;height:64px;border-color:#d7ad3d;border-style:solid}.frame::before{top:-1px;left:-1px;border-width:2px 0 0 2px}.frame::after{right:-1px;bottom:-1px;border-width:0 2px 2px 0}
    .entry-main{position:relative;width:min(940px,100%);display:grid;grid-template-columns:360px minmax(0,1fr);align-items:center;gap:76px}
    .mark{position:relative;width:330px;height:330px;display:grid;place-items:center;justify-self:center}.orbit{position:absolute;border:1px solid #4e4b35;border-radius:50%}.orbit-one{inset:0;animation:turn 16s linear infinite}.orbit-two{inset:32px;border-color:#8a7131;animation:turn-reverse 11s linear infinite}.orbit-dot{width:9px;height:9px;background:#e3b943;border-radius:50%;position:absolute;top:16px;left:54%;box-shadow:0 0 18px #e3b943}.logo-wrap{width:224px;height:224px;border:1px solid #d7ad3d;border-radius:50%;overflow:hidden;background:#050505;position:relative;z-index:1;box-shadow:0 0 0 10px #10110f,0 0 34px #bd9130}.logo-wrap img{display:block;width:100%;height:100%;object-fit:cover}.scan{position:absolute;left:48px;right:48px;height:1px;background:#f0c54a;z-index:2;animation:scan 3.8s ease-in-out infinite;box-shadow:0 0 13px #f0c54a}
    .eyebrow{color:#e6bc45;font-size:12px;letter-spacing:.18em;margin:0 0 18px}.entry h1{margin:0;color:#f7f7f1;font-size:42px;line-height:1.12;font-weight:700;letter-spacing:0}.entry p{max-width:430px;margin:18px 0 30px;color:#acafa6;font-size:16px;line-height:1.8}.enter{display:inline-flex;align-items:center;gap:18px;padding:14px 18px 14px 22px;border:1px solid #e1b743;background:#d6ac3d;color:#11120e;text-decoration:none;font-weight:700;transition:background .2s,transform .2s}.enter:hover{background:#f0c856;transform:translateY(-2px)}.enter span{display:grid;place-items:center;width:28px;height:28px;border:1px solid #3d3211;border-radius:50%;font-size:20px}.status{display:flex;gap:10px;align-items:center;margin-top:26px;color:#777a71;font-size:12px;letter-spacing:.08em}.status i{width:8px;height:8px;border-radius:50%;background:#66bf91;box-shadow:0 0 12px #66bf91}
    @keyframes turn{to{transform:rotate(360deg)}}@keyframes turn-reverse{to{transform:rotate(-360deg)}}@keyframes scan{0%,100%{top:54px;opacity:0}15%,85%{opacity:1}50%{top:276px;opacity:1}}
    @media(max-width:760px){.entry{padding:24px}.frame{inset:12px}.entry-main{grid-template-columns:1fr;gap:30px;text-align:center}.mark{width:270px;height:270px}.logo-wrap{width:184px;height:184px}.orbit-two{inset:26px}.scan{left:42px;right:42px}.entry h1{font-size:34px}.entry p{margin-left:auto;margin-right:auto}.status{justify-content:center}@keyframes scan{0%,100%{top:46px;opacity:0}15%,85%{opacity:1}50%{top:224px;opacity:1}}}
    @media(prefers-reduced-motion:reduce){.orbit,.scan{animation:none}.scan{top:50%;opacity:1}}
    </style></head><body><main class="entry"><div class="frame"></div><section class="entry-main"><div class="mark" aria-hidden="true"><div class="orbit orbit-one"><i class="orbit-dot"></i></div><div class="orbit orbit-two"></div><div class="scan"></div><div class="logo-wrap"><img src="/assets/entry-logo.jpg" alt=""></div></div><div><p class="eyebrow">QINGPEI ARCHIVE SYSTEM</p><h1>青赔资料库</h1><p>办公室局域网资料管理工作台</p><a class="enter" href="/dashboard">进入资料库 <span>→</span></a><div class="status"><i></i>系统已就绪</div></div></section></main></body></html>"""


class QingpeiHandler(BaseHTTPRequestHandler):
    server_version = "QingpeiLAN/1.0"

    def log_message(self, _format: str, *_args) -> None:
        return

    def session(self) -> tuple[dict[str, str] | None, str]:
        cookie = SimpleCookie(self.headers.get("Cookie", ""))
        token = cookie.get("qingpei_session")
        if not token:
            return None, ""
        with SESSION_LOCK:
            item = SESSIONS.get(token.value)
            if not item or float(item.get("expires", 0)) < time.time():
                SESSIONS.pop(token.value, None)
                return None, ""
            username = str(item["username"])
            account = load_accounts().get(username)
            if not account:
                return None, ""
            return {"username": username, **account}, str(item["csrf"])

    def create_session(self, username: str) -> tuple[str, str]:
        token = secrets.token_urlsafe(32)
        csrf = secrets.token_urlsafe(24)
        with SESSION_LOCK:
            SESSIONS[token] = {"username": username, "csrf": csrf, "expires": time.time() + SESSION_SECONDS}
        return token, csrf

    def send_html(self, content: str, status: HTTPStatus = HTTPStatus.OK, cookie: str = "") -> None:
        data = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        if cookie:
            self.send_header("Set-Cookie", cookie)
        self.end_headers()
        self.wfile.write(data)

    def send_entry_logo(self) -> None:
        logo_path = qingpei_db.BASE_DIR / "assets" / "entry-logo.jpg"
        if not logo_path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        data = logo_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "image/jpeg")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def redirect(self, path: str, cookie: str = "") -> None:
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", path)
        if cookie:
            self.send_header("Set-Cookie", cookie)
        self.end_headers()

    def message_redirect(self, path: str, message: str) -> None:
        joiner = "&" if "?" in path else "?"
        self.redirect(path + joiner + urlencode({"msg": message}))

    def form_data(self) -> dict[str, str]:
        length = int(self.headers.get("Content-Length", "0"))
        if length > 200_000:
            raise ValueError("提交内容过大")
        raw = self.rfile.read(length).decode("utf-8")
        return {key: values[-1] for key, values in parse_qs(raw, keep_blank_values=True).items()}

    def require_user(self) -> tuple[dict[str, str] | None, str]:
        return dict(OFFICE_USER), LOCAL_CSRF

    def require_role(self, roles: set[str]) -> tuple[dict[str, str] | None, str]:
        user, csrf = self.require_user()
        if user is not None and user.get("role") not in roles:
            self.send_html(page("无权限", '<h1>没有权限</h1><p class="hint">请使用有权限的账号操作。</p>', user, csrf), HTTPStatus.FORBIDDEN)
            return None, ""
        return user, csrf

    def verify_csrf(self, data: dict[str, str], csrf: str) -> bool:
        return bool(csrf) and hmac.compare_digest(data.get("csrf", ""), csrf)

    def do_GET(self) -> None:
        route = urlparse(self.path)
        if route.path == "/health":
            payload = b"ok"
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/plain")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
            return
        if route.path == "/assets/entry-logo.jpg":
            self.send_entry_logo()
        elif route.path in {"/", "/login"}:
            self.show_entry(route)
        elif route.path == "/dashboard":
            self.show_dashboard(route)
        elif route.path == "/record":
            self.show_record(route)
        elif route.path == "/new":
            self.show_new()
        elif route.path == "/export":
            self.export_records()
        elif route.path == "/file":
            self.download_attachment(route)
        else:
            self.send_html(page("未找到页面", "<h1>页面不存在</h1>"), HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        route = urlparse(self.path)
        try:
            data = self.form_data()
        except ValueError as exc:
            self.send_html(page("提交失败", f"<h1>提交失败</h1><p>{escape(str(exc))}</p>"), HTTPStatus.BAD_REQUEST)
            return
        if route.path == "/record/save":
            self.save_record(data)
        elif route.path == "/new":
            self.create_record(data)
        elif route.path == "/scan":
            self.scan_records(data)
        else:
            self.send_html(page("未找到页面", "<h1>页面不存在</h1>"), HTTPStatus.NOT_FOUND)

    def show_entry(self, _route) -> None:
        self.send_html(entry_page())

    def login(self, data: dict[str, str]) -> None:
        username = data.get("username", "").strip().lower()
        account = check_password(username, data.get("password", ""))
        if account is None:
            self.redirect("/login?" + urlencode({"msg": "账号或密码不正确"}))
            return
        token, _ = self.create_session(username)
        self.redirect("/", f"qingpei_session={token}; Path=/; HttpOnly; SameSite=Lax")

    def logout(self, data: dict[str, str]) -> None:
        _, csrf = self.session()
        if not self.verify_csrf(data, csrf):
            self.redirect("/login")
            return
        cookie = SimpleCookie(self.headers.get("Cookie", ""))
        token = cookie.get("qingpei_session")
        if token:
            with SESSION_LOCK:
                SESSIONS.pop(token.value, None)
        self.redirect("/login", "qingpei_session=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax")

    def show_dashboard(self, route) -> None:
        user, csrf = self.require_user()
        if user is None:
            return
        query = parse_qs(route.query)
        keyword = query.get("q", [""])[-1]
        message = query.get("msg", [""])[-1]
        rows = qingpei_db.list_records(keyword)
        all_rows = qingpei_db.list_records()
        recognized_towers = sum(bool(str(row["tower_ids"] or "").strip()) for row in all_rows)
        paid_count = sum(row["payment_status"] == "已支付" for row in all_rows)
        body_rows = []
        for row in rows:
            person_id = int(row["id"])
            location = row["address"] or qingpei_db.build_address(row["town"], row["village"], row["team"])
            body_rows.append(f'''<tr><td><a href="/record?id={person_id}">{escape(row['name'])}</a></td><td>{escape(row['tower_ids'])}</td>
            <td>{escape(location)}</td><td class="money">{float(row['total_amount'] or 0):,.2f}</td><td><span class="status">{escape(row['sign_status'])}</span></td>
            <td><span class="status">{escape(row['payment_status'])}</span></td><td><span class="status">{escape(row['archive_status'])}</span></td><td>{row['file_count']}</td><td>{escape(row['updated_at'])}</td></tr>''')
        message_html = f'<div class="notice">{escape(message)}</div>' if message else ""
        actions = ''
        if user["role"] == "admin":
            actions = f'''<form method="post" action="/scan"><input type="hidden" name="csrf" value="{escape(csrf)}"><button class="secondary" type="submit">扫描旧资料</button></form>
            <a class="button secondary" href="/export">导出汇总</a>'''
        elif user["role"] in {"review", "finance"}:
            actions = '<a class="button secondary" href="/export">导出汇总</a>'
        body = f'''<h1>资料库</h1><p class="hint">查询人员、塔号、补偿金额和流转状态。双击查看资料详情并按权限修改。</p>{message_html}
        <section class="summary"><div>登记档案<b>{len(all_rows)}</b></div><div>已识别塔号<b>{recognized_towers}</b></div><div>已支付<b>{paid_count}</b></div><div>当前查询结果<b>{len(rows)}</b></div></section>
        <form class="toolbar" method="get" action="/dashboard"><input name="q" value="{escape(keyword)}" placeholder="姓名、塔号、村或资料文件夹" style="min-width:280px;flex:1"><button type="submit">查询</button><a class="button secondary" href="/dashboard">清除</a><span class="push">{actions}</span></form>
        <div class="table-wrap"><table><thead><tr><th>姓名</th><th>铁塔编号</th><th>位置</th><th>补偿总额</th><th>签字</th><th>支付</th><th>归档</th><th>资料</th><th>更新时间</th></tr></thead>
        <tbody>{''.join(body_rows) or '<tr><td colspan="9">没有找到资料。</td></tr>'}</tbody></table></div>'''
        self.send_html(page("资料库", body, user, csrf))

    def find_row(self, value: str):
        try:
            person_id = int(value)
        except ValueError:
            return None
        return next((row for row in qingpei_db.list_records() if int(row["id"]) == person_id), None)

    def option_list(self, selected: str, options: list[str]) -> str:
        return "".join(f'<option value="{escape(item)}" {"selected" if item == selected else ""}>{escape(item)}</option>' for item in options)

    def show_record(self, route) -> None:
        user, csrf = self.require_user()
        if user is None:
            return
        row = self.find_row(parse_qs(route.query).get("id", [""])[-1])
        if row is None:
            self.message_redirect("/dashboard", "未找到该资料。")
            return
        can_data = user["role"] in {"admin", "data"}
        can_review = user["role"] in {"admin", "review"}
        can_finance = user["role"] in {"admin", "finance"}
        readonly = "" if can_data else "readonly"
        review_disabled = "" if can_review else "disabled"
        finance_disabled = "" if can_finance else "disabled"
        message = parse_qs(route.query).get("msg", [""])[-1]
        message_html = f'<div class="notice">{escape(message)}</div>' if message else ""
        folder = escape(row["folder_path"] or "暂未生成资料文件夹")
        location = escape(row["address"] or qingpei_db.build_address(row["town"], row["village"], row["team"]))
        total_source = escape(row["total_source"] or "未标注")
        attachments = qingpei_db.list_attachments(int(row["id"]))
        agreements = qingpei_db.list_agreements(int(row["id"]))
        agreement_rows = []
        for index, agreement in enumerate(agreements, start=1):
            path = Path(agreement["file_path"])
            tower_ids = escape(agreement["tower_ids"] or "暂未识别")
            amount = f"{float(agreement['total_amount']):,.2f}" if agreement["total_amount"] is not None else "暂未识别"
            action = f'<a class="button secondary" href="/file?id={agreement["id"]}" target="_blank">打开/下载</a>' if agreement["exists_flag"] else '<span class="status">文件已移动</span>'
            agreement_rows.append(
                f'<tr><td>协议 {index}</td><td>{tower_ids}</td><td class="money">{amount}</td><td>{escape(path.name)}</td><td>{action}</td></tr>'
            )
        agreement_total = sum(float(item["total_amount"]) for item in agreements if item["total_amount"] is not None)
        agreements_panel = f'''<section class="panel"><h2>协议清单（{len(agreements)} 份）</h2><p class="hint">每份协议独立列出，协议金额合计：<b>{agreement_total:,.2f} 元</b>。</p>
        <div class="table-wrap"><table><thead><tr><th>序号</th><th>铁塔编号</th><th>协议金额（元）</th><th>协议文件</th><th>操作</th></tr></thead><tbody>{''.join(agreement_rows) or '<tr><td colspan="5">暂未识别到补偿协议。</td></tr>'}</tbody></table></div></section>'''
        file_rows = []
        for attachment in attachments:
            if attachment["file_type"] == "补偿协议":
                continue
            path = Path(attachment["file_path"])
            available = bool(attachment["exists_flag"]) and path.exists()
            action = f'<a class="button secondary" href="/file?id={attachment["id"]}" target="_blank">打开/下载</a>' if available else '<span class="status">文件已移动</span>'
            file_rows.append(
                f'<tr><td>{escape(attachment["file_type"])}</td><td>{escape(path.name)}</td><td>{action}</td></tr>'
            )
        files_panel = f'''<section class="panel"><h2>其余资料文件</h2><p class="hint">点击“打开/下载”后，Word 和 Excel 会下载到当前电脑，可直接用本机 Office 或 WPS 打开。</p>
        <div class="table-wrap"><table><thead><tr><th>资料类型</th><th>文件名称</th><th>操作</th></tr></thead><tbody>{''.join(file_rows) or '<tr><td colspan="3">暂未关联资料文件。请在原自动录入系统生成后重新扫描旧资料。</td></tr>'}</tbody></table></div></section>'''
        body = f'''<h1>{escape(row['name'])}</h1><p class="hint">地址：{location or '暂未识别'}<br>补偿总额来源：{total_source}<br>资料文件夹：<code>{folder}</code></p>{message_html}
        <section class="panel"><h2>档案、身份与收款信息</h2><form method="post" action="/record/save"><input type="hidden" name="csrf" value="{escape(csrf)}"><input type="hidden" name="id" value="{row['id']}">
        <div class="form-grid"><label class="full">完整地址<input name="address" value="{escape(row['address'])}" {readonly}></label>
        <label>镇<input name="town" value="{escape(row['town'])}" {readonly}></label><label>村<input name="village" value="{escape(row['village'])}" {readonly}></label><label>队<input name="team" value="{escape(row['team'])}" {readonly}></label>
        <label>铁塔编号<input name="tower_ids" value="{escape(row['tower_ids'])}" {readonly}></label>
        <label>补偿总额（元）<input name="total_amount" inputmode="decimal" value="{float(row['total_amount'] or 0):.2f}" {readonly}></label>
        <label>身份证号码<input name="id_number" value="{escape(row['id_number'])}" {readonly}></label>
        <label>开户银行<input name="bank_name" value="{escape(row['bank_name'])}" {readonly}></label>
        <label>银行卡号<input name="bank_card" inputmode="numeric" value="{escape(row['bank_card'])}" {readonly}></label>
        <label>签字状态<select name="sign_status" {review_disabled}>{self.option_list(row['sign_status'], ['待签字', '已签字'])}</select></label>
        <label>支付状态<select name="payment_status" {finance_disabled}>{self.option_list(row['payment_status'], ['未支付', '支付中', '已支付'])}</select></label>
        <label>归档状态<select name="archive_status" {review_disabled}>{self.option_list(row['archive_status'], ['待归档', '资料不全', '已归档'])}</select></label>
        <label class="full">缺少资料或审核备注<textarea name="missing_notes" rows="3" {'readonly' if not can_review and not can_data else ''}>{escape(row['missing_notes'])}</textarea></label></div>
        <div class="actions"><button type="submit">保存修改</button><a class="button secondary" href="/dashboard">返回资料库</a></div></form></section>{agreements_panel}{files_panel}'''
        self.send_html(page("档案详情", body, user, csrf))

    def save_record(self, data: dict[str, str]) -> None:
        user, csrf = self.require_user()
        if user is None or not self.verify_csrf(data, csrf):
            self.redirect("/login")
            return
        row = self.find_row(data.get("id", ""))
        if row is None:
            self.message_redirect("/dashboard", "未找到该资料。")
            return
        role = user["role"]
        try:
            if role in {"admin", "data"}:
                qingpei_db.update_compensation(int(row["id"]), data.get("tower_ids", ""), clean_number(data.get("total_amount", "")))
                qingpei_db.update_person_location(
                    int(row["id"]), data.get("address", ""), data.get("town", ""), data.get("village", ""), data.get("team", ""),
                )
                qingpei_db.update_person_identity(
                    int(row["id"]), data.get("id_number", ""), data.get("bank_name", ""), data.get("bank_card", ""),
                )
            sign_status = data.get("sign_status", row["sign_status"]) if role in {"admin", "review"} else row["sign_status"]
            payment_status = data.get("payment_status", row["payment_status"]) if role in {"admin", "finance"} else row["payment_status"]
            archive_status = data.get("archive_status", row["archive_status"]) if role in {"admin", "review"} else row["archive_status"]
            notes = data.get("missing_notes", row["missing_notes"]) if role in {"admin", "data", "review"} else row["missing_notes"]
            qingpei_db.update_status(int(row["id"]), sign_status, payment_status, archive_status, notes)
        except ValueError:
            self.message_redirect(f"/record?id={row['id']}", "补偿总额请填写正确的数字。")
            return
        self.message_redirect(f"/record?id={row['id']}", "已保存。")

    def show_new(self) -> None:
        user, csrf = self.require_role({"admin", "data"})
        if user is None:
            return
        body = f'''<h1>新增简要档案</h1><p class="hint">先登记姓名、塔号和金额；完整协议、证件和补偿表仍可通过原“自动录入系统”生成并自动合并。</p>
        <section class="panel"><form method="post" action="/new"><input type="hidden" name="csrf" value="{escape(csrf)}"><div class="form-grid">
        <label>姓名<input name="name" required></label><label>铁塔编号<input name="tower_ids"></label><label>补偿总额（元）<input name="total_amount" inputmode="decimal"></label>
        <label>镇<input name="town"></label><label>村<input name="village"></label><label>队<input name="team"></label>
        <label class="full">资料备注<textarea name="missing_notes" rows="3"></textarea></label></div><div class="actions"><button type="submit">保存档案</button><a class="button secondary" href="/dashboard">取消</a></div></form></section>'''
        self.send_html(page("新增档案", body, user, csrf))

    def create_record(self, data: dict[str, str]) -> None:
        user, csrf = self.require_role({"admin", "data"})
        if user is None or not self.verify_csrf(data, csrf):
            self.redirect("/login")
            return
        try:
            person_id = qingpei_db.create_web_record(
                data.get("name", ""), data.get("tower_ids", ""), clean_number(data.get("total_amount", "")),
                data.get("town", ""), data.get("village", ""), data.get("team", ""), data.get("missing_notes", ""),
            )
        except ValueError as exc:
            self.message_redirect("/new", str(exc))
            return
        self.message_redirect(f"/record?id={person_id}", "新档案已保存。")

    def scan_records(self, data: dict[str, str]) -> None:
        user, csrf = self.require_role({"admin"})
        if user is None or not self.verify_csrf(data, csrf):
            self.redirect("/login")
            return
        imported, files_count = qingpei_db.scan_existing_folders()
        self.message_redirect("/dashboard", f"扫描完成：新增 {imported} 户，关联 {files_count} 个资料文件。")

    def show_users(self, route) -> None:
        user, csrf = self.require_role({"admin"})
        if user is None:
            return
        message = parse_qs(route.query).get("msg", [""])[-1]
        rows = "".join(
            f"<tr><td>{escape(username)}</td><td>{escape(account.get('display_name', ''))}</td><td>{escape(ROLES.get(account.get('role', ''), ''))}</td>"
            f"<td><form class=\"form-row\" method=\"post\" action=\"/users\"><input type=\"hidden\" name=\"csrf\" value=\"{escape(csrf)}\"><input type=\"hidden\" name=\"action\" value=\"reset\"><input type=\"hidden\" name=\"username\" value=\"{escape(username)}\"><input name=\"password\" type=\"password\" placeholder=\"重设密码\" required><button class=\"secondary\" type=\"submit\">重设</button></form></td></tr>"
            for username, account in sorted(load_accounts().items())
        )
        notice = f'<div class="notice">{escape(message)}</div>' if message else ""
        body = f'''<h1>账号管理</h1><p class="hint">请为每位同事建立独立账号，不要共用管理员账号。</p>{notice}
        <section class="panel"><h2>现有账号</h2><div class="table-wrap"><table><thead><tr><th>账号</th><th>姓名</th><th>角色</th><th>重设密码</th></tr></thead><tbody>{rows}</tbody></table></div></section>
        <section class="panel"><h2>新增账号</h2><form method="post" action="/users"><input type="hidden" name="csrf" value="{escape(csrf)}"><input type="hidden" name="action" value="add"><div class="form-grid">
        <label>登录账号<input name="username" pattern="[A-Za-z0-9_]{3,20}" required></label><label>姓名或称呼<input name="display_name" required></label>
        <label>角色<select name="role">{self.option_list('data', ['data', 'review', 'finance', 'admin'])}</select></label><label>初始密码<input name="password" type="password" minlength="8" required></label>
        </div><div class="actions"><button type="submit">创建账号</button></div></form></section>'''
        self.send_html(page("账号管理", body, user, csrf))

    def save_user(self, data: dict[str, str]) -> None:
        user, csrf = self.require_role({"admin"})
        if user is None or not self.verify_csrf(data, csrf):
            self.redirect("/login")
            return
        accounts = load_accounts()
        username = data.get("username", "").strip().lower()
        password = data.get("password", "")
        if len(password) < 8:
            self.message_redirect("/users", "密码至少需要 8 位。")
            return
        if data.get("action") == "reset":
            if username not in accounts:
                self.message_redirect("/users", "账号不存在。")
                return
            set_password(username, password)
            self.message_redirect("/users", "密码已重设。")
            return
        if not username.replace("_", "").isalnum() or not 3 <= len(username) <= 20 or username in accounts:
            self.message_redirect("/users", "账号须为 3 到 20 位字母、数字或下划线，且不能重复。")
            return
        role = data.get("role", "data")
        if role not in ROLES:
            self.message_redirect("/users", "角色不正确。")
            return
        salt, digest = password_hash(password)
        accounts[username] = {"display_name": data.get("display_name", "").strip() or username, "role": role, "salt": salt, "password_hash": digest}
        save_accounts(accounts)
        self.message_redirect("/users", "账号已创建。")

    def show_password(self, route) -> None:
        user, csrf = self.require_user()
        if user is None:
            return
        message = parse_qs(route.query).get("msg", [""])[-1]
        notice = f'<div class="notice">{escape(message)}</div>' if message else ""
        body = f'''<h1>修改密码</h1>{notice}<section class="panel"><form method="post" action="/password"><input type="hidden" name="csrf" value="{escape(csrf)}">
        <div class="form-grid"><label>当前密码<input name="old_password" type="password" required></label><label>新密码（至少 8 位）<input name="new_password" type="password" minlength="8" required></label></div>
        <div class="actions"><button type="submit">保存新密码</button></div></form></section>'''
        self.send_html(page("修改密码", body, user, csrf))

    def change_own_password(self, data: dict[str, str]) -> None:
        user, csrf = self.require_user()
        if user is None or not self.verify_csrf(data, csrf):
            self.redirect("/login")
            return
        if check_password(user["username"], data.get("old_password", "")) is None:
            self.message_redirect("/password", "当前密码不正确。")
            return
        new_password = data.get("new_password", "")
        if len(new_password) < 8:
            self.message_redirect("/password", "新密码至少需要 8 位。")
            return
        set_password(user["username"], new_password)
        self.message_redirect("/password", "密码已修改。")

    def export_records(self) -> None:
        user, _ = self.require_role({"admin", "review", "finance"})
        if user is None:
            return
        output_path = qingpei_db.export_records()
        data = output_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Disposition", "attachment; filename=qingpei_summary.csv")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def download_attachment(self, route) -> None:
        user, _ = self.require_user()
        if user is None:
            return
        try:
            attachment_id = int(parse_qs(route.query).get("id", [""])[-1])
        except ValueError:
            self.send_html(page("文件不存在", "<h1>文件不存在</h1>", user), HTTPStatus.NOT_FOUND)
            return
        attachment = qingpei_db.get_attachment(attachment_id)
        if attachment is None:
            self.send_html(page("文件不存在", "<h1>文件不存在</h1>", user), HTTPStatus.NOT_FOUND)
            return
        path = Path(attachment["file_path"])
        try:
            path.resolve().relative_to(qingpei_db.BASE_DIR.resolve())
        except ValueError:
            self.send_html(page("无法访问文件", "<h1>文件路径不安全</h1>", user), HTTPStatus.FORBIDDEN)
            return
        if not path.is_file():
            self.send_html(page("文件不存在", "<h1>文件可能已移动或删除</h1>", user), HTTPStatus.NOT_FOUND)
            return
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", "attachment; filename*=UTF-8''" + __import__("urllib.parse").parse.quote(path.name))
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main() -> int:
    parser = argparse.ArgumentParser(description="青赔资料局域网管理系统")
    parser.add_argument("--port", type=int, default=PORT)
    parser.add_argument("--no-browser", action="store_true", help="启动后不自动打开浏览器")
    args = parser.parse_args()
    qingpei_db.initialize()
    load_accounts()
    server = ThreadingHTTPServer(("0.0.0.0", args.port), QingpeiHandler)
    address = f"http://{local_ip()}:{args.port}"
    print("青赔资料局域网管理系统已启动")
    print(f"本机和办公室同事访问：{address}")
    print("按 Ctrl+C 可停止服务。")
    if not args.no_browser:
        threading.Timer(0.4, lambda: webbrowser.open(f"http://127.0.0.1:{args.port}")).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
