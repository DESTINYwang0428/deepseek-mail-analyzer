#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
deepseek_mail_analyzer.py — 今日邮件智能分析（重要度 + 传达含义 + 摘要 + 原文）

============================================================================
本脚本的特点：
  1. 完全不经过 DSH，不调用 DS Agent，不消耗任何 DSH / Agent token；
     唯一的外部调用是 DeepSeek 官方大语言模型 API（https://api.deepseek.com），
     计费走你自己的 DeepSeek API Key（在 https://platform.deepseek.com 申请）。
  2. 通过 IMAP 协议直连你的邮箱，读取“今日”邮件（也可改为最近 N 小时）。
  3. 对每封邮件输出：重要度评分(1-5)、分类标签、传达含义、摘要、建议动作、
     发件人可信度；对“重要”邮件（重要度 >= 阈值）自动附上完整原文。
  4. 纯 Python 标准库实现（Python 3.8+），无需 pip 安装任何第三方包。

常用命令：
    python3 deepseek_mail_analyzer.py                    # 读今日邮件并分析（自动缓存，重复邮件不再消耗 token）
    python3 deepseek_mail_analyzer.py --demo             # 用内置示例邮件（不连邮箱）
    python3 deepseek_mail_analyzer.py --demo --dry-run   # 示例邮件 + 不调用 API
    python3 deepseek_mail_analyzer.py --hours 48         # 最近 48 小时
    python3 deepseek_mail_analyzer.py --threshold 3      # 重要度 >=3 视为重要并给原文
    python3 deepseek_mail_analyzer.py --unread-only      # 只看未读邮件
    python3 deepseek_mail_analyzer.py --important-only   # 规则预过滤，订阅/营销类不调用 API
    python3 deepseek_mail_analyzer.py --no-cache         # 关闭缓存（每次都重新分析）
    python3 deepseek_mail_analyzer.py --no-overview      # 不生成今日总览（省一次调用）
    python3 deepseek_mail_analyzer.py --cache-clean      # 手动清理缓存（移除超过 30 天的旧记录）
    python3 deepseek_mail_analyzer.py --cache-clean --cache-max-age 30   # 按 30 天清理
    python3 deepseek_mail_analyzer.py --cache-clear      # 彻底清空缓存
    python3 deepseek_mail_analyzer.py --desktop-report   # 生成 HTML 报告到桌面
    python3 deepseek_mail_analyzer.py --popup            # macOS 弹窗：通知 + 有重要邮件自动打开报告
    python3 deepseek_mail_analyzer.py --json > r.json    # 输出原始 JSON

省 token 说明：
    * 读邮件（IMAP）本身 0 token；token 只花在“把邮件发给 LLM 分析”和“生成总览”。
    * 脚本按 Message-ID 缓存分析结果：同一封邮件只分析一次，之后每天重复运行 0 新增 token。
    * --important-only 会用本地规则跳过订阅/营销/周刊类邮件，不调用 API。
    * 建议在非高峰时段（每天 9:00-14:00 之外）运行，DeepSeek V4 错峰价约为高峰的一半。

配置优先级（低 -> 高）：
    本文件顶部 CONFIG 区 -> .env 文件 -> 环境变量 -> 命令行参数
============================================================================
"""

import argparse
import datetime as dt
import email
import imaplib
import json
import os
import re
import socket
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
from html import unescape

# ============================================================
# 微软 OAuth2（XOAUTH2 / 设备代码流）：
# 微软 2022 年起禁用 Office365 的“账号密码”登录 IMAP，学校/企业邮箱
# （如 qmul.ac.uk）必须走 OAuth2：浏览器授权一次拿令牌，之后自动续期。
# 默认用 Thunderbird 注册的多租户公共客户端 ID（开源邮件客户端通用做法），一般无需修改。
# ============================================================
OAUTH_DEFAULT_CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
OAUTH_SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"

# ============================================================
# CONFIG 区：在这里修改默认配置（也可用 .env / 环境变量 / 命令行参数覆盖）
# ============================================================
DEFAULTS = {
    # ---- DeepSeek API（平台：platform.deepseek.com，费用走你自己的 Key）----
    "api_key": "sk-在此填入你的DeepSeekAPIKey",   # 环境变量: DEEPSEEK_API_KEY
    "api_base": "https://api.deepseek.com",       # 环境变量: DEEPSEEK_BASE_URL
    "model": "deepseek-chat",                     # 环境变量: DEEPSEEK_MODEL

    # ---- 邮箱（IMAP）----
    # 常用服务器：QQ 邮箱 imap.qq.com / 163 imap.163.com / Gmail imap.gmail.com
    #            / Outlook outlook.office365.com / 126 imap.126.com
    "imap_host": "imap.qq.com",       # 环境变量: IMAP_HOST
    "imap_port": 993,                 # 环境变量: IMAP_PORT（993 为 SSL，143 为明文/STARTTLS）
    "imap_user": "you@example.com",   # 环境变量: IMAP_USER
    # 注意：QQ/163/Gmail 等通常不能用登录密码，需在邮箱设置里开启 IMAP 并
    #       生成“授权码 / 应用专用密码”填写到这里。
    "imap_password": "你的IMAP授权码",  # 环境变量: IMAP_PASSWORD
    "imap_ssl": True,                 # 环境变量: IMAP_SSL（1/0）
    "imap_folder": "INBOX",           # 环境变量: IMAP_FOLDER

    # ---- 微软 Outlook / 学校邮箱 OAuth2（微软已禁用密码登录 IMAP，必须用 OAuth）----
    "imap_auth": "password",          # password | oauth2。环境变量: IMAP_AUTH
    "oauth_client_id": OAUTH_DEFAULT_CLIENT_ID,  # 环境变量: OAUTH_CLIENT_ID（一般不用改）
    "oauth_tenant": "common",         # 环境变量: OAUTH_TENANT
    "oauth_token_file": ".outlook_oauth.json",  # 环境变量: OAUTH_TOKEN_FILE

    # ---- 分析行为 ----
    "lookback_hours": None,   # None = 今日 0 点起；填数字 = 最近 N 小时。环境变量: LOOKBACK_HOURS
    "max_emails": 100,        # 最多处理多少封。环境变量: MAX_EMAILS
    "threshold": 4,           # 重要度 >= 该值视为“重要”，输出原文。环境变量: IMPORTANT_THRESHOLD
    "body_max": 6000,         # 送入 API 的每封邮件正文上限（字符）。环境变量: BODY_MAX
    "print_body_cap": 8000,   # 终端打印原文时的字符上限（--full-body 可取消）
    "batch_size": 20,         # 批量分析时每批邮件数。环境变量: BATCH_SIZE

    # ---- 省 token 机制 ----
    # 按 Message-ID 缓存分析结果，重复邮件不再调用 API（--no-cache 关闭）
    "cache_file": ".mail_analyzer_state.json",   # 环境变量: CACHE_FILE
    # 缓存自动清理：超过该天数的旧分析记录会在每次运行时自动移除
    "cache_max_age_days": 30,                    # 环境变量: CACHE_MAX_AGE_DAYS
    # --important-only 时按发件人域名/关键词预过滤：命中以下子串的发件人跳过 API
    "skip_senders": "newsletter,weekly,digest,promo,bounce,unsubscribe,mailer,edm,marketing",  # 环境变量: SKIP_SENDERS

    # ---- 费用估算（元/百万 tokens，默认按 deepseek V4-Flash 空闲时段价；高峰 9:00-14:00 约翻倍）----
    "price_input": 1.5,       # 环境变量: DEEPSEEK_PRICE_INPUT
    "price_output": 4.5,      # 环境变量: DEEPSEEK_PRICE_OUTPUT

    # ---- 桌面报告 / 弹窗（--desktop-report / --popup）----
    "desktop_dir": "~/Desktop",   # 报告保存目录。环境变量: DESKTOP_DIR
}

# .env 文件名（与脚本同目录，或 --env 指定）
DEFAULT_ENV_FILE = ".env"


# ============================================================
# 工具函数
# ============================================================

def load_dotenv(path):
    """极简 .env 加载器（不覆盖已存在的环境变量），无第三方依赖。"""
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k, v = k.strip(), v.strip().strip('"').strip("'")
            if k:
                os.environ.setdefault(k, v)


def strtobool(s):
    return str(s).strip().lower() in ("1", "true", "yes", "on", "y")


def int_or_none(s):
    s = str(s).strip().lower()
    if s in ("", "none", "today", "null"):
        return None
    return int(s)


_CODES = {
    "reset": "0", "bold": "1", "dim": "2",
    "red": "31", "green": "32", "yellow": "33",
    "blue": "34", "magenta": "35", "cyan": "36",
}


def col(text, code, use_color):
    if not use_color:
        return text
    return "\033[%sm%s\033[0m" % (_CODES[code], text)


def decode_mime(s):
    """解码邮件头（Subject / From 等，可能是 base64/quoted-printable + 多编码）。"""
    if not s:
        return ""
    try:
        return str(make_header(decode_header(s)))
    except Exception:
        return s


def decode_part(part):
    payload = part.get_payload(decode=True)
    if payload is None:
        payload = part.get_payload()
        if isinstance(payload, list):
            payload = ""
    charset = part.get_content_charset() or "utf-8"
    try:
        return payload.decode(charset, errors="replace")
    except Exception:
        return payload.decode("utf-8", errors="replace")


_TAG_RE = re.compile(r"<[^>]+>")


def strip_html(html):
    """把 HTML 正文粗略转成纯文本（去掉脚本/样式/标签/实体）。"""
    if not html:
        return ""
    html = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", html)
    html = re.sub(r"(?s)<!--.*?-->", " ", html)
    html = _TAG_RE.sub(" ", html)
    html = unescape(html)
    html = re.sub(r"[ \t\r\f\v]+", " ", html)
    html = re.sub(r"\n\s*\n+", "\n", html)
    return html.strip()


def get_body(msg):
    """返回 (纯文本正文, html正文)。优先 text/plain，否则用 text/html 转纯文本。"""
    text_parts, html_parts = [], []
    try:
        if msg.is_multipart():
            for part in msg.walk():
                cd = str(part.get("Content-Disposition") or "")
                if "attachment" in cd:
                    continue
                ctype = part.get_content_type()
                try:
                    if ctype == "text/plain":
                        text_parts.append(decode_part(part))
                    elif ctype == "text/html":
                        html_parts.append(decode_part(part))
                except Exception:
                    pass
        else:
            ctype = msg.get_content_type()
            if ctype == "text/plain":
                text_parts.append(decode_part(msg))
            elif ctype == "text/html":
                html_parts.append(decode_part(msg))
    except Exception:
        pass
    text = "\n".join(text_parts).strip()
    html = "\n".join(html_parts).strip()
    if not text and html:
        text = strip_html(html)
    return text, html


def _internaldate(msgdata, tz_local):
    """从 IMAP 响应里解析 INTERNALDATE 作为 Date 头缺失时的兜底。"""
    try:
        if not msgdata or not isinstance(msgdata[0], bytes):
            return None
        m = re.search(rb'INTERNALDATE "([^"]+)"', msgdata[0])
        if not m:
            return None
        s = m.group(1).decode().strip()
        return dt.datetime.strptime(s, "%d-%b-%Y %H:%M:%S %z").astimezone(tz_local)
    except Exception:
        return None


def parse_message(raw, msgdata, tz_local):
    msg = email.message_from_bytes(raw)
    subject = decode_mime(msg.get("Subject"))
    from_ = decode_mime(msg.get("From"))
    to_ = decode_mime(msg.get("To"))
    date_hdr = msg.get("Date")
    when = None
    if date_hdr:
        try:
            when = parsedate_to_datetime(date_hdr)
        except Exception:
            when = None
    if when is not None and when.tzinfo is None:
        when = when.replace(tzinfo=tz_local)
    if when is None:
        when = _internaldate(msgdata, tz_local)
    if when is None:
        when = dt.datetime.now(tz_local)
    text, html = get_body(msg)

    # 未读标记：FLAGS 一般在响应的最后一段，也可能拼在首段
    flags = b""
    for piece in (msgdata or ()):
        if isinstance(piece, bytes) and b"FLAGS" in piece:
            flags = piece
            break
    unread = b"\\Seen" not in flags

    mid = msg.get("Message-ID")
    if mid:
        mid = mid.strip()
        if mid.startswith("<") and mid.endswith(">"):
            mid = mid[1:-1]

    return {"subject": subject, "from_": from_, "to_": to_,
            "date": when, "body": text, "html": html, "unread": unread,
            "message_id": mid or None, "sender_email": extract_email(from_)}


def compute_since(cfg):
    """取邮件的时间起点：默认今日 0 点；配置了 lookback_hours 则取最近 N 小时。"""
    now = dt.datetime.now().astimezone()
    h = cfg.get("lookback_hours")
    if h:
        return now - dt.timedelta(hours=h)
    return now.replace(hour=0, minute=0, second=0, microsecond=0)


def extract_email(addr):
    """从 "名字 <a@b.com>" 里提取纯邮箱地址（小写），提取不到返回空串。"""
    if not addr:
        return ""
    m = re.search(r"<([^<>@\s]+@[^<>@\s]+)>", addr)
    if m:
        return m.group(1).lower()
    m = re.search(r"[\w.+-]+@[\w.-]+", addr)
    return m.group(0).lower() if m else ""


def ensure_message_id(em):
    """保证每封邮件有稳定的 message_id 作为缓存键；没有 Message-ID 头时用内容哈希兜底。"""
    if em.get("message_id"):
        return
    import hashlib
    key = "%s|%s|%s" % (em.get("from_", ""), em.get("subject", ""),
                        em.get("date", "").isoformat())
    em["message_id"] = "fallback-" + hashlib.md5(key.encode("utf-8")).hexdigest()


def load_cache(path):
    """读取分析结果缓存（{emails: {message_id: {analysis, analyzed_at, ...}}}）。"""
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and isinstance(data.get("emails"), dict):
            return data
    except Exception:
        pass
    return {"emails": {}}


def save_cache(path, cache):
    """原子写入缓存文件。"""
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)
        os.replace(tmp, path)
    except Exception as e:
        print("  [警告] 缓存写入失败（不影响本次分析）: %s" % e, file=sys.stderr)


def clean_cache(cache, max_age_days):
    """按时间清理过期缓存条目（定期清理）。max_age_days<=0 或 None 时不清理。
    返回移除的条数。"""
    if not max_age_days or max_age_days <= 0:
        return 0
    cutoff = dt.datetime.now().astimezone() - dt.timedelta(days=max_age_days)
    m = cache.setdefault("emails", {})
    expired = []
    for k, v in m.items():
        a = v.get("analyzed_at")
        if not a:
            continue  # 无时间戳的旧记录不擅自删除
        try:
            if dt.datetime.fromisoformat(a) < cutoff:
                expired.append(k)
        except Exception:
            continue
    for k in expired:
        m.pop(k, None)
    return len(expired)


def is_llm_candidate(em, cfg):
    """--important-only 用的本地预过滤：命中营销发件人或明显低优先级关键词则不调用 API。"""
    sender = em.get("sender_email", "")
    skips = [s.strip().lower() for s in str(cfg.get("skip_senders", "")).split(",") if s.strip()]
    if skips and any(s in sender for s in skips):
        return False
    s = ((em.get("subject") or "") + " " + (em.get("body") or ""))[:800].lower()
    if any(k in s for k in _LOW_KW) and not any(k in s for k in _HIGH_KW):
        return False
    return True


# ============================================================
# 获取邮件
# ============================================================

def imap_login(M, user, password):
    """登录 IMAP；账号/密码含非 ASCII 字符时退回 SASL PLAIN（UTF-8）。"""
    try:
        M.login(user, password)
    except UnicodeEncodeError:
        import base64
        token = base64.b64encode(
            ("\0%s\0%s" % (user, password)).encode("utf-8")
        ).decode("ascii")
        M.authenticate("PLAIN", lambda x: token)


# ============================================================
# 微软 OAuth2（XOAUTH2）支持：Outlook / Office365 学校邮箱专用
# 微软已禁用 Office365 的密码登录 IMAP，学校邮箱必须走 OAuth2：
#   第一次运行时浏览器授权一次 -> 脚本保存 refresh_token -> 之后自动续期。
# ============================================================

def _urlopen_robust(url, data=None, headers=None, timeout=60):
    """带证书兜底的 HTTPS 请求：
    1) 优先用 certifi 的证书库；2) 退回系统默认；
    3) 若本机证书库缺失（macOS python.org 常见），降级为不校验证书并警告。"""
    import ssl

    def attempt(ctx):
        req = urllib.request.Request(url, data=data, headers=headers or {})
        return urllib.request.urlopen(req, timeout=timeout, context=ctx)

    try:
        try:
            import certifi
            return attempt(ssl.create_default_context(cafile=certifi.where()))
        except Exception:
            pass
        return attempt(ssl.create_default_context())
    except urllib.error.URLError as e:
        reason = str(getattr(e, "reason", e))
        if "CERTIFICATE_VERIFY_FAILED" in reason or "certificate verify failed" in reason:
            print("  [警告] 本机 Python 证书库缺失，HTTPS 已降级为不校验证书（建议运行 Python "
                  "安装目录下的 Install Certificates.command 修复）", file=sys.stderr)
            return attempt(ssl._create_unverified_context())
        raise


def _post_form(url, data):
    """POST 表单并返回 JSON（微软 OAuth 接口用）。HTTP 400（如 authorization_pending）
    且响应体是合法 JSON 时也返回该 JSON，由调用方判断。"""
    body = urllib.parse.urlencode(data).encode("utf-8")
    try:
        with _urlopen_robust(url, data=body, headers={
                "Content-Type": "application/x-www-form-urlencoded",
        }, timeout=60) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            detail = json.loads(raw)
            if isinstance(detail, dict):
                return detail
        except Exception:
            pass
        raise RuntimeError("微软 OAuth 请求失败: HTTP %s %s" % (e.code, raw[:300]))


def oauth_device_flow(cfg):
    """设备代码流：打印浏览器网址+代码，等待用户登录授权，返回 (access_token, refresh_token)。"""
    base = "https://login.microsoftonline.com/%s/oauth2/v2.0" % cfg.get("oauth_tenant", "common")
    client_id = cfg.get("oauth_client_id") or OAUTH_DEFAULT_CLIENT_ID
    d = _post_form(base + "/devicecode", {"client_id": client_id, "scope": OAUTH_SCOPE})
    print()
    print("=" * 64)
    print("需要完成一次微软登录（只需这一次，之后自动续期）：")
    print("  1) 用浏览器打开：%s" % d.get("verification_uri", "https://microsoft.com/devicelogin"))
    print("  2) 输入代码：%s" % d.get("user_code", ""))
    print("  3) 用 %s 账号登录并点击“同意”" % cfg.get("imap_user", ""))
    print("=" * 64)
    print()
    sys.stdout.flush()
    interval = max(int(d.get("interval", 5)), 2)
    deadline = dt.datetime.now() + dt.timedelta(seconds=int(d.get("expires_in", 900)))
    while dt.datetime.now() < deadline:
        time.sleep(interval)
        r = _post_form(base + "/token", {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id,
            "device_code": d["device_code"],
        })
        if r.get("access_token"):
            return r["access_token"], r.get("refresh_token", "")
        err = r.get("error", "")
        if err == "authorization_declined":
            raise RuntimeError("你拒绝了授权，已退出。")
        if err == "expired_token":
            raise RuntimeError("验证码已过期，请重新运行 --oauth-login。")
        # authorization_pending -> 继续等待用户完成登录
    raise RuntimeError("登录超时，请重新运行 --oauth-login。")


def oauth_refresh(cfg, refresh_token):
    """用 refresh_token 换新的 access_token。返回 (access_token, refresh_token, expires_in)。"""
    base = "https://login.microsoftonline.com/%s/oauth2/v2.0" % cfg.get("oauth_tenant", "common")
    client_id = cfg.get("oauth_client_id") or OAUTH_DEFAULT_CLIENT_ID
    r = _post_form(base + "/token", {
        "grant_type": "refresh_token",
        "client_id": client_id,
        "refresh_token": refresh_token,
        "scope": OAUTH_SCOPE,
    })
    if not r.get("access_token"):
        raise RuntimeError("刷新令牌失败: %s %s"
                           % (r.get("error", "?"), r.get("error_description", "")))
    return r["access_token"], r.get("refresh_token") or refresh_token, int(r.get("expires_in", 3600))


def load_oauth_tokens(path):
    try:
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
        if isinstance(d, dict) and d.get("refresh_token"):
            return d
    except Exception:
        pass
    return None


def save_oauth_tokens(path, tokens):
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(tokens, f, ensure_ascii=False, indent=1)
        os.replace(tmp, path)
        os.chmod(path, 0o600)
    except Exception as e:
        print("  [警告] OAuth 令牌保存失败（不影响本次连接）: %s" % e, file=sys.stderr)


def get_access_token(cfg, mode="auto"):
    """获取可用的微软 access_token。
    mode: auto=优先用缓存，过期则刷新；refresh=强制刷新；login=强制重新浏览器登录。"""
    path = cfg.get("oauth_token_file") or ".outlook_oauth.json"
    tok = None if mode == "login" else load_oauth_tokens(path)
    if tok and mode != "refresh":
        expires_at = tok.get("expires_at", 0)
        if expires_at - time.time() > 300:
            return tok["access_token"], path
    if tok:
        try:
            at, rt, exp = oauth_refresh(cfg, tok["refresh_token"])
            tok = {"access_token": at, "refresh_token": rt,
                   "expires_at": time.time() + exp, "user": cfg.get("imap_user")}
            save_oauth_tokens(path, tok)
            return at, path
        except Exception as e:
            print("  [OAuth] 刷新令牌失败（%s），需要重新登录一次。" % e, file=sys.stderr)
    at, rt = oauth_device_flow(cfg)
    tok = {"access_token": at, "refresh_token": rt,
           "expires_at": time.time() + 3600, "user": cfg.get("imap_user")}
    save_oauth_tokens(path, tok)
    return at, path


def imap_login_oauth(M, user, access_token):
    """用 XOAUTH2 登录 IMAP（Outlook / Office365 学校邮箱）。"""
    auth = "user=%s\x01auth=Bearer %s\x01\x01" % (user, access_token)
    M.authenticate("XOAUTH2", lambda x: auth)


def fetch_emails(cfg):
    """通过 IMAP 拉取“今日”邮件，返回 (邮件列表, 时间起点)。"""
    since = compute_since(cfg)
    host, port = cfg["imap_host"], cfg["imap_port"]
    socket.setdefaulttimeout(60)

    if cfg["imap_ssl"]:
        M = imaplib.IMAP4_SSL(host, port)
    else:
        M = imaplib.IMAP4(host, port)
        try:
            M.starttls()
        except Exception:
            pass

    try:
        if cfg.get("imap_auth") == "oauth2":
            # Outlook / Office365 学校邮箱：XOAUTH2
            at, _ = get_access_token(cfg, mode="auto")
            try:
                imap_login_oauth(M, cfg["imap_user"], at)
            except Exception as e:
                print("  [OAuth] 令牌验证失败（%s），自动刷新后重试..." % e, file=sys.stderr)
                at, _ = get_access_token(cfg, mode="refresh")
                imap_login_oauth(M, cfg["imap_user"], at)
        else:
            imap_login(M, cfg["imap_user"], cfg["imap_password"])
        status, data = M.select(cfg["imap_folder"], readonly=True)
        if status != "OK":
            raise RuntimeError("无法打开邮箱文件夹 %r: %s" % (cfg["imap_folder"], data))
        status, data = M.search(None, "SINCE", since.strftime("%d-%b-%Y"))
        if status != "OK":
            raise RuntimeError("IMAP 搜索失败: %s" % (data,))
        nums = data[0].split()
        if not nums:
            return [], since
        # 只取最近 max_emails 封候选，再按 Date 头精确过滤
        nums = nums[-cfg["max_emails"]:]
        tz_local = dt.datetime.now().astimezone().tzinfo
        emails = []
        for num in nums:
            try:
                status, msgdata = M.fetch(num, "(RFC822 FLAGS)")
                if status != "OK" or not msgdata or msgdata[0] is None:
                    continue
                em = parse_message(msgdata[0][1], msgdata, tz_local)
                if em and em["date"] >= since:
                    emails.append(em)
            except Exception as e:
                print("  [警告] 解析邮件 %s 失败: %s" % (num, e), file=sys.stderr)
                continue
    finally:
        try:
            M.logout()
        except Exception:
            pass

    emails.sort(key=lambda e: e["date"], reverse=True)
    for i, em in enumerate(emails, 1):
        em["index"] = i
    return emails, since


def demo_emails():
    """内置示例邮件，用于不连邮箱也能看到完整效果（--demo）。"""
    tz = dt.datetime.now().astimezone().tzinfo
    now = dt.datetime.now(tz)
    samples = [
        {
            "subject": "[告警] 生产服务器 web-03 CPU 持续 95%，请立即处理",
            "from_": "监控平台 <alert@monitor.example.com>",
            "to_": "you@example.com",
            "date": now - dt.timedelta(minutes=40),
            "unread": True,
            "body": (
                "告警详情\n"
                "--------\n"
                "主机：web-03（10.0.0.13）\n"
                "指标：CPU 使用率 95%（阈值 80%），持续时间 30 分钟\n"
                "时间：" + now.strftime("%Y-%m-%d %H:%M:%S") + "\n"
                "可能原因：上午 10 点发布的新版本流量增长，疑似存在死循环或连接未释放。\n"
                "建议：登录服务器执行 top 查看占用进程，必要时回滚到上一个稳定版本。\n"
                "本条消息为自动发送，请勿直接回复。"
            ),
        },
        {
            "subject": "【重要】本周五项目评审会改期至下周一 10:00",
            "from_": "张三 <zhangsan@example.com>",
            "to_": "you@example.com",
            "date": now - dt.timedelta(hours=2),
            "unread": True,
            "body": (
                "你好，\n"
                "因产品经理临时出差，原定本周五 14:00 的《新版本需求评审会》调整到下周一 10:00，"
                "地点不变（3 楼大会议室）。\n"
                "请提前把各自负责模块的进度更新到评审文档，周日晚 20:00 前完成。\n"
                "如时间冲突请尽快回复我协调。\n"
                "谢谢！"
            ),
        },
        {
            "subject": "您有一笔退款已到账：¥128.00",
            "from_": "支付平台 <noreply@pay.example.com>",
            "to_": "you@example.com",
            "date": now - dt.timedelta(hours=5),
            "unread": False,
            "body": (
                "尊敬的用户：\n"
                "您订单 20250819xxxx 的退款 ¥128.00 已原路退回，预计 1-2 个工作日到账，"
                "请留意银行/支付账户变动。如有疑问请通过官方客服渠道咨询。"
            ),
        },
        {
            "subject": "【订阅】每周技术周刊 Vol.128",
            "from_": "技术周刊 <weekly@newsletter.example.com>",
            "to_": "you@example.com",
            "date": now - dt.timedelta(hours=9),
            "unread": False,
            "body": (
                "本期看点：\n"
                "1. Python 3.13 性能改进实测\n"
                "2. IMAP 协议实现邮件客户端的最佳实践\n"
                "3. 本周开源项目推荐\n"
                "如不想再收到，请点击底部退订链接。"
            ),
        },
    ]
    for i, em in enumerate(samples, 1):
        em["index"] = i
        em.setdefault("message_id", "demo-%d@example.com" % i)
        em.setdefault("sender_email", extract_email(em["from_"]))
        em.setdefault("mailbox", "示例")
    return samples


# ============================================================
# DeepSeek API 调用（唯一的对外 LLM 调用，不经过 DSH）
# ============================================================

def call_deepseek(cfg, messages, max_tokens=4096, json_mode=True):
    payload = {
        "model": cfg["model"],
        "messages": messages,
        "temperature": 0.3,
        "max_tokens": max_tokens,
        "stream": False,
    }
    if json_mode:
        payload["response_format"] = {"type": "json_object"}
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    url = cfg["api_base"].rstrip("/") + "/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + cfg["api_key"],
    }
    try:
        with _urlopen_robust(url, data=body, headers=headers, timeout=180) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")
        raise RuntimeError("DeepSeek API HTTP %s: %s" % (e.code, detail[:500]))
    except urllib.error.URLError as e:
        raise RuntimeError("DeepSeek API 连接失败: %s" % (e.reason,))
    choice = data["choices"][0]
    usage = data.get("usage", {})
    return choice["message"]["content"], choice.get("finish_reason"), usage


def extract_json(text):
    """从模型输出中提取 JSON 对象（容错：去掉 markdown 代码块围栏）。"""
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("响应中未找到 JSON 对象")
    return json.loads(text[start:end + 1])


def build_email_doc(em, body_max):
    """把一封邮件整理成送给模型看的 JSON 文档。"""
    return {
        "index": em["index"],
        "from": em["from_"],
        "to": em["to_"],
        "date": em["date"].strftime("%Y-%m-%d %H:%M:%S %z"),
        "subject": em["subject"],
        "body": em["body"][:body_max],
    }


def add_usage(a, b):
    for k in ("prompt_tokens", "completion_tokens", "total_tokens"):
        a[k] = a.get(k, 0) + b.get(k, 0)
    return a


# ---- 提示词 ----

BATCH_SYSTEM_PROMPT = (
    "你是一个专业的邮件助理，负责帮用户快速筛选邮箱。"
    "你会收到一批 JSON 格式的邮件（字段：index/from/to/date/subject/body）。"
    "请对每一封邮件给出分析，并只输出一个 JSON 对象，格式如下：\n"
    "{\n"
    '  "results": [\n'
    '    {\n'
    '      "index": <对应输入的 index>,\n'
    '      "importance": <1到5的整数，5=最重要；请有区分度，不要把全部邮件都标成4或5>,\n'
    '      "label": <"紧急"|"重要"|"日常"|"通知"|"订阅"|"垃圾营销"|"其他">,\n'
    '      "meaning": "<用一两句中文说明这封邮件在传达什么>",\n'
    '      "summary": "<用3-5句中文概括邮件要点>",\n'
    '      "action": "<建议动作，如：立即回复/尽快处理/稍后阅读/归档/忽略/删除>",\n'
    '      "sender_trust": "<可信|可疑|钓鱼风险>"\n'
    "    }\n"
    "  ],\n"
    '  "overview": "<对整批邮件的中文概述：整体优先级、最需要注意的1-2件事>"\n'
    "}\n"
    "注意：results 的数量必须与输入邮件数量一致，index 必须一一对应。"
    "请综合发件人、主题、正文判断，警惕可疑发件人与营销邮件。"
)

ONE_SYSTEM_PROMPT = (
    "你是一个专业的邮件助理。请分析下面这封邮件，只输出一个 JSON 对象（不要输出其他内容）：\n"
    '{\n'
    '  "importance": <1到5的整数，5=最重要>,\n'
    '  "label": <"紧急"|"重要"|"日常"|"通知"|"订阅"|"垃圾营销"|"其他">,\n'
    '  "meaning": "<用一两句中文说明这封邮件在传达什么>",\n'
    '  "summary": "<用3-5句中文概括邮件要点>",\n'
    '  "action": "<建议动作>",\n'
    '  "sender_trust": "<可信|可疑|钓鱼风险>"\n'
    "}"
)

OVERVIEW_SYSTEM_PROMPT = (
    "你是用户的邮件助理。下面是用户今日邮件的简要分析列表（JSON，字段：subject/from/label/"
    "importance/meaning）。请用中文写一段 80~150 字的“今日邮件总览”：整体繁忙程度、最需要优先"
    "处理的事、有没有需要警惕的风险（可疑发件人/钓鱼）。只输出 JSON 对象：{\"overview\": \"...\"}"
)


# ============================================================
# 分析（批量 + 逐封兜底 + dry-run 启发式）
# ============================================================

_HIGH_KW = ("告警", "紧急", "urgent", "alert", "截止", "due", "故障", "宕机",
            "评审", "上线", "合同", "发票", "欠费", "中断", "停机")
_LOW_KW = ("优惠", "促销", "折扣", "订阅", "周报", "newsletter", "unsubscribe",
           "推广", "广告", "weekly", "digest", "周刊", "退款", "到账")


def heuristic_analyze(em):
    """dry-run 模式用的简单关键词分析（不调用 API）。"""
    s = ((em.get("subject") or "") + " " + (em.get("body") or ""))[:500].lower()
    imp = 3
    if any(k in s for k in _HIGH_KW):
        imp = 5
    if any(k in s for k in _LOW_KW):
        imp = min(imp, 2)
    label = {1: "订阅", 2: "通知", 3: "日常", 4: "重要", 5: "紧急"}[imp]
    return {
        "importance": imp,
        "label": label,
        "meaning": em.get("subject") or "(无主题)",
        "summary": (em.get("body") or "")[:80].replace("\n", " "),
        "action": "立即处理" if imp >= 4 else "稍后阅读",
        "sender_trust": "可信",
    }


def analyze_one(cfg, em, dry_run, usage_tot=None):
    """逐封分析（dry_run 时用启发式，不消耗 token）。"""
    if dry_run:
        return heuristic_analyze(em)
    doc = build_email_doc(em, cfg["body_max"])
    content, finish, usage = call_deepseek(
        cfg,
        [{"role": "system", "content": ONE_SYSTEM_PROMPT},
         {"role": "user", "content": "邮件(JSON)：\n" + json.dumps(doc, ensure_ascii=False)}],
        max_tokens=1500,
    )
    if usage_tot is not None:
        add_usage(usage_tot, usage)
    try:
        return extract_json(content)
    except Exception as e:
        return {"importance": 3, "label": "其他",
                "meaning": "分析失败: %s" % e,
                "summary": content[:300], "action": "手动查看", "sender_trust": "未知"}


def analyze_batch(cfg, emails, dry_run):
    """批量分析（默认方式，省 token）。返回 (index -> 分析结果, 用量统计)。"""
    usage_tot = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
    results = {}
    if dry_run:
        for em in emails:
            results[em["index"]] = heuristic_analyze(em)
        return results, usage_tot

    docs = [build_email_doc(em, cfg["body_max"]) for em in emails]
    for start in range(0, len(docs), cfg["batch_size"]):
        chunk = docs[start:start + cfg["batch_size"]]
        content, finish, usage = call_deepseek(
            cfg,
            [{"role": "system", "content": BATCH_SYSTEM_PROMPT},
             {"role": "user",
              "content": "邮件列表(JSON)：\n" + json.dumps(chunk, ensure_ascii=False)}],
        )
        add_usage(usage_tot, usage)
        try:
            parsed = extract_json(content).get("results", [])
        except Exception as e:
            # 批量解析失败 -> 该批逐封重试
            print("  [警告] 批量响应解析失败(%s)，该批改为逐封分析" % e, file=sys.stderr)
            for em in chunk:
                try:
                    results[em["index"]] = analyze_one(cfg, em, False, usage_tot)
                except Exception as e2:
                    results[em["index"]] = {"importance": 3, "label": "其他",
                                            "meaning": "分析失败: %s" % e2,
                                            "summary": "", "action": "手动查看",
                                            "sender_trust": "未知"}
            continue
        if finish == "length":
            print("  [警告] 批量响应被截断，该批改为逐封分析", file=sys.stderr)
            for em in chunk:
                try:
                    results[em["index"]] = analyze_one(cfg, em, False, usage_tot)
                except Exception as e2:
                    results[em["index"]] = {"importance": 3, "label": "其他",
                                            "meaning": "分析失败: %s" % e2,
                                            "summary": "", "action": "手动查看",
                                            "sender_trust": "未知"}
            continue
        for r in parsed:
            try:
                results[int(r.get("index", -1))] = r
            except Exception:
                continue
    return results, usage_tot


def make_overview(cfg, emails, analyses, dry_run):
    """生成今日邮件总览。返回 (总览文本, 用量统计)。"""
    brief = [
        {"subject": em["subject"], "from": em["from_"],
         "label": a.get("label"), "importance": a.get("importance"),
         "meaning": a.get("meaning")}
        for em, a in zip(emails, analyses)
    ]
    if dry_run:
        return "（dry-run 模式未调用 API，无总览）", \
            {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
    content, finish, usage = call_deepseek(
        cfg,
        [{"role": "system", "content": OVERVIEW_SYSTEM_PROMPT},
         {"role": "user", "content": json.dumps(brief, ensure_ascii=False)}],
        max_tokens=1000,
    )
    try:
        return extract_json(content).get("overview", content), usage
    except Exception:
        return content, usage


# ============================================================
# 输出
# ============================================================

def print_report(cfg, emails, analyses, overview_text, usage, since, args):
    use_color = args.color and sys.stdout.isatty()
    now = dt.datetime.now().astimezone()
    width = 64
    sep = col("─" * width, "dim", use_color)

    print()
    print(col("═" * width, "cyan", use_color))
    print(col("  今日邮件智能分析报告  %s" % now.strftime("%Y-%m-%d %H:%M"), "bold", use_color))
    print(col("═" * width, "cyan", use_color))

    n_important = sum(1 for a in analyses
                      if int(a.get("importance", 3)) >= cfg["threshold"])
    n_cached = sum(1 for em in emails if "cached_analysis" in em)
    n_rule = sum(1 for em in emails if em.get("rule_skipped"))
    print("统计：共 %d 封 ｜ 新分析 %d ｜ 缓存复用 %d ｜ 规则过滤 %d ｜ 重要(≥%d) %d 封 ｜ 范围：%s 起"
          % (len(emails), len(emails) - n_cached - n_rule, n_cached, n_rule,
             cfg["threshold"], n_important, since.strftime("%m-%d %H:%M")))

    if overview_text:
        print()
        print(col("【今日总览】", "bold", use_color))
        print(overview_text)

    for em, a in zip(emails, analyses):
        imp = int(a.get("importance", 3))
        label = a.get("label", "其他")
        trust = a.get("sender_trust", "")
        important = imp >= cfg["threshold"] or "紧急" in str(label)
        stars = col("★" * imp + "☆" * (5 - imp),
                    "red" if imp >= 4 else ("yellow" if imp == 3 else "green"),
                    use_color)
        tag = col("[重要] ", "red", use_color) if important else ""
        unread = col("未读", "magenta", use_color) if em["unread"] else ""
        if "cached_analysis" in em:
            src_tag = col("（缓存，未耗 token）", "cyan", use_color)
        elif em.get("rule_skipped"):
            src_tag = col("（规则过滤，未耗 token）", "dim", use_color)
        else:
            src_tag = ""
        print()
        print(sep)
        print("%s%d. %s  %s ｜ 发件人可信度：%s  %s  %s"
              % (tag, em["index"], stars, label, trust, unread, src_tag))
        mb = "邮箱：%s  ｜  " % em.get("mailbox", "") if em.get("mailbox") else ""
        print(col("   %s时间：%s  ｜  发件人：%s"
                  % (mb, em["date"].strftime("%Y-%m-%d %H:%M"), em["from_"]), "dim", use_color))
        print("   主题：%s" % (em["subject"] or "(无主题)"))
        print("   含义：%s" % a.get("meaning", "-"))
        print("   摘要：%s" % a.get("summary", "-"))
        print("   建议：%s" % a.get("action", "-"))
        if (important or args.bodies) and em["body"]:
            print(col("   ── 原文 ──", "cyan", use_color))
            body = em["body"]
            if not args.full_body and len(body) > cfg["print_body_cap"]:
                body = body[:cfg["print_body_cap"]] + \
                    "\n   …(原文过长已截断，可用 --full-body 查看全文，共 %d 字)" % len(em["body"])
            for line in body.splitlines():
                print("   " + line)

    print()
    print(sep)
    if args.dry_run:
        print("dry-run 模式：未调用 DeepSeek API，未消耗任何 token。")
    else:
        print("已消耗 DeepSeek token：%s 输入 + %s 输出（共 %s）"
              % (usage.get("prompt_tokens", 0), usage.get("completion_tokens", 0),
                 usage.get("total_tokens", 0)))
        cost = (usage.get("prompt_tokens", 0) * float(cfg.get("price_input", 1.5))
                + usage.get("completion_tokens", 0) * float(cfg.get("price_output", 4.5))) / 1e6
        print("预估费用：约 ¥%.4f（按 deepseek V4-Flash 空闲时段价：输入 ¥%.2f/M、输出 ¥%.2f/M；"
              "高峰时段 9:00-14:00 约翻倍）" % (cost, float(cfg.get("price_input", 1.5)),
                                             float(cfg.get("price_output", 4.5))))
    print("本脚本不经过 DSH，不调用 DS Agent，仅使用 DeepSeek 官方 API。")


# ============================================================
# 桌面报告 + macOS 弹窗（--desktop-report / --popup）
# ============================================================

_HTML_CSS = """
body { font-family: -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif;
       max-width: 880px; margin: 24px auto; padding: 0 16px; color: #222; background: #fafafa; }
h1 { font-size: 22px; }
.stat { color: #555; font-size: 14px; }
.overview { background: #eef4ff; border: 1px solid #c9d8f0; border-radius: 8px; padding: 12px 16px; margin: 16px 0; }
.mail { background: #fff; border: 1px solid #e2e2e2; border-radius: 10px; padding: 14px 18px; margin: 14px 0; }
.mail.important { border-color: #e0a0a0; box-shadow: 0 0 0 1px #f2c9c9; }
.head { font-size: 15px; font-weight: 600; }
.idx { color: #999; margin-right: 6px; font-weight: 400; }
.stars { color: #e8a33d; }
.label { display: inline-block; padding: 1px 10px; border-radius: 10px; font-size: 12px; color: #fff; margin: 0 8px; }
.label.imp { background: #d9534f; } .label.mid { background: #f0ad4e; } .label.low { background: #5cb85c; }
.trust { color: #777; font-size: 13px; font-weight: 400; }
.src { color: #999; font-size: 12px; font-weight: 400; margin-left: 6px; }
.unread { color: #c0392b; font-size: 12px; font-weight: 400; margin-left: 8px; }
.meta, .subject { color: #444; margin-top: 6px; font-size: 13.5px; }
.line { color: #333; margin-top: 4px; font-size: 13.5px; }
.body { margin-top: 8px; background: #f7f7f7; border-radius: 6px; padding: 8px 12px; }
.body pre { white-space: pre-wrap; word-break: break-word; font-family: inherit; font-size: 12.5px; color: #333; margin: 4px 0 0 0; }
.foot { color: #999; font-size: 12px; margin: 20px 0 40px; }
"""


def write_html_report(cfg, emails, analyses, overview_text, usage, since, args):
    """把完整报告写成自包含 HTML 存到桌面，返回文件路径。"""
    from html import escape
    desktop = os.path.expanduser(cfg.get("desktop_dir") or "~/Desktop")
    os.makedirs(desktop, exist_ok=True)
    now = dt.datetime.now().astimezone()
    path = os.path.join(desktop, "今日邮件分析_%s.html" % now.strftime("%Y-%m-%d"))

    n_important = sum(1 for a in analyses if int(a.get("importance", 3)) >= cfg["threshold"])
    n_cached = sum(1 for em in emails if "cached_analysis" in em)
    n_rule = sum(1 for em in emails if em.get("rule_skipped"))
    n_new = len(emails) - n_cached - n_rule

    rows = []
    for em, a in zip(emails, analyses):
        imp = int(a.get("importance", 3))
        important = imp >= cfg["threshold"] or "紧急" in str(a.get("label", ""))
        if "cached_analysis" in em:
            src = "缓存"
        elif em.get("rule_skipped"):
            src = "规则过滤"
        else:
            src = "LLM"
        body_html = ""
        if (important or args.bodies) and em["body"]:
            body_html = "<div class='body'><b>原文：</b><pre>%s</pre></div>" % escape(em["body"])
        rows.append(
            "<div class='mail%s'>"
            "<div class='head'><span class='idx'>%d</span><span class='stars'>%s</span>"
            "<span class='label %s'>%s</span><span class='trust'>发件人可信度：%s</span>"
            "<span class='src'>（%s）</span>%s</div>"
            "<div class='meta'>%s时间：%s ｜ 发件人：%s</div>"
            "<div class='subject'>主题：%s</div>"
            "<div class='line'>含义：%s</div>"
            "<div class='line'>摘要：%s</div>"
            "<div class='line'>建议：%s</div>"
            "%s</div>"
            % (" important" if important else "", em["index"],
               "★" * imp + "☆" * (5 - imp),
               "imp" if imp >= 4 else ("mid" if imp == 3 else "low"),
               escape(str(a.get("label", ""))),
               escape(str(a.get("sender_trust", ""))), src,
               "  <span class='unread'>未读</span>" if em["unread"] else "",
               "邮箱：%s ｜ " % escape(em.get("mailbox", "")) if em.get("mailbox") else "",
               em["date"].strftime("%Y-%m-%d %H:%M"), escape(em["from_"]),
               escape(em["subject"] or "(无主题)"),
               escape(str(a.get("meaning", "-"))),
               escape(str(a.get("summary", "-"))),
               escape(str(a.get("action", "-"))),
               body_html)
        )

    overview_html = ("<div class='overview'><b>今日总览：</b>%s</div>" % escape(overview_text)) \
        if overview_text else ""
    if args.dry_run:
        foot = "dry-run 模式：未调用 DeepSeek API，未消耗任何 token。"
    else:
        foot = ("已消耗 DeepSeek token：%s 输入 + %s 输出（共 %s）｜ 预估费用：约 ¥%.4f"
                % (usage.get("prompt_tokens", 0), usage.get("completion_tokens", 0),
                   usage.get("total_tokens", 0),
                   (usage.get("prompt_tokens", 0) * float(cfg.get("price_input", 1.5))
                    + usage.get("completion_tokens", 0) * float(cfg.get("price_output", 4.5))) / 1e6))
    foot += " ｜ 本脚本不经过 DSH，仅使用 DeepSeek 官方 API。"

    html = ("<!DOCTYPE html>\n<html lang='zh'><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width, initial-scale=1'>"
            "<title>今日邮件分析 %s</title><style>%s</style></head><body>"
            "<h1>今日邮件智能分析报告 %s</h1>"
            "<div class='stat'>共 %d 封 ｜ 重要(≥%d) %d 封 ｜ 新分析 %d ｜ 缓存 %d ｜ 规则过滤 %d ｜ 范围：%s 起</div>"
            "%s%s"
            "<div class='foot'>%s</div></body></html>") % (
        now.strftime("%Y-%m-%d"), _HTML_CSS, now.strftime("%Y-%m-%d %H:%M"),
        len(emails), cfg["threshold"], n_important, n_new, n_cached, n_rule,
        since.strftime("%m-%d %H:%M"), overview_html, "".join(rows), foot)

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def macos_notify(title, message):
    """macOS 系统通知（右上角横幅，非阻塞；失败静默）。"""
    import subprocess
    title = title.replace("\\", "\\\\").replace('"', '\\"')
    message = message.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")[:300]
    script = 'display notification "%s" with title "%s" sound name "default"' % (message, title)
    try:
        subprocess.run(["osascript", "-e", script], timeout=30, capture_output=True)
    except Exception:
        pass


def open_file(path):
    """用系统默认程序打开文件（macOS: open / Windows: start / Linux: xdg-open）。"""
    import subprocess
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif sys.platform.startswith("win"):
            subprocess.Popen(["cmd", "/c", "start", "", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ============================================================
# 命令行入口
# ============================================================

def build_cfg(args):
    cfg = dict(DEFAULTS)
    env_map = {
        "api_key": ("DEEPSEEK_API_KEY", str),
        "api_base": ("DEEPSEEK_BASE_URL", str),
        "model": ("DEEPSEEK_MODEL", str),
        "imap_host": ("IMAP_HOST", str),
        "imap_port": ("IMAP_PORT", int),
        "imap_user": ("IMAP_USER", str),
        "imap_password": ("IMAP_PASSWORD", str),
        "imap_ssl": ("IMAP_SSL", strtobool),
        "imap_folder": ("IMAP_FOLDER", str),
        "imap_auth": ("IMAP_AUTH", str),
        "oauth_client_id": ("OAUTH_CLIENT_ID", str),
        "oauth_tenant": ("OAUTH_TENANT", str),
        "oauth_token_file": ("OAUTH_TOKEN_FILE", str),
        "lookback_hours": ("LOOKBACK_HOURS", int_or_none),
        "max_emails": ("MAX_EMAILS", int),
        "threshold": ("IMPORTANT_THRESHOLD", int),
        "body_max": ("BODY_MAX", int),
        "batch_size": ("BATCH_SIZE", int),
        "cache_file": ("CACHE_FILE", str),
        "cache_max_age_days": ("CACHE_MAX_AGE_DAYS", int),
        "skip_senders": ("SKIP_SENDERS", str),
        "price_input": ("DEEPSEEK_PRICE_INPUT", float),
        "price_output": ("DEEPSEEK_PRICE_OUTPUT", float),
        "desktop_dir": ("DESKTOP_DIR", str),
    }
    for k, (envk, cast) in env_map.items():
        v = os.environ.get(envk)
        if v is not None and v != "":
            try:
                cfg[k] = cast(v)
            except (ValueError, TypeError):
                pass
    # 命令行参数覆盖（优先级最高）
    for k in ("imap_host", "imap_port", "imap_user", "imap_password",
              "imap_folder", "api_key", "model", "api_base"):
        v = getattr(args, k, None)
        if v:
            cfg[k] = v
    if args.hours is not None:
        cfg["lookback_hours"] = args.hours
    if args.max_emails:
        cfg["max_emails"] = args.max_emails
    if args.threshold:
        cfg["threshold"] = args.threshold
    if args.body_max:
        cfg["body_max"] = args.body_max
    if args.skip_senders:
        cfg["skip_senders"] = args.skip_senders
    if args.cache_file:
        cfg["cache_file"] = args.cache_file
    if args.cache_max_age is not None:
        cfg["cache_max_age_days"] = args.cache_max_age
    if args.imap_auth:
        cfg["imap_auth"] = args.imap_auth
    if args.oauth_token_file:
        cfg["oauth_token_file"] = args.oauth_token_file
    if args.ssl is not None:
        cfg["imap_ssl"] = args.ssl
    return cfg


def collect_mailboxes(cfg):
    """收集所有要读取的邮箱配置。
    默认邮箱用 .env 主配置（IMAP_HOST/IMAP_USER/...）；
    可在 .env 里用 MAILBOX2_NAME、MAILBOX2_IMAP_HOST、MAILBOX2_IMAP_USER、
    MAILBOX2_IMAP_PASSWORD、MAILBOX2_IMAP_AUTH ... 前缀追加第 2、3... 个邮箱，
    一次运行会合并分析所有邮箱。返回带 mailbox_name 的配置列表。"""
    def cast(k, v):
        if k == "imap_port":
            try:
                return int(v)
            except ValueError:
                return None
        if k == "imap_ssl":
            return strtobool(v)
        return v

    boxes = [dict(cfg)]
    boxes[0]["mailbox_name"] = os.environ.get("MAILBOX_NAME") or cfg["imap_host"]
    i = 2
    while True:
        prefix = "MAILBOX%d_" % i
        name = os.environ.get(prefix + "NAME")
        if not name:
            break
        box = dict(cfg)
        box["mailbox_name"] = name
        for k, envk in {
            "imap_host": prefix + "IMAP_HOST",
            "imap_port": prefix + "IMAP_PORT",
            "imap_user": prefix + "IMAP_USER",
            "imap_password": prefix + "IMAP_PASSWORD",
            "imap_auth": prefix + "IMAP_AUTH",
            "imap_folder": prefix + "IMAP_FOLDER",
            "imap_ssl": prefix + "IMAP_SSL",
            "oauth_token_file": prefix + "OAUTH_TOKEN_FILE",
        }.items():
            v = os.environ.get(envk)
            if v is None or v == "":
                continue
            cv = cast(k, v)
            if cv is not None:
                box[k] = cv
        boxes.append(box)
        i += 1
    return boxes


def parse_args():
    p = argparse.ArgumentParser(
        description="通过 DeepSeek API 分析今日邮件（不经过 DSH，不消耗 DSH/Agent token）",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--host", dest="imap_host", default=None,
                   help="IMAP 服务器（默认 %s）" % DEFAULTS["imap_host"])
    p.add_argument("--port", dest="imap_port", type=int, default=None,
                   help="IMAP 端口（默认 %s）" % DEFAULTS["imap_port"])
    p.add_argument("-u", "--user", dest="imap_user", default=None, help="IMAP 用户名/邮箱")
    p.add_argument("-p", "--password", dest="imap_password", default=None,
                   help="IMAP 授权码（QQ/163/Gmail 用授权码而非登录密码）")
    p.add_argument("--folder", dest="imap_folder", default=None, help="邮箱文件夹（默认 INBOX）")
    p.add_argument("--no-ssl", dest="ssl", action="store_false", help="不使用 SSL（端口 143）")
    p.add_argument("--hours", type=float, default=None,
                   help="最近 N 小时（默认：今日 0 点起）")
    p.add_argument("--max-emails", type=int, default=None, help="最多处理邮件数")
    p.add_argument("--threshold", type=int, default=None,
                   help="重要度阈值，>= 该值输出原文（默认 4）")
    p.add_argument("--body-max", type=int, default=None,
                   help="送入 API 的单封邮件正文上限（字符，默认 6000，调小更省 token）")
    p.add_argument("--api-key", dest="api_key", default=None, help="DeepSeek API Key")
    p.add_argument("--model", default=None, help="DeepSeek 模型（默认 deepseek-chat）")
    p.add_argument("--api-base", dest="api_base", default=None, help="DeepSeek API 地址")
    p.add_argument("--env", default=DEFAULT_ENV_FILE, help=".env 文件路径")
    p.add_argument("--json", action="store_true", help="输出原始 JSON 结果")
    p.add_argument("--no-batch", action="store_true", help="逐封分析（更稳但更耗 token）")
    p.add_argument("--no-overview", action="store_true", help="不生成今日总览（省一次调用）")
    p.add_argument("--full-body", action="store_true", help="重要邮件原文不截断")
    p.add_argument("--bodies", action="store_true", help="所有邮件都打印原文")
    p.add_argument("--demo", action="store_true", help="使用内置示例邮件（不连邮箱）")
    p.add_argument("--dry-run", action="store_true",
                   help="不调用 DeepSeek API（用本地关键词规则代替）")
    p.add_argument("--no-color", dest="color", action="store_false", help="关闭颜色输出")
    p.add_argument("--unread-only", action="store_true", help="只看未读邮件")
    p.add_argument("--important-only", action="store_true",
                   help="规则预过滤：订阅/营销类邮件不调用 API（用本地规则判定）")
    p.add_argument("--no-cache", action="store_true",
                   help="关闭结果缓存（同一封邮件每次都重新分析）")
    p.add_argument("--cache-file", default=None,
                   help="缓存文件路径（默认 .mail_analyzer_state.json）")
    p.add_argument("--cache-max-age", type=int, default=None,
                   help="缓存保留天数（默认 90；每次运行自动清理更旧的记录）")
    p.add_argument("--cache-clean", action="store_true",
                   help="手动清理缓存：移除超过 --cache-max-age 天的旧记录")
    p.add_argument("--cache-clear", action="store_true",
                   help="彻底清空缓存文件")
    p.add_argument("--desktop-report", action="store_true",
                   help="把报告写成 HTML 存到桌面（默认 ~/Desktop/今日邮件分析_日期.html）")
    p.add_argument("--popup", action="store_true",
                   help="macOS 弹窗：系统通知 + 有重要邮件时自动在浏览器打开报告（隐含 --desktop-report）")
    p.add_argument("--open-report", action="store_true",
                   help="配合 --popup：无论有无重要邮件都自动打开报告")
    p.add_argument("--skip-senders", default=None,
                   help="预过滤时跳过这些发件人关键字（逗号分隔，默认 newsletter,weekly,...）")
    p.add_argument("--auth", dest="imap_auth", default=None, choices=["password", "oauth2"],
                   help="邮箱登录方式：password=账号密码/授权码（默认）；"
                        "oauth2=Outlook/Office365 学校邮箱（浏览器授权一次）")
    p.add_argument("--oauth-login", action="store_true",
                   help="手动执行一次微软 OAuth 浏览器授权并保存令牌（之后运行无需再登录）")
    p.add_argument("--oauth-reset", action="store_true",
                   help="删除已保存的 OAuth 令牌（下次运行需重新授权）")
    p.add_argument("--oauth-token-file", dest="oauth_token_file", default=None,
                   help="OAuth 令牌文件路径（默认 .outlook_oauth.json）")
    return p.parse_args()


def main():
    args = parse_args()
    if args.popup:
        args.desktop_report = True  # 弹窗需要报告文件
    load_dotenv(args.env)
    cfg = build_cfg(args)
    # --json 时：进度/提示信息走 stderr，stdout 只输出纯 JSON
    progress = sys.stderr if args.json else sys.stdout

    # ---- OAuth 辅助命令（不需要 API Key）----
    if args.oauth_reset:
        path = cfg.get("oauth_token_file") or ".outlook_oauth.json"
        if os.path.exists(path):
            os.remove(path)
            print("已删除 OAuth 令牌文件：%s" % path)
        else:
            print("没有找到 OAuth 令牌文件：%s" % path)
        return 0
    if args.oauth_login:
        if not cfg["imap_user"] or cfg["imap_user"].startswith("you@"):
            print("[错误] 请先在 .env 配置 IMAP_USER（你的邮箱地址）", file=sys.stderr)
            return 2
        at, path = get_access_token(cfg, mode="login")
        print("OAuth 登录成功！令牌已保存到 %s，之后直接运行脚本即可读邮件。" % path)
        return 0

    # ---- 缓存维护命令（不需要 API Key）----
    if args.cache_clear:
        path = cfg.get("cache_file") or ".mail_analyzer_state.json"
        if os.path.exists(path):
            os.remove(path)
            print("已清空缓存文件：%s" % path)
        else:
            print("没有找到缓存文件：%s" % path)
        return 0
    if args.cache_clean:
        path = cfg.get("cache_file") or ".mail_analyzer_state.json"
        days = cfg.get("cache_max_age_days")
        if args.cache_max_age is not None:
            days = args.cache_max_age
        cache = load_cache(path)
        removed = clean_cache(cache, days)
        total = len(cache.get("emails", {}))
        if os.path.exists(path):
            save_cache(path, cache)
        print("缓存清理完成：移除了 %d 条超过 %d 天的记录（剩余 %d 条）"
              % (removed, days, total))
        return 0

    _key_ok = (bool(cfg["api_key"]) and cfg["api_key"].startswith("sk-")
               and "在此填入" not in cfg["api_key"] and "你的" not in cfg["api_key"])
    if not args.demo and not _key_ok and not args.dry_run:
        print("[错误] DeepSeek API Key 未配置（可在脚本顶部 CONFIG 区、.env 或 --api-key 设置）",
              file=sys.stderr)
        return 2

    try:
        if args.demo:
            emails = demo_emails()
            since = compute_since(cfg)
            print("（demo 模式：使用内置示例邮件，未连接邮箱）", file=progress)
        else:
            since = compute_since(cfg)
            emails = []
            for box in collect_mailboxes(cfg):
                print("正在连接邮箱 %s（%s）..."
                      % (box.get("mailbox_name", box["imap_host"]), box["imap_host"]),
                      file=progress)
                box_emails, _ = fetch_emails(box)
                for em in box_emails:
                    em["mailbox"] = box.get("mailbox_name", box["imap_host"])
                emails.extend(box_emails)
            if emails:
                emails.sort(key=lambda e: e["date"], reverse=True)
                for i, em in enumerate(emails, 1):
                    em["index"] = i

        if not emails:
            print("自 %s 起没有符合条件的邮件。"
                  % since.strftime("%Y-%m-%d %H:%M"), file=progress)
            if args.popup:
                macos_notify("今日邮件分析", "今日无新邮件（0 封）")
            return 0

        if args.unread_only:
            before = len(emails)
            emails = [em for em in emails if em["unread"]]
            print("--unread-only：保留 %d 封未读邮件（过滤掉 %d 封已读）"
                  % (len(emails), before - len(emails)), file=progress)
            if not emails:
                print("没有未读邮件。", file=progress)
                if args.popup:
                    macos_notify("今日邮件分析", "今日无未读邮件")
                return 0

        for em in emails:
            ensure_message_id(em)

        # ---- 缓存：同一封邮件（Message-ID）只分析一次 ----
        use_cache = not args.no_cache
        cache = load_cache(cfg["cache_file"]) if use_cache else {"emails": {}}
        cached_map = cache.setdefault("emails", {})
        # 定期清理：每次运行时自动移除超过 cache_max_age_days 天的旧记录
        if use_cache:
            removed = clean_cache(cache, cfg.get("cache_max_age_days"))
            if removed:
                print("缓存清理：自动移除了 %d 条超过 %d 天的旧记录"
                      % (removed, cfg.get("cache_max_age_days")), file=progress)
                save_cache(cfg["cache_file"], cache)  # 立即写回，保证清理持久化
        for em in emails:
            rec = cached_map.get(em["message_id"])
            if rec and rec.get("analysis"):
                em["cached_analysis"] = rec["analysis"]
        new_emails = [em for em in emails if "cached_analysis" not in em]

        print("共获取 %d 封邮件（新邮件 %d 封 / 缓存复用 %d 封）"
              % (len(emails), len(new_emails), len(emails) - len(new_emails)), file=progress)

        # ---- 规则预过滤：明显低优先级的邮件不调用 API ----
        to_analyze = list(new_emails)
        if args.important_only:
            cand, skipped = [], []
            for em in new_emails:
                (cand if is_llm_candidate(em, cfg) else skipped).append(em)
            to_analyze = cand
            for em in skipped:
                em["rule_skipped"] = True
            print("--important-only：规则预过滤 %d 封，不调用 API（省 token）" % len(skipped), file=progress)

        if args.dry_run:
            print("（dry-run：跳过 DeepSeek API 调用，不消耗 token）", file=progress)

        # ---- 分析（只分析真正需要的新邮件）----
        usage = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
        results = {}
        if to_analyze:
            if args.no_batch:
                for em in to_analyze:
                    try:
                        results[em["index"]] = analyze_one(cfg, em, args.dry_run, usage)
                    except Exception as e:
                        results[em["index"]] = {"importance": 3, "label": "其他",
                                                "meaning": "分析失败: %s" % e,
                                                "summary": "", "action": "手动查看",
                                                "sender_trust": "未知"}
            else:
                results, usage = analyze_batch(cfg, to_analyze, args.dry_run)

            # 写回缓存（dry-run 只在 demo 模式写，避免把本地规则结果当成真实分析）
            if use_cache and (not args.dry_run or args.demo):
                now_iso = dt.datetime.now().astimezone().isoformat()
                for em in to_analyze:
                    a = results.get(em["index"])
                    if a:
                        cached_map[em["message_id"]] = {
                            "analyzed_at": now_iso, "analysis": a,
                            "subject": em["subject"], "from": em["from_"],
                        }
                if len(cached_map) > 3000:  # 只保留最近 3000 封
                    for k in sorted(cached_map, key=lambda k: cached_map[k].get("analyzed_at", ""))[
                            :len(cached_map) - 3000]:
                        cached_map.pop(k, None)
                save_cache(cfg["cache_file"], cache)

        analyses = []
        for em in emails:
            if "cached_analysis" in em:
                analyses.append(em["cached_analysis"])
            elif em.get("rule_skipped"):
                analyses.append(heuristic_analyze(em))
            else:
                analyses.append(results.get(em["index"], {}))

        overview_text = ""
        if not args.no_overview:
            if to_analyze or not use_cache:
                overview_text, u2 = make_overview(cfg, emails, analyses, args.dry_run)
                usage = add_usage(usage, u2)
            else:
                overview_text = ("今日没有新邮件需要分析（全部命中缓存），"
                                 "已自动跳过总览调用，未消耗 token。")

        if args.json:
            emails_out = []
            for em in emails:
                d = build_email_doc(em, cfg["print_body_cap"])
                d["unread"] = em["unread"]
                d["cached"] = "cached_analysis" in em
                d["rule_skipped"] = em.get("rule_skipped", False)
                d["mailbox"] = em.get("mailbox", "")
                emails_out.append(d)
            out = {
                "since": since.isoformat(),
                "emails": emails_out,
                "analyses": analyses,
                "overview": overview_text,
                "usage": usage,
            }
            print(json.dumps(out, ensure_ascii=False, indent=2))
        else:
            print_report(cfg, emails, analyses, overview_text, usage, since, args)

        # ---- 桌面报告 + 弹窗（--desktop-report / --popup）----
        if args.popup or args.desktop_report:
            if not emails:
                if args.popup:
                    macos_notify("今日邮件分析", "今日无新邮件（0 封）")
            else:
                report_path = write_html_report(cfg, emails, analyses,
                                                overview_text, usage, since, args)
                print("报告已保存：%s" % report_path, file=progress)
                if args.popup:
                    n_imp = sum(1 for a in analyses
                                if int(a.get("importance", 3)) >= cfg["threshold"])
                    top = ""
                    best = max(zip(emails, analyses),
                               key=lambda t: int(t[1].get("importance", 3)))
                    if best:
                        top = "最需关注：%s" % (best[0]["subject"] or "(无主题)")[:40]
                    macos_notify("今日邮件分析：共 %d 封 · 重要 %d 封"
                                 % (len(emails), n_imp), top or "今日无重要邮件")
                    if n_imp > 0 or args.open_report:
                        open_file(report_path)
        return 0

    except RuntimeError as e:
        print("[错误] %s" % e, file=sys.stderr)
        print("提示：检查邮箱服务器/授权码是否正确；DeepSeek Key 是否有效、账户是否有余额。",
              file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\n已中断。", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
