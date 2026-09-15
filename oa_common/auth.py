# -*- coding: utf-8 -*-
"""OA 自动登录模块：凭据存 Windows 凭据管理器(keyring)，登录用 Playwright 真实点击。

用法：
  from oa_auth import get_credentials, login
  acc, pwd = get_credentials()      # 首次会弹小窗口让你输入一次
  base = login(page, acc, pwd)      # 返回登录成功的站点 base url
"""
import json
import time
from pathlib import Path

import keyring

import sys as _sys
_sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from oa_common import paths
CONFIG = paths.STATE / "oa_config.json"
SERVICE = "qwenwork-oa-tus-sound"
HOSTS = ["http://oa.tus-sound.com", "http://101.200.140.240"]
LOGIN_PATH = "/login/Login.jsp?logintype=1"


def _load_cfg():
    if CONFIG.exists():
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    return {}


def _save_cfg(d):
    CONFIG.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


def get_credentials():
    """返回 (账号, 密码)。密码从凭据管理器取；没有则弹窗输入一次并保存。"""
    cfg = _load_cfg()
    acc = cfg.get("account")
    if acc:
        pwd = keyring.get_password(SERVICE, acc)
        if pwd:
            return acc, pwd
    # 首次：弹窗输入
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.title("OA 凭据首次设置（仅这一次）")
    root.geometry("360x170")
    root.attributes("-topmost", True)
    ttk.Label(root, text="请输入 OA 登录账号和密码，将加密存入 Windows 凭据管理器：").pack(padx=12, pady=(12, 4))
    e_acc = ttk.Entry(root, width=40)
    e_pwd = ttk.Entry(root, width=40, show="*")
    e_acc.pack(padx=12, pady=4)
    e_pwd.pack(padx=12, pady=4)
    result = {}

    def ok():
        result["acc"] = e_acc.get().strip()
        result["pwd"] = e_pwd.get()
        root.destroy()

    ttk.Button(root, text="保存", command=ok).pack(pady=8)
    root.mainloop()

    acc, pwd = result.get("acc"), result.get("pwd")
    if not acc or not pwd:
        raise RuntimeError("未输入账号密码")
    keyring.set_password(SERVICE, acc, pwd)
    cfg["account"] = acc
    _save_cfg(cfg)
    return acc, pwd


def login(page, account=None, password=None, verbose=True):
    """在给定 Playwright Page 上执行登录，返回登录成功的站点 base。"""
    if not account or not password:
        account, password = get_credentials()

    last_err = None
    for base in HOSTS:
        try:
            page.goto(base + LOGIN_PATH, wait_until="domcontentloaded", timeout=25000)
            page.wait_for_selector("#loginid", timeout=10000)
            page.fill("#loginid", account)
            page.fill("#userpassword", password)
            page.click("#login")
            deadline = time.time() + 25
            while time.time() < deadline:
                try:
                    body = page.inner_text("body")
                except Exception:
                    body = ""
                if any(k in body for k in ("上次登录", "您已经登录", "个人门户")):
                    if verbose:
                        print(f"[auth] 登录成功: {base}", flush=True)
                    return base
                if any(k in body for k in ("密码错误", "用户名或密码", "帐号或密码", "验证码")):
                    raise RuntimeError(f"登录被拒: {body[:200]!r}")
                u = (page.url or "").lower()
                if "login" not in u and body:
                    if verbose:
                        print(f"[auth] 登录成功(跳转): {base}", flush=True)
                    return base
                time.sleep(1)
            last_err = f"{base} 登录后状态未知: {body[:200]!r}"
            if verbose:
                print(f"[auth] {last_err}", flush=True)
        except Exception as e:
            last_err = f"{base} 异常: {str(e)[:150]}"
            if verbose:
                print(f"[auth] {last_err}", flush=True)
    raise RuntimeError(f"所有站点登录失败，最后错误: {last_err}")


if __name__ == "__main__":
    acc, pwd = get_credentials()
    print("凭据就绪，账号:", acc)
