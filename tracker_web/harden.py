#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把默认弱密码替换为强密码：同时更新 tracker.db 与 server.py 种子默认值。"""
import os, re, secrets, sqlite3, hashlib

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracker.db")
SERVER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "server.py")

def hash_pw(pw, salt=None):
    if salt is None:
        salt = secrets.token_hex(8)
    h = hashlib.sha256((salt + pw).encode("utf-8")).hexdigest()
    return salt + ":" + h

# 生成 3 个强密码（url-safe，约 19 字符，无歧义字符）
new_pw = {
    "admin": secrets.token_urlsafe(14),
    "zhang": secrets.token_urlsafe(14),
    "view":  secrets.token_urlsafe(14),
}

# 1) 更新数据库
c = sqlite3.connect(DB)
updated = 0
for u, pw in new_pw.items():
    cur = c.execute("SELECT 1 FROM users WHERE username=?", (u,)).fetchone()
    if cur:
        c.execute("UPDATE users SET password_hash=? WHERE username=?", (hash_pw(pw), u))
        updated += 1
c.commit(); c.close()
print("DB 更新账号数:", updated)

# 2) 同步 server.py 种子默认值（仅当该行还是旧密码时才替换，避免误改）
src = open(SERVER, "r", encoding="utf-8").read()
mapping = {"admin123": new_pw["admin"], "zhang123": new_pw["zhang"], "view123": new_pw["view"]}
changed = 0
for old, npw in mapping.items():
    if old in src:
        src = src.replace('"%s"' % old, '"%s"' % npw, 1)
        changed += 1
open(SERVER, "w", encoding="utf-8").write(src)
print("server.py 种子默认值替换数:", changed)

print("\n===== 新密码（请妥善保存，下同）=====")
print("admin (管理员):", new_pw["admin"])
print("zhang (跟单员):", new_pw["zhang"])
print("view  (只读访客):", new_pw["view"])
print("=====================================")
