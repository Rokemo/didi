#!/usr/bin/env python3
"""
github_push.py — 通过 GitHub REST API 把本地仓库内容推送到 GitHub。

适用场景
--------
当标准 git 传输 (github.com:443) 被网络出口策略封锁、但 api.github.com 可达时，
用 GitHub Git Data API (blob -> tree -> commit -> ref) 绕过 443 封锁完成推送。

设计要点
--------
* 本地仓库是「真相源」：推送的内容来自 `git ls-files`（本地已跟踪文件）。
* 两种模式：
    - 默认（additive）：额外保留远程存在、但本地未跟踪的文件（安全，不会误删远程
      文件，例如 GitHub 自动生成的 README）。
    - --sync：让远程与本地完全一致——本地已 `git rm` 删除的文件会从远程一并删除。
* 认证：环境变量 GH_TOKEN（Personal Access Token，需 Contents: Read and Write）。
* 仓库/分支：环境变量 GH_REPO / GH_BRANCH，或命令行 --repo / --branch。
* 增量友好：每次都在远程当前 HEAD 之上叠加一个新提交（fast-forward，无需 force）。
* 中文路径：用 `git ls-files -z` 读取，避免 git 的 C 风格转义把中文名破坏。

用法
----
  # 只读检查远程状态 + 将要推送的文件清单
  python tools/github_push.py --dry-run

  # 普通推送（保留远程额外文件，如 README）
  GH_TOKEN=xxx python tools/github_push.py

  # 同步模式（远程 = 本地，删除本地已移除的文件）
  GH_TOKEN=xxx python tools/github_push.py --sync

  # 指定仓库 / 分支 / 提交信息
  GH_TOKEN=xxx python tools/github_push.py --repo owner/repo --branch main --message "update"
"""
import os, sys, json, base64, subprocess, argparse, urllib.request, urllib.error

API = "https://api.github.com"
DEFAULT_REPO = "Rokemo/didi"
DEFAULT_BRANCH = "main"


def api(method, path, token=None, data=None):
    url = API + path
    req = urllib.request.Request(url, method=method)
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("User-Agent", "github-push-tool")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    if token:
        req.add_header("Authorization", "Bearer " + token)
    body = None
    if data is not None:
        req.add_header("Content-Type", "application/json")
        body = json.dumps(data).encode()
    try:
        with urllib.request.urlopen(req, data=body, timeout=60) as r:
            txt = r.read().decode()
            return r.status, (json.loads(txt) if txt else {})
    except urllib.error.HTTPError as e:
        return e.code, {"__error__": e.read().decode(errors="replace")[:600]}


def local_files():
    # -z: NUL 分隔且不转义，避免中文路径被 git 包裹成 "\345\220..."
    out = subprocess.check_output(["git", "ls-files", "-z"])
    return [p.decode("utf-8") for p in out.split(b"\x00") if p]


def get_remote_state(repo, branch):
    st, ref = api("GET", f"/repos/{repo}/git/refs/heads/{branch}")
    if st != 200:
        return st, ref
    head = ref["object"]["sha"]
    st, commit = api("GET", f"/repos/{repo}/git/commits/{head}")
    if st != 200:
        return st, commit
    tree_sha = commit["tree"]["sha"]
    st, tree = api("GET", f"/repos/{repo}/git/trees/{tree_sha}?recursive=1")
    if st != 200:
        return st, tree
    files = [(e["path"], e["type"], e["mode"], e["sha"]) for e in tree.get("tree", [])]
    return 200, {"head": head, "tree": tree_sha,
                 "truncated": tree.get("truncated"), "files": files}


def push(repo, branch, token, message, sync):
    st, state = get_remote_state(repo, branch)
    if st != 200:
        return f"远程状态获取失败 {st}: {state}"
    parent = state["head"]
    remote_files = {p: (t, m, s) for (p, t, m, s) in state["files"]}
    print(f"[info] 远程 {branch} = {parent[:10]}, 远程文件数={len(remote_files)}")

    # 1) 上传本地文件为 blob
    tree_entries = {}
    for p in local_files():
        with open(p, "rb") as f:
            raw = f.read()
        b64 = base64.b64encode(raw).decode()
        st, blob = api("POST", f"/repos/{repo}/git/blobs", token=token,
                       data={"content": b64, "encoding": "base64"})
        if st != 201:
            return f"blob 上传失败 {st}: {blob}  (文件 {p})"
        tree_entries[p] = {"path": p, "mode": "100644", "type": "blob", "sha": blob["sha"]}
        print(f"[blob] {p} -> {blob['sha'][:10]}")

    # 2) 处理远程额外文件
    if sync:
        dropped = [p for p in remote_files if p not in tree_entries]
        if dropped:
            print(f"[sync] 将从远程删除 {len(dropped)} 个本地已移除的文件:")
            for p in dropped:
                print("   -", p)
        # sync 模式：tree 仅含本地文件（远程-only 文件不保留）
    else:
        kept = 0
        for p, (t, m, s) in remote_files.items():
            if p not in tree_entries:
                tree_entries[p] = {"path": p, "mode": m, "type": t, "sha": s}
                kept += 1
        if kept:
            print(f"[additive] 保留 {kept} 个远程额外文件（本地未跟踪）")

    # 3) 创建 tree（不使用 base_tree：显式给出完整文件清单，
    #    远程中存在但此处未列出的文件才会被真正删除，保证 --sync 生效）
    st, new_tree = api("POST", f"/repos/{repo}/git/trees", token=token,
                       data={"tree": list(tree_entries.values())})
    if st != 201:
        return f"tree 创建失败 {st}: {new_tree}"
    print(f"[tree] {new_tree['sha'][:10]} (共 {len(tree_entries)} 项)")

    # 4) 创建 commit（父 = 远程当前 HEAD，fast-forward）
    st, new_commit = api("POST", f"/repos/{repo}/git/commits", token=token,
                         data={"message": message, "tree": new_tree["sha"], "parents": [parent]})
    if st != 201:
        return f"commit 创建失败 {st}: {new_commit}"
    print(f"[commit] {new_commit['sha'][:10]}")

    # 5) 更新 ref
    st, upd = api("PATCH", f"/repos/{repo}/git/refs/heads/{branch}",
                  token=token, data={"sha": new_commit["sha"], "force": False})
    if st != 200:
        return f"ref 更新失败 {st}: {upd}"
    return f"SUCCESS 已推送 -> https://github.com/{repo}/commit/{new_commit['sha']}"


def main():
    ap = argparse.ArgumentParser(
        description="通过 GitHub API 推送本地仓库（绕过 github.com:443 封锁）")
    ap.add_argument("--repo", default=os.environ.get("GH_REPO", DEFAULT_REPO), help="owner/repo")
    ap.add_argument("--branch", default=os.environ.get("GH_BRANCH", DEFAULT_BRANCH), help="分支名")
    ap.add_argument("--message", default="chore: 更新仓库内容", help="提交信息")
    ap.add_argument("--sync", action="store_true", help="让远程与本地一致（删除本地已移除的文件）")
    ap.add_argument("--dry-run", action="store_true", help="仅检查远程状态并打印将推送的文件")
    args = ap.parse_args()

    if args.dry_run:
        st, s = get_remote_state(args.repo, args.branch)
        if st != 200:
            print("DRY-RUN FAIL", st, s)
            sys.exit(1)
        print("remote head:", s["head"][:10])
        print("--- 远程文件 ---")
        for p, t, m, sha in s["files"]:
            print(f"  {m:6} {t:5} {p}")
        print("--- 本地将推送的文件 ---")
        for p in local_files():
            print("  +", p)
        return

    tok = os.environ.get("GH_TOKEN")
    if not tok:
        print("缺少 GH_TOKEN 环境变量（需 Contents: Read and Write 权限）")
        sys.exit(1)
    print(push(args.repo, args.branch, tok, args.message, args.sync))


if __name__ == "__main__":
    main()
