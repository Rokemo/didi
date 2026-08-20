# Cloudflare 命名隧道 · 详细设置指南（稳定外网链接）

> 目标：把本机 `http://127.0.0.1:8731` 的跟单系统，通过 Cloudflare 命名隧道开放到外网，
> 得到一个**永久固定**的外网地址（重启电脑、重跑脚本都不变）。
> 耗时约 10 分钟，完全免费，只需一个 Cloudflare 账号。

---

## 原理（为什么它"稳定"）

| 对比 | 快速隧道 `cloudflared tunnel --url` | **命名隧道 `cloudflared tunnel run 名字`** |
|------|------------------------------------|--------------------------------------------|
| 外网地址 | 每次运行都随机换 `xxx.trycloudflare.com` | **固定** `https://<隧道UUID>.cfargotunnel.com`，永不改变 |
| 需要登录 | 不需要 | 需要（一次性，免费账号） |
| 适合 | 临时演示 | **长期稳定使用（推荐）** |

命名隧道创建后，Cloudflare 会给它一个**固定的 UUID**，该 UUID 就是地址的一部分；只要你本机用同一个凭证运行，地址永远不变。

> 如果你有自己的域名（托管在 Cloudflare），还能绑定成 `didi.你的域名.com` 这种更好记的地址（可选，见第 5 步）。

---

## 第 0 步：准备（一次性检查）

- [ ] 一台**能长期开机**的电脑（跑跟单系统和隧道；关机关隧道）
- [ ] 一个免费 Cloudflare 账号：https://dash.cloudflare.com/sign-up
- [ ] （可选）你的域名托管在 Cloudflare（用于自定义子域名，没有也能用）

---

## 第 1 步：安装 cloudflared（两种方式任选）

**方式 A（推荐，winget 一条命令）**
打开 PowerShell 或 CMD，执行：
```
winget install --id Cloudflare.cloudflared
```
装完重开窗口，验证：
```
cloudflared --version
```
能看到版本号即成功。

**方式 B（手动下载）**
1. 打开 https://github.com/cloudflare/cloudflared/releases
2. 下载 Windows 版 `cloudflared-windows-amd64.exe`
3. 重命名为 `cloudflared.exe`，放到 `C:\Windows`（这样任何窗口都能直接调用）
4. 验证：`cloudflared --version`

> 本机如果没有 winget 或下载慢，优先方式 B；下载地址若打不开，用手机热点试试。

---

## 第 2 步：登录 Cloudflare（一次性）

```
cloudflared login
```
- 会打开浏览器，要求登录你的 Cloudflare 账号；
- 选一个域名授权（没有域名也能授权成功，用于生成凭证）；
- 成功后会在 `C:\Users\你的用户名\.cloudflared\cert.pem` 生成登录凭证。
- 如果浏览器没自动弹，终端里会有个 URL，手动复制到浏览器打开。

> 这一步**只需做一次**。之后这台电脑就可以创建/运行命名隧道。

---

## 第 3 步：创建命名隧道（一次性）

```
cloudflared tunnel create didi
```
输出类似：
```
Created tunnel didi with id 9b1f4e2c-xxxx-xxxx-xxxx-xxxxxxxxxxxx
Created credentials file C:\Users\你的用户名\.cloudflared\9b1f4e2c-xxx.json
```
记下这个 **UUID**。你的固定外网地址就是：
```
https://9b1f4e2c-xxxx-xxxx-xxxx-xxxxxxxxxxxx.cfargotunnel.com
```
**这个地址永久有效。**（可以先复制到记事本存着）

> 查看已创建的隧道：`cloudflared tunnel list`

---

## 第 4 步：写配置文件，让隧道指向本地 8731

用记事本在 `C:\Users\你的用户名\.cloudflared\config.yml` 新建一个文件（UTF-8 编码），内容：

```yaml
tunnel: didi
credentials-file: C:\Users\你的用户名\.cloudflared\9b1f4e2c-xxxx-xxxx-xxxx-xxxxxxxxxxxx.json
ingress:
  - service: http://localhost:8731
```

> 把 `credentials-file` 那行的路径改成你第 3 步实际生成的 `.json` 文件路径。
> ⚠️ **注意（容易踩的坑）**：不绑定自定义域名时，ingress **只写这一条**即可，
> 不要再加 `- service: http_status:404` 兜底行——因为第一条没写 hostname 会匹配所有流量，
> 再加 404 行会导致 cloudflared 报错 `Rule #1 is matching the hostname ''...` 拒绝启动。
> 只有绑定自定义域名（第 5 步）且每条规则都写了 hostname 后，才需要在最后加 404 兜底。

---

## 第 5 步（可选）绑定自己的域名子地址

如果想让地址变成 `didi.你的域名.com`（需要域名托管在 Cloudflare）：
```
cloudflared tunnel route dns didi didi.你的域名.com
```
然后修改 `config.yml` 的 ingress 第一条，加 hostname：
```yaml
ingress:
  - hostname: didi.你的域名.com
    service: http://localhost:8731
  - service: http_status:404
```
`route dns` 会自动建一条 CNAME 指向隧道。约 1 分钟后可用。

> 没有域名就跳过本步，直接用第 3 步的 `<UUID>.cfargotunnel.com` 地址即可。

---

## 第 6 步：日常启动

**方式一：双击我们已做好的 `外网_稳定_Cloudflare.bat`**
它会：后台启动本地服务（`server.py`）→ 检查 cloudflared → 运行 `cloudflared tunnel run didi`。
窗口里会打印出固定地址，复制给任何人就能访问。

**方式二：手动两步**
```
# 窗口 1：启动本地服务
启动.bat

# 窗口 2：启动隧道
cloudflared tunnel run didi
```

> 看到 `Registered tunnel connection` / `Requesting new certificate` 等日志且不退出 = 隧道已上线。
> 关隧道窗口 = 外网访问停止（本地服务仍在）。

---

## 第 7 步（推荐）后台常驻 + 开机自启

让隧道和本地服务在**后台**跑、开机自动启动，不用每次手点：

1. **先把配置文件放到默认位置**（第 4 步已做），并测一次 `cloudflared tunnel run didi` 能正常连上；
2. 管理员身份打开 CMD，把 cloudflared 注册成 Windows 服务：
```
cloudflared service install
```
   - 它会读取 `config.yml`，以后随 Windows 自动启动、崩溃自动拉起；
3. 本地服务 `server.py` 的开机自启：按 `Win+R` → 输入 `shell:startup` → 把 `启动.bat` 的**快捷方式**放进去（或建"任务计划程序"指向它）。

> 想要手动控制服务：`net start cloudflared` / `net stop cloudflared`。

---

## 第 8 步：验证

| 检查项 | 方法 | 期望 |
|--------|------|------|
| 本地服务 | 浏览器开 `http://127.0.0.1:8731` | 能打开 |
| 隧道日志 | 隧道窗口无报错 | `Registered tunnel connection` |
| 外网访问 | 手机流量（关 Wi-Fi）打开固定地址 | 能打开页面 |
| 稳定性 | 重启电脑后再开 | 地址不变，重新连上即可 |

---

## ⚠️ 重要：安全提醒（必读）

当前跟单系统**默认免登录**（`AUTH_ENABLED=0`）——也就是说，一旦开放外网，
**任何拿到你地址的人都能查看甚至修改你的数据**（订单、合同金额、工序进度）。

开放外网前请务必做至少一项：
1. **开启密码登录**：启动服务前设置环境变量 `AUTH_ENABLED=1`（然后浏览器会要求登录，
   用户名密码在 `server.py` 或首次启动日志中；这会让外网访问需要密码）；或
2. 只把地址发给可信同事，且**用完即关隧道**（关掉隧道窗口）；
3. 有域名的话，可以用 Cloudflare Access（免费额度）做邮件/口令二次鉴权，最稳。

---

## 常见问题

- **`cloudflared login` 打不开浏览器** → 复制终端里的 URL 到浏览器手动打开。
- **`cloudflared tunnel create` 报错** → 确认已 login；`cloudflared tunnel list` 看是否已存在。
- **`tunnel run` 一直重连** → 检查 `config.yml` 里 `tunnel:` 名称与 `credentials-file` 路径是否正确、文件是否存在。
- **外网打不开但本地能开** → 用手机热点（关 Wi-Fi）测试，排除公司网络拦截；看隧道窗口有无报错。
- **地址每次变** → 你用的是快速隧道（`cloudflared tunnel --url ...`），要换成命名隧道（`cloudflared tunnel run didi`）。
- **想删除隧道** → `cloudflared tunnel delete didi`。

---

## 与 WorkBuddy 一键外网的区别

| | WorkBuddy 一键外网 | Cloudflare 命名隧道（本方案） |
|--|--------------------|-------------------------------|
| 数据 | 生成**独立副本**（独立库/密码，曾出现"账号密码错误"） | 直连你本机数据库，无副本 |
| 地址 | 每次不同 | **固定** |
| 稳定性 | 依赖 WorkBuddy 会话 | 依赖你本机 + Cloudflare 边缘，非常稳定 |
| 费用 | 随订阅 | 免费 |

**结论**：要长期、稳定、数据一致，用命名隧道；一键外网只适合临时看看。
