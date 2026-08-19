# DeepSeek 今日邮件智能分析（deepseek_mail_analyzer.py）

一个**独立运行**的 Python 脚本：直连你的邮箱（IMAP）读取**今日**邮件，调用
**DeepSeek 官方大语言模型 API** 逐封输出——**重要度评分、传达含义、摘要、建议动作、
发件人可信度**，并自动附上重要邮件的**原文**。

> 关键特性：**完全不经过 DSH，不调用 DS Agent，不消耗任何 DSH/Agent token。**
> 唯一的对外调用是 DeepSeek 官方 API（`https://api.deepseek.com`），计费走你自己的
> DeepSeek API Key。纯 Python 标准库实现，**无需 pip 安装任何第三方包**（Python 3.8+）。

---

## 一、准备工作

### 1. DeepSeek API Key（必须，除 `--dry-run` 外）
- 到 <https://platform.deepseek.com> 注册并创建 API Key（形如 `sk-...`）。
- 充值少量余额即可（分析几十封邮件通常只需几分钱）。

### 2. 邮箱 IMAP 授权码（多数邮箱不能用登录密码）
| 邮箱 | IMAP 服务器 | 端口 | 授权码获取方式 |
|------|-------------|------|----------------|
| QQ 邮箱 | `imap.qq.com` | 993 | 设置 → 账号 → 开启「IMAP/SMTP 服务」→ 生成**授权码** |
| 163 邮箱 | `imap.163.com` | 993 | 设置 → POP3/SMTP/IMAP → 开启 → 获取**授权码** |
| Gmail | `imap.gmail.com` | 993 | 开启两步验证后，生成**应用专用密码** |
| Outlook/Office365 | `outlook.office365.com` | 993 | 账户安全设置中开启 IMAP / 应用密码 |
| 126 邮箱 | `imap.126.com` | 993 | 同 163 |

## 二、配置（三种方式，优先级从低到高）

### 方式 1：修改脚本顶部 `CONFIG` 区
直接编辑 `deepseek_mail_analyzer.py` 顶部的 `DEFAULTS`，填入邮箱与 API Key。

### 方式 2：`.env` 文件（推荐，密钥不入代码）
复制项目里的 `.env.example` 为 `.env`，填入你的配置：

```bash
cp .env.example .env
# 编辑 .env，填入 DEEPSEEK_API_KEY 与邮箱信息
```

`.env` 与脚本同目录时会被自动读取（也可用 `--env 其它路径` 指定）。

### 方式 3：环境变量 / 命令行参数
```bash
export DEEPSEEK_API_KEY=sk-xxxx
export IMAP_HOST=imap.qq.com
export IMAP_USER=you@qq.com
export IMAP_PASSWORD=你的授权码

python3 deepseek_mail_analyzer.py \
  --host imap.qq.com --user you@qq.com --password 授权码 \
  --api-key sk-xxxx
```

## 三、运行

```bash
# 1) 先不花钱测试：内置示例邮件 + 不调用 API
python3 deepseek_mail_analyzer.py --demo --dry-run

# 2) 不连邮箱，看完整效果（会真实调用 DeepSeek API）
python3 deepseek_mail_analyzer.py --demo

# 3) 正式运行：读取今日邮件并分析
python3 deepseek_mail_analyzer.py

# 常用参数
python3 deepseek_mail_analyzer.py --hours 48            # 最近 48 小时
python3 deepseek_mail_analyzer.py --threshold 3         # 重要度 >=3 就输出原文
python3 deepseek_mail_analyzer.py --bodies              # 所有邮件都输出原文
python3 deepseek_mail_analyzer.py --full-body           # 重要邮件原文不截断
python3 deepseek_mail_analyzer.py --json > report.json  # 输出原始 JSON
python3 deepseek_mail_analyzer.py --no-batch            # 逐封分析（更稳但更耗 token）
python3 deepseek_mail_analyzer.py --unread-only         # 只看未读邮件
python3 deepseek_mail_analyzer.py --important-only      # 规则预过滤，营销/订阅类不调用 API
python3 deepseek_mail_analyzer.py --no-cache            # 关闭缓存（每次重新分析）
python3 deepseek_mail_analyzer.py --body-max 1500       # 送入 API 的正文上限调小，更省 token
python3 deepseek_mail_analyzer.py --cache-clean         # 手动清理缓存（移除超过 90 天的旧记录）
python3 deepseek_mail_analyzer.py --cache-clean --cache-max-age 30   # 按 30 天清理
python3 deepseek_mail_analyzer.py --cache-clear         # 彻底清空缓存
```

## 四、为什么每天运行会消耗 token？怎么省？（重要）

### token 花在哪、不花在哪
- **读邮件（IMAP 拉取）本身 0 token**：这只是网络协议，不经过任何大模型。
- **token 只花在两处**：① 每封邮件的正文要作为**输入 token** 发给 DeepSeek LLM 做分析；
  ② 生成“今日总览”还要一次调用。
- 所以“每天消耗 token”的真相是：**每次运行都把当天所有邮件重新发给了 LLM 分析一遍**。

### 本脚本内置的省 token 机制（全部默认开启或一键开启）

| 机制 | 说明 | 效果 |
|------|------|------|
| **分析结果缓存**（默认开） | 按 Message-ID 记忆每封邮件的分析结果，存到 `.mail_analyzer_state.json` | 同一封邮件**只分析一次**，之后每天重复运行 **0 新增 token**；总览也会自动跳过 |
| **缓存定期清理**（默认开） | 每次运行自动移除超过 `CACHE_MAX_AGE_DAYS`（默认 90 天）的旧记录；也可手动 `--cache-clean` / `--cache-clear` | 缓存文件不会无限膨胀 |
| **规则预过滤** `--important-only` | 订阅/周刊/营销类发件人（可 `--skip-senders` 调整）或明显低优先级邮件，用本地关键词规则直接判定，**不发给 LLM** | 典型邮件箱可省掉 50%~90% 的 API 调用 |
| **只看未读** `--unread-only` | 已读邮件完全不处理 | 日常增量运行只处理真正的新邮件 |
| **正文截断** `--body-max` | 送入 API 的正文上限默认 6000 字符，可调小 | 输入 token 与成本几乎成正比 |
| **跳过总览** `--no-overview` | 不生成今日总览 | 省最后一次调用 |
| **批量分析**（默认） | 一批 20 封合并在一次请求里分析 | 比逐封少很多请求开销 |

### 完全 0 token 的选项
- `--dry-run`：用本地关键词规则代替 LLM（**0 token**），但没有智能总结；可与 `--important-only` 之外的任意参数组合，先预览效果。
- 把脚本放进 `crontab` 每天定时跑（`--unread-only` + 缓存），通常每天只分析新到的几封邮件。

### 成本估算（2026-08-17 起 DeepSeek V4 峰谷定价）
`deepseek-chat`（V4-Flash）：空闲时段 输入 ¥1.5/百万 tokens、输出 ¥4.5/百万 tokens；
**高峰时段（每天 9:00-14:00）约翻倍**。一封 6000 字的邮件约几千输入 token + 几百输出 token，
约 **¥0.01~0.02/封**。脚本每次运行末尾会打印实际 token 与预估费用
（价格可在脚本顶部 `price_input`/`price_output` 或环境变量 `DEEPSEEK_PRICE_INPUT/OUTPUT` 调整）。

**建议**：每天在非高峰时段（如早上 8 点前）定时运行，配合缓存 + `--unread-only`，
一个月通常只需几分钱。

## 五、输出示例（节选）

```
════════════════════════════════════════════════════════════════
  今日邮件智能分析报告  2025-01-15 09:41
════════════════════════════════════════════════════════════════
统计：共 4 封 ｜ 重要(≥4) 1 封 ｜ 时间范围：01-15 00:00 起

【今日总览】
今天邮件整体中等繁忙，最需要优先处理的是服务器告警；会议改期通知
需要尽快确认；其余为退款通知与订阅周刊，可稍后阅读……

────────────────────────────────────────────────────────────────
[重要] 1. ★★★★★  紧急 ｜ 发件人可信度：可信  未读
   时间：2025-01-15 09:00  ｜  发件人：监控平台 <alert@monitor.example.com>
   主题：[告警] 生产服务器 web-03 CPU 持续 95%，请立即处理
   含义：生产服务器负载异常，需要运维介入……
   摘要：……
   建议：立即处理
   ── 原文 ──
   告警详情
   ……
```

## 五、桌面报告 + 开机自动分析（macOS）

### 手动弹窗
```bash
python3 deepseek_mail_analyzer.py --desktop-report   # 把报告写成 HTML 放到桌面
python3 deepseek_mail_analyzer.py --popup            # 系统通知 + 有重要邮件时自动在浏览器打开报告
```

### 开机自动分析（LaunchAgent，已配置）
已创建 `/Users/huzi/Library/LaunchAgents/com.huzi.deepseek-mail-analyzer.plist` 并加载，
效果：**每次登录电脑自动运行**（另加每天 09:30 兜底一次），流程为：
读 qmul 邮箱 → 分析（只看未读 + 过滤营销，最省）→ 桌面生成 HTML 报告 →
弹出系统通知（重要邮件会自动打开浏览器报告），**不用你记得去开**。

- 运行日志：`/tmp/mail-analyzer.out.log`（stdout）、`/tmp/mail-analyzer.err.log`（stderr）
- 临时停用（本次会话）：`launchctl bootout gui/$(id -u)/com.huzi.deepseek-mail-analyzer`
- 彻底移除：删除上面的 plist 文件再执行 `launchctl bootout gui/$(id -u)/com.huzi.deepseek-mail-analyzer`
- 改运行参数：编辑 plist 里的 `ProgramArguments`，然后 `launchctl bootout ... && launchctl bootstrap gui/$(id -u) <plist路径>` 重新加载

## 六、多个邮箱一起分析

在 `.env` 里用 `MAILBOX2_` / `MAILBOX3_` ... 前缀追加更多邮箱，一次运行会**合并分析所有邮箱**（同一份报告、同一个弹窗），每封邮件会标注来自哪个邮箱：

```ini
# 主邮箱（qmul）照旧 ...
IMAP_HOST=outlook.office365.com
IMAP_USER=your-id@qmul.ac.uk
IMAP_AUTH=oauth2

# 追加北邮（腾讯企业邮）
MAILBOX2_NAME=bupt
MAILBOX2_IMAP_HOST=imap.exmail.qq.com
MAILBOX2_IMAP_PORT=993
MAILBOX2_IMAP_USER=你的学号@bupt.edu.cn
MAILBOX2_IMAP_PASSWORD=你的客户端专用密码
MAILBOX2_IMAP_AUTH=password
```

支持的前缀变量：`MAILBOX2_NAME / IMAP_HOST / IMAP_PORT / IMAP_USER / IMAP_PASSWORD / IMAP_AUTH / IMAP_FOLDER / IMAP_SSL / OAUTH_TOKEN_FILE`（OAuth 邮箱务必单独指定令牌文件，如 `MAILBOX2_OAUTH_TOKEN_FILE=.outlook_oauth_2.json`）。

> 北邮邮箱 = 腾讯企业邮：IMAP 服务器 `imap.exmail.qq.com`（或 `imap.bupt.edu.cn`），SSL 端口 993。第三方客户端登录密码一般需要在网页版邮箱设置里生成"**客户端专用密码**"（设置 → 收发信设置 → 客户端设置），账号密码可能无法直接登录 IMAP。

## 七、常见问题

- **登录失败 / AUTH 错误**：QQ、163、Gmail 必须用**授权码/应用专用密码**，不是登录密码；检查 IMAP 服务器是否已开启。
- **SSL 报错**：换用 `--no-ssl`（端口 143）试试，或确认端口填 993。
- **今日没有邮件**：脚本默认取“今日 0 点起”，可加 `--hours 24` 扩大范围。
- **API 报 401/402**：检查 `DEEPSEEK_API_KEY` 是否正确、账户是否有余额。
- **想完全离线试用**：`--demo --dry-run` 不需要任何 Key 和网络。
- **想只测邮箱连通性**：`--dry-run`（不调用 API，只读邮件）。
