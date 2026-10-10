# 美股深度周报（GitHub Actions 自动版）

每周六北京时间约 08:30 运行，采用美东时间最近结束的周一至周五作为报告区间。自动生成中文 Word 周报，通过 Gmail SMTP 以真实本地 DOCX 文件为附件发送，并通过 Gmail IMAP 已发送记录核对邮件附件。

## 上线前必须由账号所有者配置

前往仓库 **Settings → Secrets and variables → Actions → New repository secret**，创建以下四个 Repository secrets：
- `GEMINI_API_KEY`：有可用 API 额度的 Google AI Studio 免费层 Gemini API Key（不要关联付费计费账号，避免收费）。
- `GMAIL_USER`：用于发送报告的 Gmail 地址。
- `GMAIL_APP_PASSWORD`：启用 Google 两步验证后生成的 **应用专用密码**，不要使用 Gmail 登录密码。该功能是否提供取决于账号设置。
- `MAIL_TO`：你希望接收周报的邮箱地址。

不要在聊天、仓库文件、Issue 或 Actions 日志中粘贴任何密码。

报告使用 Gemini 2.5 Flash 免费层 + Google Search Grounding 实时研究，并由 yfinance 抓取公开价格资料。公共数据接口未保证可用，如缺乏关键交易日收盘数据，工作流会明确失败而非编造。模型可通过仓库 **Actions Variables** 里的 `REPORT_MODEL` 覆盖；默认 `gemini-2.5-flash`。Gemini 2.5 Flash 官方免费额度内模型与 Google Search Grounding 可免费使用；超过配额时任务可能报错，绝不自动切换付费模型。Google 免费 API 输入可能用于改进产品，且不保证免费层长期不变。

## 使用

- 手动：GitHub → Actions → **US Market Weekly Report** → Run workflow。
- 测试：选择 `dry_run=true`，只生成报告不发送邮件。
- 自动：每周六北京时间约 08:30；GitHub cron 在高峰时可能延迟。
- Actions 的 **Artifacts** 中可以下载本次 Word 与 `status.json`，工作流失败会在 Actions 页面显示。
- 如 IMAP 核查到本周准确主题的邮件已带 .docx，则自动跳过，不重复发信。
- 若 SMTP 已接受邮件但 IMAP 记录尚未同步，流水线不会不加区分地重复发送，而是报告“已提交，附件尚待验证”。收件箱实际投递不能仅靠已发送文件夹证明。

研究标准：标普 500、纳指 100、VGT；宏观和收益率；行业 3–5 个；个股 4–8 只；财报超预期/低于预期的比较；资金、技术位置、两周催化和条件情景。尽可能用公司 IR / SEC / BLS / Treasury / Fed 等官方来源，区分可验证事实与研究推断，无法确认的数字写明未获得可靠数据。

**注意**：GitHub Actions 不共享 ChatGPT 已连接 Gmail 的授权，也无法从 ChatGPT 获取你的 Gemini API 密钥。必须完成上述四个 Secret 的一次性设置后才可发信。确认连续成功后，再暂停此前 ChatGPT 内的重复定时发送任务。
