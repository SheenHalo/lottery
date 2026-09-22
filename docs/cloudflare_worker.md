# Cloudflare Worker 爬虫

本项目使用 TypeScript Worker 在 Cloudflare Workers 上每天抓取最新一期“全国彩票销售情况”。

## 当前能力

- 仅抓取财政部搜索结果第 1 页。
- 选择标题中可解析出年月的报告，并按日期顺序补齐 D1 中缺失的月份。
- 优先从财政部移动站抓取报告链接；移动站不可用时，回退到综合信息列表页和全文检索接口。
- 访问文章页并提取附件链接。
- HTTP 访问 `/latest` 返回 JSON。
- HTTP 访问 `/state` 返回 D1 表计数和 R2 对象计数。
- HTTP 访问 `/sync-latest` 会同步缺失的新报告，下载附件、解析 Excel 并写入月度/地区汇总；如果月度数据已到最新月份，会直接返回 `up_to_date`。
- HTTP 访问 `/` 返回 ECharts 彩票销售数据仪表盘。
- HTTP 访问 `/api/dashboard` 返回仪表盘预计算快照，避免页面访问时重复聚合历史数据。
- HTTP 访问 `/api/usage` 查看免费额度周期和当前可计算占比。
- HTTP 访问 `/api/months`、`/api/monthly`、`/api/regions`、`/api/attachments` 读取 D1 数据。
- Cron Trigger 每天执行一次 `/sync-latest` 的同等逻辑。
- Excel 附件会解析为 `lottery_monthly_summary` 和 `lottery_region_summary`，支持从已落后的月份开始逐期补齐。

## 云端资源

- Worker：`lottery-latest-crawler`
- D1：`lottery-data`
- R2：`lottery-downloads`
- R2 key 前缀：`downloads/`

## D1 表结构

D1 业务表结构保存在 `migrations/0001_initial_schema.sql`。新建或重建 D1 时执行：

```powershell
npm run db:migrations:apply:remote
```

本地 D1 调试可执行：

```powershell
npm run db:migrations:apply:local
```

## 本地运行

```powershell
npm install
npm run check
npm run dev
```

访问：

```powershell
Invoke-RestMethod http://127.0.0.1:8787/latest
```

测试定时任务：

```powershell
npm run dev:scheduled
Invoke-WebRequest "http://127.0.0.1:8787/cdn-cgi/handler/scheduled?cron=0+2+*+*+*"
```

## 部署

```powershell
npm run deploy
```

`wrangler.jsonc` 中的 cron 为 `0 2 * * *`，Cloudflare Cron 使用 UTC 时间，对应北京时间每天 10:00。

验证云端状态：

```powershell
Invoke-RestMethod https://lottery.aback.fun/state
Invoke-RestMethod https://lottery.aback.fun/sync-latest
```

## D1 取数接口

所有取数接口都支持 `limit` 和 `offset`，默认 `limit=100`，最大 `limit=500`。`date` 可以传 `YYYY-MM` 或 `YYYY-MM-01`；不传时默认查对应表最新月份。

```powershell
Invoke-RestMethod "https://lottery.aback.fun/api/usage"
Invoke-RestMethod "https://lottery.aback.fun/api/months?limit=12"
Invoke-RestMethod "https://lottery.aback.fun/api/monthly?date=2026-04&limit=50"
Invoke-RestMethod "https://lottery.aback.fun/api/regions?date=2015-02&region=北京"
Invoke-RestMethod "https://lottery.aback.fun/api/attachments?date=2026-04&extension=xlsx"
```

