# 彩票销售数据 Worker

## 项目简介
本项目通过 Cloudflare Workers 每天定时检查财政部“全国彩票销售情况”最新一期报告。Worker 会对比 D1 中已结构化入库的最新月份；发现新月份后，会按顺序下载报告附件到 R2，将附件索引写入 D1，并解析 Excel 写入月度和地区汇总表。Worker 会优先使用财政部移动站，随后回退到综合信息列表页和全文检索接口。

当前云端数据已经完成一次性导入：

- D1：结构化历史数据与附件索引。
- R2：历史 Excel 附件对象。
- Cron：北京时间每天 10:00 自动运行一次。

## Cloudflare 免费额度速查

以下额度按 Cloudflare 官方文档在 2026-06-04 的说明整理。额度可能调整，最终以链接中的官方页面为准。

### 本项目用到的资源

| 资源 | 当前用途 | 免费额度要点 | 本项目当前占用 |
| --- | --- | --- | --- |
| Workers | API、最新数据同步、定时任务入口 | 100,000 请求/天；每次 HTTP 请求 10 ms CPU；128 MB 内存；每次调用 50 个外部 subrequests；账号最多 100 个 Workers | 1 个 Worker：`lottery-latest-crawler` |
| Cron Triggers | 每天北京时间 10:00 自动执行 `/sync-latest` 同等逻辑 | 免费账号最多 5 个 Cron Triggers；Cron 使用 UTC；变更最多可能需要 15 分钟传播 | 1 个 Cron：`0 2 * * *` |
| Custom Domain | 绑定 `https://lottery.aback.fun` | 每个 zone 最多 100 个 Custom Domains；每个 zone 最多 1,000 条 routes | 1 个 Custom Domain |
| D1 | 保存月度汇总、地区汇总、附件索引 | 5,000,000 rows read/天；100,000 rows written/天；5 GB 总存储；D1 不收数据传输费用 | 3 张业务表，当前用量见 `/state` |
| R2 Standard | 保存历史 Excel 附件 | 10 GB-month/月；1,000,000 Class A 操作/月；10,000,000 Class B 操作/月；互联网 egress 免费 | 当前对象数量见 `/state` |

### 需要注意的计费点

- Workers 请求额度按账号每天统计，免费计划 100,000 次/天，UTC 00:00 重置。
- Workers 的 CPU 时间只计算实际执行代码的 CPU；等待网络请求、D1、R2 的时间不算 CPU。
- D1 的 reads/writes 按扫描或写入的行数计费；Dashboard 或 Wrangler 手动查询也会计入用量。
- D1 免费 reads/writes 每天 UTC 00:00 重置；D1 存储是账号总量。
- R2 免费额度按月计算，且免费层只适用于 Standard storage，不适用于 Infrequent Access。
- R2 `PutObject`、`ListObjects` 等属于 Class A；`GetObject`、`HeadObject` 等属于 Class B；删除对象免费。

### 查看当前项目用量

项目内免费额度接口：

```powershell
Invoke-RestMethod https://lottery.aback.fun/api/usage
```

项目内状态接口：

```powershell
Invoke-RestMethod https://lottery.aback.fun/state
```

Cloudflare Dashboard：

- Workers：Workers & Pages > `lottery-latest-crawler` > Metrics
- Cron：Workers & Pages > `lottery-latest-crawler` > Settings > Triggers
- D1：Workers & Pages > D1 SQL Database > `lottery-data` > Metrics > Row Metrics
- R2：R2 Object Storage > `lottery-downloads` > Metrics

官方文档：

- [Workers Limits](https://developers.cloudflare.com/workers/platform/limits/)
- [Workers Pricing](https://developers.cloudflare.com/workers/platform/pricing/)
- [Cron Triggers](https://developers.cloudflare.com/workers/configuration/cron-triggers/)
- [D1 Pricing](https://developers.cloudflare.com/d1/platform/pricing/)
- [R2 Pricing](https://developers.cloudflare.com/r2/pricing/)

## 主要文件
- `src/index.ts`：Worker 路由、抓取、D1/R2 同步逻辑。
- `wrangler.jsonc`：Worker 名称、D1/R2 绑定和定时任务配置。
- `migrations/0001_initial_schema.sql`：D1 业务表结构。
- `worker-configuration.d.ts`：Wrangler 生成的绑定类型。
- `docs/cloudflare_worker.md`：部署、调试和云端验证命令。
- `docs/api.md`：线上接口文档。

## 本地开发
```powershell
npm install
npm run check
npm run dev
```

定时任务本地调试：

```powershell
npm run dev:scheduled
Invoke-WebRequest "http://127.0.0.1:8787/cdn-cgi/handler/scheduled?cron=0+2+*+*+*"
```

## 部署
```powershell
npm run deploy
```

新建或重建 D1 时，先应用 migration：

```powershell
npm run db:migrations:apply:remote
```

部署后可用以下接口检查状态：

```powershell
Invoke-RestMethod https://lottery.aback.fun/state
Invoke-RestMethod https://lottery.aback.fun/sync-latest
```

## 接口
- `/`：ECharts 彩票销售数据仪表盘，读取预计算快照，不在页面访问时重新聚合历史数据。
- `/health`：基础健康检查。
- `/latest`：抓取并返回远端最新报告信息，不写入存储。
- `/state`：返回 D1 表计数、最新月份和 R2 对象数量。
- `/sync-latest`：同步最新一期；如果 D1/R2 已经包含最新月份则返回 `up_to_date`。
- `/api/dashboard`：返回仪表盘预计算快照。
- `/api/dashboard/rebuild`：手动重建仪表盘快照。
- `/api/usage`：返回 Cloudflare 免费额度、周期和当前可计算占比。
- `/api/months`：返回 D1 中已有的报告月份和各表行数。
- `/api/monthly`：返回月度汇总数据，支持 `date`、`lottery_group`、`lottery_type`、`limit`、`offset`。
- `/api/regions`：返回地区汇总数据，支持 `date`、`region`、`lottery_group`、`limit`、`offset`。
- `/api/attachments`：返回附件索引，支持 `date`、`extension`、`limit`、`offset`。

取数示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/api/usage"
Invoke-RestMethod "https://lottery.aback.fun/api/dashboard"
Invoke-RestMethod "https://lottery.aback.fun/api/months?limit=12"
Invoke-RestMethod "https://lottery.aback.fun/api/monthly?date=2026-04&limit=50"
Invoke-RestMethod "https://lottery.aback.fun/api/regions?date=2015-02&region=北京"
Invoke-RestMethod "https://lottery.aback.fun/api/attachments?date=2026-04&extension=xlsx"
```

## 数据说明
本地 `data/` 与 `downloads/` 目录只作为备份或临时文件位置，默认不纳入版本控制。正式运行以 Cloudflare D1 和 R2 为准。

仪表盘只聚合 D1 中已经结构化入库的销售数据。Worker 会自动同步搜索结果中缺失的新报告，解析 Excel 并写入 `lottery_monthly_summary` / `lottery_region_summary`，然后刷新仪表盘快照。

