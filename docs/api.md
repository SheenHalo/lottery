# 彩票销售数据 API 文档

## 基础信息

- Base URL：`https://lottery.aback.fun`
- 数据格式：JSON，UTF-8
- 当前鉴权：无
- CORS：允许跨域 `GET` 和 `OPTIONS`

## 通用分页

以下取数接口均支持分页：

| 参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `limit` | integer | `100` | 每页条数，最小 `1`，最大 `500` |
| `offset` | integer | `0` | 偏移量，最小 `0`，最大 `100000` |

分页响应统一包含：

```json
{
  "pagination": {
    "limit": 100,
    "offset": 0,
    "total": 7
  },
  "rows": []
}
```

## 日期参数

`/api/monthly`、`/api/regions`、`/api/attachments` 支持 `date` 参数：

- 可传 `YYYY-MM`，例如 `2026-04`
- 可传 `YYYY-MM-01`，例如 `2026-04-01`
- 不传或传 `latest` 时，默认查询该表最新月份

## 免费额度用量

### `GET /api/usage`

返回本项目用到的 Cloudflare 免费额度、额度周期，以及当前能从 Worker 直接计算的占比。

注意：Cloudflare 账号级真实用量，例如 Workers 今日请求数、D1 今日 rows read/write、R2 本月 Class A/B 操作数，不能通过 Worker 绑定直接读取。接口会将这些项标记为 `unavailable`，并保留免费额度与周期信息，真实累计量请在 Cloudflare Dashboard 或 Analytics API 中查看。

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/api/usage"
```

顶层字段：

| 字段 | 说明 |
| --- | --- |
| `checkedAt` | 检查时间 |
| `baseUrl` | 当前 API 基础域名 |
| `note` | 统计说明 |
| `periodSummary` | 按周期归类的额度 key，包含 `daily`、`weekly`、`monthly` 等 |
| `resources` | 当前项目资源信息 |
| `quotas` | 免费额度明细 |

`quotas` 行字段：

| 字段 | 说明 |
| --- | --- |
| `key` | 额度唯一标识 |
| `resource` | Cloudflare 资源类型 |
| `name` | 额度名称 |
| `period` | 额度周期，例如 `daily`、`monthly`、`per_invocation`、`account` |
| `periodLabel` | 中文周期说明，例如 `日限`、`月限` |
| `unit` | 单位 |
| `freeLimit` | 免费额度上限 |
| `currentUsage` | 当前可计算用量；不可从 Worker 读取时为 `null` |
| `remaining` | 剩余额度；不可计算时为 `null` |
| `percentOfFreeLimit` | 当前占免费额度百分比；不可计算时为 `null` |
| `status` | `available` 或 `unavailable` |
| `source` | 用量来源 |
| `notes` | 备注 |

响应片段：

```json
{
  "checkedAt": "2026-06-04T10:00:00.000Z",
  "baseUrl": "https://lottery.aback.fun",
  "periodSummary": {
    "daily": [
      "workers.requests",
      "d1.rows_read",
      "d1.rows_written"
    ],
    "weekly": [],
    "monthly": [
      "r2.storage",
      "r2.class_a_operations",
      "r2.class_b_operations"
    ]
  },
  "quotas": [
    {
      "key": "d1.storage",
      "resource": "D1",
      "name": "D1 存储",
      "period": "total",
      "periodLabel": "总量限制",
      "unit": "bytes",
      "freeLimit": 5000000000,
      "currentUsage": 618496,
      "remaining": 4999381504,
      "percentOfFreeLimit": 0.01,
      "status": "available"
    }
  ]
}
```

## 健康检查

### `GET /`

返回 ECharts 彩票销售数据仪表盘。页面本体由 Worker 静态返回，图表数据来自 `/api/dashboard` 的 D1 预计算快照，不会在每次页面访问时重新扫描并聚合历史数据。

示例：

```powershell
Start-Process "https://lottery.aback.fun/"
```

## 仪表盘快照

### `GET /api/dashboard`

返回仪表盘使用的预计算 JSON 快照，包括全国销售趋势、福彩/体彩趋势、彩票类型结构、同比/环比、地区排名和地区热力图数据。快照存储在 D1 的 `dashboard_cache` 表中；如果首次访问时尚未生成，Worker 会生成一次并写入缓存。快照只聚合 D1 中已经结构化入库的销售数据；正常同步会先解析 Excel 并写入月度、地区汇总，再刷新快照。

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/api/dashboard"
```

### `GET /api/dashboard/rebuild`

手动重建仪表盘快照。`/sync-latest` 发现新报告、写入附件和结构化销售数据后也会尝试刷新该快照。

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/api/dashboard/rebuild"
```

### `GET /health`

返回 Worker 基础可用状态。

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/health"
```

响应：

```json
{
  "ok": true
}
```

## 云端状态

### `GET /state`

返回 D1 三张业务表的行数、最早月份、最新月份，以及 R2 附件数量。

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/state"
```

响应字段：

| 字段 | 说明 |
| --- | --- |
| `monthly` | `lottery_monthly_summary` 表状态 |
| `region` | `lottery_region_summary` 表状态 |
| `attachments` | `report_attachments` 表状态 |
| `r2` | R2 `downloads/` 前缀下对象数量 |

响应示例：

```json
{
  "monthly": {
    "count": 1381,
    "minReportDate": "2010-01-01",
    "maxReportDate": "2026-04-01"
  },
  "region": {
    "count": 2304,
    "minReportDate": "2010-01-01",
    "maxReportDate": "2015-02-01"
  },
  "attachments": {
    "count": 269,
    "minReportDate": "2009-03-01",
    "maxReportDate": "2026-04-01"
  },
  "r2": {
    "prefix": "downloads/",
    "objectCount": 269
  }
}
```

## 远端最新报告

### `GET /latest`

实时访问财政部搜索页和文章页，返回远端最新一期报告信息。该接口只抓取，不写入 D1/R2。

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/latest"
```

响应字段：

| 字段 | 说明 |
| --- | --- |
| `checkedAt` | 检查时间 |
| `title` | 搜索结果标题 |
| `url` | 报告文章 URL |
| `reportDate` | 报告月份，格式 `YYYY-MM-01` |
| `pageTitle` | 文章页标题 |
| `attachments` | 文章附件列表 |

## 同步最新报告

### `GET /sync-latest`

检查远端最新一期报告，并与 D1 中已保存的最新月份对比：

- 如果远端没有更新，返回 `up_to_date`
- 如果远端不可访问，返回 `remote_unavailable`
- 如果发现新月份，按日期顺序下载附件到 R2，写入 `report_attachments`，并解析 Excel 写入月度和地区汇总表
- 如果数据库落后多个已发布月份，会在同一次同步中尝试补齐搜索结果页中可见的缺失报告

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/sync-latest"
```

响应字段：

| 字段 | 说明 |
| --- | --- |
| `checkedAt` | 检查时间 |
| `action` | `created`、`up_to_date` 或 `remote_unavailable` |
| `storedLatestReportDate` | D1 中原本的最新月份 |
| `remoteLatestReportDate` | 远端最新月份 |
| `error` | 远端不可访问时的错误信息 |
| `report` | 新增时的报告详情 |
| `storedAttachments` | 本次写入 R2 的附件信息 |

响应示例：

```json
{
  "checkedAt": "2026-06-03T13:57:44.252Z",
  "action": "up_to_date",
  "storedLatestReportDate": "2026-04-01",
  "remoteLatestReportDate": "2026-04-01",
  "storedAttachments": []
}
```

## 月份列表

### `GET /api/months`

返回 D1 中已有的报告月份，以及每个月在三张表里的行数。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `limit` | integer | 分页条数 |
| `offset` | integer | 分页偏移 |

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/api/months?limit=12"
```

响应行字段：

| 字段 | 说明 |
| --- | --- |
| `reportDate` | 报告月份 |
| `monthlyRows` | 月度汇总行数 |
| `regionRows` | 地区汇总行数 |
| `attachmentRows` | 附件索引行数 |

响应示例：

```json
{
  "pagination": {
    "limit": 3,
    "offset": 0,
    "total": 196
  },
  "rows": [
    {
      "reportDate": "2026-04-01",
      "monthlyRows": 7,
      "regionRows": 0,
      "attachmentRows": 1
    }
  ]
}
```

## 月度汇总

### `GET /api/monthly`

返回 `lottery_monthly_summary` 中的月度汇总数据。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `date` | string | 报告月份；不传默认最新月份 |
| `lottery_group` | string | 彩票大类，可用别名 `lotteryGroup` |
| `lottery_type` | string | 彩票类型，可用别名 `lotteryType` |
| `limit` | integer | 分页条数 |
| `offset` | integer | 分页偏移 |

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/api/monthly?date=2026-04&limit=50"
```

响应行字段：

| 字段 | 说明 |
| --- | --- |
| `reportDate` | 报告月份 |
| `lotteryGroup` | 彩票大类 |
| `lotteryType` | 彩票类型 |
| `monthlySales` | 当月销售额 |

响应示例：

```json
{
  "pagination": {
    "limit": 3,
    "offset": 0,
    "total": 7
  },
  "rows": [
    {
      "reportDate": "2026-04-01",
      "lotteryGroup": "体彩",
      "lotteryType": "乐透数字型",
      "monthlySales": 650317.4023
    }
  ],
  "date": "2026-04-01"
}
```

## 地区汇总

### `GET /api/regions`

返回 `lottery_region_summary` 中的地区汇总数据。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `date` | string | 报告月份；不传默认最新月份 |
| `region` | string | 地区名，可用别名 `region_name`、`regionName` |
| `lottery_group` | string | 彩票大类，可用别名 `lotteryGroup` |
| `limit` | integer | 分页条数 |
| `offset` | integer | 分页偏移 |

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/api/regions?date=2015-02&region=北京"
```

响应行字段：

| 字段 | 说明 |
| --- | --- |
| `reportDate` | 报告月份 |
| `regionName` | 地区名 |
| `lotteryGroup` | 彩票大类 |
| `monthlySales` | 当月销售额 |

响应示例：

```json
{
  "pagination": {
    "limit": 3,
    "offset": 0,
    "total": 64
  },
  "rows": [
    {
      "reportDate": "2015-02-01",
      "regionName": "上海",
      "lotteryGroup": "体彩",
      "monthlySales": 43104.3785
    }
  ],
  "date": "2015-02-01"
}
```

## 附件索引

### `GET /api/attachments`

返回 `report_attachments` 中的附件索引。该接口返回 R2 key，不直接下载文件。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `date` | string | 报告月份；不传默认最新月份 |
| `extension` | string | 附件扩展名，可用别名 `ext`，支持 `pdf`、`doc`、`docx`、`xls`、`xlsx`、`zip`、`rar` |
| `limit` | integer | 分页条数 |
| `offset` | integer | 分页偏移 |

示例：

```powershell
Invoke-RestMethod "https://lottery.aback.fun/api/attachments?date=2026-04&extension=xlsx"
```

响应行字段：

| 字段 | 说明 |
| --- | --- |
| `reportDate` | 报告月份 |
| `pageTitle` | 报告页面标题 |
| `sourceUrl` | 报告文章 URL |
| `attachmentUrl` | 财政部原始附件 URL |
| `originalName` | 页面中的原始附件名 |
| `filename` | 保存后的文件名 |
| `r2Key` | R2 对象 key |
| `extension` | 扩展名 |
| `fileSize` | 文件大小，单位 byte |
| `createdAt` | 写入时间 |
| `updatedAt` | 更新时间 |

响应示例：

```json
{
  "pagination": {
    "limit": 3,
    "offset": 0,
    "total": 1
  },
  "rows": [
    {
      "reportDate": "2026-04-01",
      "pageTitle": "2026年4月份全国彩票销售情况",
      "sourceUrl": "http://zhs.mof.gov.cn/zonghexinxi/202605/t20260525_3990465.htm",
      "attachmentUrl": "http://zhs.mof.gov.cn/zonghexinxi/202605/P020260525375098921904.xlsx",
      "originalName": "附件.xlsx",
      "filename": "2026年4月份全国彩票销售情况_1.xlsx",
      "r2Key": "downloads/2026年4月份全国彩票销售情况_1.xlsx",
      "extension": ".xlsx",
      "fileSize": 23410,
      "createdAt": "2026-06-03 13:07:18",
      "updatedAt": "2026-06-03 13:07:18"
    }
  ],
  "date": "2026-04-01",
  "filters": {
    "extension": ".xlsx"
  }
}
```

## 错误响应

错误响应也是 JSON。

| 状态码 | 场景 |
| --- | --- |
| `400` | 参数错误，例如日期格式、分页参数或扩展名无效 |
| `404` | 路由不存在 |
| `405` | `/api/*` 路由使用了非 `GET` 方法 |
| `500` | D1 查询或 Worker 内部错误 |
| `502` | 财政部远端页面抓取失败 |

错误示例：

```json
{
  "error": "api request failed",
  "detail": "Invalid date; expected YYYY-MM or YYYY-MM-01."
}
```

