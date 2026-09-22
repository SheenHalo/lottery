import * as XLSX from "xlsx";

const SEARCH_ENDPOINT = "https://search.mof.gov.cn/was5/web/search";
const REPORT_INDEX_ENDPOINT = "https://zhs.mof.gov.cn/zonghexinxi/index.htm";
const MOBILE_REPORT_INDEX_ENDPOINT = "https://m.mof.gov.cn/";

const SEARCH_PARAMS: Record<string, string> = {
  channelid: "274731",
  searchword: "全国彩票销售情况",
  keyword: "全国彩票销售情况",
  perpage: "10",
  outlinepage: "10",
  searchscope: "",
  timescope: "",
  page: "1",
};

const DEFAULT_HEADERS = {
  "User-Agent":
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
  "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
  "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
};

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

const DEFAULT_PAGE_LIMIT = 100;
const MAX_PAGE_LIMIT = 500;
const MAX_PAGE_OFFSET = 100_000;
const BYTES_PER_GB = 1_000_000_000;
const DASHBOARD_CACHE_KEY = "main";
const YOY_MINIMUM_BASE_SALES = 10_000;
const BILLION_YUAN_TO_WAN = 10_000;

const FREE_LIMITS = {
  workersRequestsPerDay: 100_000,
  workersCpuMsPerInvocation: 10,
  workersMemoryBytes: 128 * 1024 * 1024,
  workersExternalSubrequestsPerInvocation: 50,
  workersInternalSubrequestsPerInvocation: 1_000,
  workersPerAccount: 100,
  cronTriggersPerAccount: 5,
  customDomainsPerZone: 100,
  routesPerZone: 1_000,
  d1RowsReadPerDay: 5_000_000,
  d1RowsWrittenPerDay: 100_000,
  d1StorageBytes: 5 * BYTES_PER_GB,
  r2StorageBytesMonth: 10 * BYTES_PER_GB,
  r2ClassAOperationsPerMonth: 1_000_000,
  r2ClassBOperationsPerMonth: 10_000_000,
} as const;

const ATTACHMENT_EXTENSIONS = new Set([
  ".pdf",
  ".doc",
  ".docx",
  ".xls",
  ".xlsx",
  ".zip",
  ".rar",
]);

interface ArticleLink {
  url: string;
  title: string;
  reportDate: string;
}

interface Attachment {
  url: string;
  originalName: string;
  extension: string;
}

interface LatestReport {
  checkedAt: string;
  title: string;
  url: string;
  reportDate: string;
  pageTitle: string;
  attachments: Attachment[];
}

interface SyncResult {
  checkedAt: string;
  action: "created" | "remote_unavailable" | "up_to_date";
  storedLatestReportDate: string | null;
  remoteLatestReportDate: string | null;
  error?: string;
  report?: LatestReport;
  storedAttachments: StoredAttachment[];
}

interface StoredAttachment {
  filename: string;
  r2Key: string;
  fileSize: number;
}

interface StoredAttachmentResult {
  stored: StoredAttachment;
  content?: ArrayBuffer;
}

interface ParsedMonthlyRecord {
  lotteryGroup: string;
  lotteryType: string;
  monthlySales: number;
}

interface ParsedRegionRecord {
  regionName: string;
  lotteryGroup: string;
  monthlySales: number;
}

interface ParsedLotteryData {
  monthly: ParsedMonthlyRecord[];
  region: ParsedRegionRecord[];
}

interface DatabaseState {
  monthly: TableState;
  region: TableState;
  attachments: TableState;
  r2: R2State;
}

interface TableState {
  count: number;
  minReportDate: string | null;
  maxReportDate: string | null;
}

interface R2State {
  prefix: string;
  objectCount: number;
  totalSizeBytes: number;
  totalSizeHuman: string;
}

interface Pagination {
  limit: number;
  offset: number;
  total: number;
}

interface PageRequest {
  limit: number;
  offset: number;
}

interface ApiRowsResponse<Row> {
  date?: string | null;
  filters?: Record<string, string>;
  pagination: Pagination;
  rows: Row[];
}

interface AvailableMonthRow {
  reportDate: string;
  monthlyRows: number;
  regionRows: number;
  attachmentRows: number;
}

interface MonthlySummaryRow {
  reportDate: string;
  lotteryGroup: string;
  lotteryType: string;
  monthlySales: number | null;
}

interface RegionSummaryRow {
  reportDate: string;
  regionName: string;
  lotteryGroup: string;
  monthlySales: number | null;
}

interface AttachmentRow {
  reportDate: string;
  pageTitle: string;
  sourceUrl: string | null;
  attachmentUrl: string | null;
  originalName: string | null;
  filename: string;
  r2Key: string;
  extension: string | null;
  fileSize: number | null;
  createdAt: string | null;
  updatedAt: string | null;
}

interface DashboardCacheRow {
  report_date: string | null;
  payload: string;
  updated_at: string | null;
}

interface DashboardSnapshot {
  generatedAt: string;
  source: {
    unit: "万元";
    salesDisplayUnit: "亿元";
    note: string;
  };
  coverage: {
    monthly: TableState;
    region: TableState;
    attachments: TableState;
  };
  latest: {
    reportDate: string | null;
    totalSales: number;
    welfareSales: number;
    sportsSales: number;
    monthOverMonthPercent: number | null;
    yearOverYearPercent: number | null;
    topType: string | null;
    topTypeSales: number;
    topTypeSharePercent: number;
    regionReportDate: string | null;
  };
  monthlyTrend: DashboardTrendPoint[];
  annualTrend: DashboardAnnualPoint[];
  groupTrend: DashboardGroupPoint[];
  typeTrend: DashboardTypePoint[];
  latestGroupBreakdown: DashboardBreakdownPoint[];
  latestTypeBreakdown: DashboardTypeBreakdownPoint[];
  growthBars: DashboardGrowthPoint[];
  regionLatest: {
    reportDate: string | null;
    rows: DashboardRegionPoint[];
  };
  regionHeatmap: {
    months: string[];
    regions: string[];
    values: Array<[number, number, number]>;
  };
}

interface DashboardTrendPoint {
  reportDate: string;
  totalSales: number;
  welfareSales: number;
  sportsSales: number;
  monthOverMonthPercent: number | null;
  yearOverYearPercent: number | null;
}

interface DashboardAnnualPoint {
  year: string;
  totalSales: number;
  welfareSales: number;
  sportsSales: number;
  welfareSharePercent: number;
  sportsSharePercent: number;
  monthCount: number;
  isPartial: boolean;
}

interface DashboardGroupPoint {
  reportDate: string;
  lotteryGroup: string;
  monthlySales: number;
}

interface DashboardTypePoint {
  reportDate: string;
  lotteryGroup: string;
  lotteryType: string;
  monthlySales: number;
}

interface DashboardBreakdownPoint {
  name: string;
  monthlySales: number;
  sharePercent: number;
}

interface DashboardTypeBreakdownPoint {
  lotteryGroup: string;
  lotteryType: string;
  monthlySales: number;
  sharePercent: number;
}

interface DashboardGrowthPoint {
  reportDate: string;
  monthOverMonthPercent: number | null;
  yearOverYearPercent: number | null;
}

interface DashboardRegionPoint {
  regionName: string;
  welfareSales: number;
  sportsSales: number;
  totalSales: number;
}

interface UsageResponse {
  checkedAt: string;
  baseUrl: string;
  note: string;
  periodSummary: Record<QuotaPeriod, string[]>;
  resources: {
    worker: {
      name: string;
      customDomain: string;
      cron: string;
    };
    d1: {
      databaseName: string;
      databaseId: string;
      estimatedStorageBytes: number | null;
      estimatedStorageHuman: string | null;
    };
    r2: R2State;
  };
  quotas: UsageQuota[];
}

type QuotaPeriod =
  | "daily"
  | "weekly"
  | "monthly"
  | "per_invocation"
  | "account"
  | "zone"
  | "total";

interface UsageQuota {
  key: string;
  resource: "Workers" | "Cron" | "Custom Domain" | "D1" | "R2";
  name: string;
  period: QuotaPeriod;
  periodLabel: string;
  unit: string;
  freeLimit: number;
  currentUsage: number | null;
  remaining: number | null;
  percentOfFreeLimit: number | null;
  status: "available" | "unavailable";
  source: "worker_config" | "worker_runtime" | "d1_meta" | "r2_list" | "cloudflare_analytics_required";
  notes?: string;
}

type QueryValue = string | number;
type ReportTable =
  | "lottery_monthly_summary"
  | "lottery_region_summary"
  | "report_attachments";

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    if (url.pathname === "/health") {
      return jsonResponse({ ok: true });
    }

    if (url.pathname === "/") {
      if (request.method !== "GET") {
        return htmlResponse("Method not allowed", 405);
      }
      return htmlResponse(DASHBOARD_HTML);
    }

    if (url.pathname.startsWith("/api/")) {
      if (request.method !== "GET") {
        return jsonResponse({ error: "Method not allowed" }, 405);
      }

      try {
        const body = await handleApiRequest(url, env);
        const cacheHeaders: Record<string, string> =
          url.pathname === "/api/dashboard"
            ? {
                "Cache-Control": "no-store, max-age=0",
              }
            : {};
        return jsonResponse(body, 200, cacheHeaders);
      } catch (error) {
        console.error("api request failed", serializeError(error));
        return jsonResponse(
          { error: "api request failed", detail: errorMessage(error) },
          statusForError(error),
        );
      }
    }

    if (url.pathname === "/state") {
      return jsonResponse(await getCloudState(env));
    }

    if (url.pathname === "/sync-latest") {
      try {
        return jsonResponse(await syncLatestReport(env));
      } catch (error) {
        console.error("latest sync failed", serializeError(error));
        return jsonResponse(
          { error: "latest sync failed", detail: errorMessage(error) },
          502,
        );
      }
    }

    if (url.pathname !== "/latest") {
      return jsonResponse({ error: "Not found" }, 404);
    }

    try {
      const latest = await crawlLatestReport();
      return jsonResponse(latest);
    } catch (error) {
      console.error("latest crawl failed", serializeError(error));
      return jsonResponse(
        { error: "latest crawl failed", detail: errorMessage(error) },
        502,
      );
    }
  },

  scheduled(
    controller: ScheduledController,
    env: Env,
    ctx: ExecutionContext,
  ): void {
    ctx.waitUntil(runScheduledCrawl(controller, env));
  },
};

async function handleApiRequest(url: URL, env: Env): Promise<unknown> {
  if (url.pathname === "/api/dashboard") {
    return getDashboardSnapshot(env.DB);
  }
  if (url.pathname === "/api/dashboard/rebuild") {
    return refreshDashboardSnapshot(env.DB);
  }
  if (url.pathname === "/api/usage") {
    return getUsage(env);
  }
  if (url.pathname === "/api/months") {
    return getAvailableMonths(env.DB, url);
  }
  if (url.pathname === "/api/monthly") {
    return getMonthlySummary(env.DB, url);
  }
  if (url.pathname === "/api/regions") {
    return getRegionSummary(env.DB, url);
  }
  if (url.pathname === "/api/attachments") {
    return getAttachments(env.DB, url);
  }

  throw new HttpError(404, "API route not found.");
}

async function runScheduledCrawl(
  controller: ScheduledController,
  env: Env,
): Promise<void> {
  const result = await syncLatestReport(env);
  console.log(
    JSON.stringify({
      event: "lottery_latest_sync",
      cron: controller.cron,
      scheduledTime: new Date(controller.scheduledTime).toISOString(),
      action: result.action,
      storedLatestReportDate: result.storedLatestReportDate,
      remoteLatestReportDate: result.remoteLatestReportDate,
      attachments: result.storedAttachments.length,
      url: result.report?.url,
    }),
  );
}

async function syncLatestReport(env: Env): Promise<SyncResult> {
  const storedLatestReportDate = await getLatestStoredReportDate(env.DB);
  let latestLinks: ArticleLink[];
  try {
    latestLinks = await crawlReportLinks();
  } catch (error) {
    return {
      checkedAt: new Date().toISOString(),
      action: "remote_unavailable",
      storedLatestReportDate,
      remoteLatestReportDate: null,
      error: errorMessage(error),
      storedAttachments: [],
    };
  }

  const remoteLatestLink = latestLinks[0];
  const newLinks = latestLinks
    .filter(
      (link) =>
        !storedLatestReportDate || link.reportDate > storedLatestReportDate,
    )
    .sort((left, right) => left.reportDate.localeCompare(right.reportDate));

  if (newLinks.length === 0) {
    return {
      checkedAt: new Date().toISOString(),
      action: "up_to_date",
      storedLatestReportDate,
      remoteLatestReportDate: remoteLatestLink.reportDate,
      storedAttachments: [],
    };
  }

  const storedAttachments: StoredAttachment[] = [];
  let latestReport: LatestReport | undefined;
  for (const link of newLinks) {
    latestReport = await fetchReportDetails(link);
    const reportAttachments = await syncReportData(env, latestReport);
    storedAttachments.push(...reportAttachments);
  }

  try {
    await refreshDashboardSnapshot(env.DB);
  } catch (error) {
    console.error("dashboard snapshot refresh failed", serializeError(error));
  }

  return {
    checkedAt: new Date().toISOString(),
    action: "created",
    storedLatestReportDate,
    remoteLatestReportDate: remoteLatestLink.reportDate,
    report: latestReport,
    storedAttachments,
  };
}

async function syncReportData(
  env: Env,
  report: LatestReport,
): Promise<StoredAttachment[]> {
  const storedAttachments: StoredAttachment[] = [];
  const parsedMonthly = new Map<string, ParsedMonthlyRecord>();
  const parsedRegion = new Map<string, ParsedRegionRecord>();
  let spreadsheetCount = 0;
  let parsedWorkbookCount = 0;
  const parseErrors: string[] = [];

  for (const [index, attachment] of report.attachments.entries()) {
    const storedResult = await storeAttachment(
      env.DOWNLOADS,
      report,
      attachment,
      index + 1,
    );
    await upsertAttachmentMetadata(
      env.DB,
      report,
      attachment,
      storedResult.stored,
    );
    storedAttachments.push(storedResult.stored);

    if (isSpreadsheetAttachment(attachment) && storedResult.content) {
      spreadsheetCount += 1;
      try {
        const parsed = parseLotteryWorkbook(storedResult.content);
        if (parsed.monthly.length > 0 || parsed.region.length > 0) {
          parsedWorkbookCount += 1;
        }
        for (const row of parsed.monthly) {
          parsedMonthly.set(`${row.lotteryGroup}|${row.lotteryType}`, row);
        }
        for (const row of parsed.region) {
          parsedRegion.set(`${row.regionName}|${row.lotteryGroup}`, row);
        }
      } catch (error) {
        parseErrors.push(`${attachment.originalName}: ${errorMessage(error)}`);
      }
    }
  }

  if (spreadsheetCount > 0 && parsedWorkbookCount === 0) {
    throw new Error(
      `No structured lottery data parsed from ${spreadsheetCount} spreadsheet attachment(s). ${parseErrors.join(" ")}`,
    );
  }

  if (parsedMonthly.size > 0 || parsedRegion.size > 0) {
    await upsertLotteryData(env.DB, report.reportDate, {
      monthly: [...parsedMonthly.values()],
      region: [...parsedRegion.values()],
    });
  }

  return storedAttachments;
}

export async function crawlLatestReport(): Promise<LatestReport> {
  return fetchReportDetails(await crawlLatestLink());
}

async function crawlLatestLink(): Promise<ArticleLink> {
  return (await crawlReportLinks())[0];
}

async function crawlReportLinks(): Promise<ArticleLink[]> {
  const sourceUrls = [
    MOBILE_REPORT_INDEX_ENDPOINT,
    REPORT_INDEX_ENDPOINT,
    buildSearchUrl(),
  ];
  const errors: string[] = [];

  for (const sourceUrl of sourceUrls) {
    try {
      const html = await fetchText(sourceUrl);
      const links = parseSearchLinks(html, sourceUrl);
      if (links.length > 0) {
        return [...links].sort((left: ArticleLink, right: ArticleLink) =>
          right.reportDate.localeCompare(left.reportDate),
        );
      }
      errors.push(`${sourceUrl}: no matching report links`);
    } catch (error) {
      errors.push(`${sourceUrl}: ${errorMessage(error)}`);
    }
  }

  throw new Error(`No matching lottery report links found. ${errors.join(" ")}`);
}

async function fetchReportDetails(latestLink: ArticleLink): Promise<LatestReport> {
  const searchUrl = buildSearchUrl();
  const articleHtml = await fetchText(latestLink.url, {
    headers: { Referer: searchUrl },
  });
  const pageTitle = sanitizeTitle(
    extractPageTitle(articleHtml) || latestLink.title,
  );
  const attachments = gatherAttachments(articleHtml, latestLink.url);

  return {
    checkedAt: new Date().toISOString(),
    title: latestLink.title,
    url: latestLink.url,
    reportDate: latestLink.reportDate,
    pageTitle,
    attachments,
  };
}

function buildSearchUrl(): string {
  const url = new URL(SEARCH_ENDPOINT);
  for (const [key, value] of Object.entries(SEARCH_PARAMS)) {
    url.searchParams.set(key, value);
  }
  return url.toString();
}

function parseSearchLinks(
  html: string,
  baseUrl = SEARCH_ENDPOINT,
): ArticleLink[] {
  const links: ArticleLink[] = [];
  const seenUrls = new Set<string>();
  const anchorPattern = /<a\b[^>]*href=(["'])(.*?)\1[^>]*>([\s\S]*?)<\/a>/gi;

  for (const match of html.matchAll(anchorPattern)) {
    const href = decodeHtml(match[2]).trim();
    const title = cleanText(stripTags(match[3]));
    if (!title.includes("彩票")) {
      continue;
    }

    const reportDate = extractReportDate(title);
    if (!reportDate) {
      continue;
    }

    const articleUrl = normalizeMofUrl(href, baseUrl);
    if (!articleUrl || seenUrls.has(articleUrl)) {
      continue;
    }

    links.push({ url: articleUrl, title, reportDate });
    seenUrls.add(articleUrl);
  }

  return links;
}

function extractPageTitle(html: string): string | undefined {
  const metaTitle = html.match(
    /<meta\b[^>]*name=(["'])ArticleTitle\1[^>]*content=(["'])(.*?)\2[^>]*>/i,
  );
  if (metaTitle?.[3]) {
    return cleanText(decodeHtml(metaTitle[3]));
  }

  const title = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  if (title?.[1]) {
    return cleanText(stripTags(title[1]));
  }

  const heading = html.match(/<h[12][^>]*>([\s\S]*?)<\/h[12]>/i);
  return heading?.[1] ? cleanText(stripTags(heading[1])) : undefined;
}

function gatherAttachments(html: string, baseUrl: string): Attachment[] {
  const attachments: Attachment[] = [];
  const seenUrls = new Set<string>();
  const anchorPattern = /<a\b[^>]*href=(["'])(.*?)\1[^>]*>([\s\S]*?)<\/a>/gi;

  for (const match of html.matchAll(anchorPattern)) {
    const href = decodeHtml(match[2]).trim();
    const fullUrl = normalizeMofUrl(href, baseUrl);
    if (!fullUrl || seenUrls.has(fullUrl)) {
      continue;
    }

    const linkText = cleanText(stripTags(match[3]));
    const extension = inferExtension(fullUrl, linkText);
    if (!extension) {
      continue;
    }

    attachments.push({
      url: fullUrl,
      originalName: linkText || basenameFromUrl(fullUrl),
      extension,
    });
    seenUrls.add(fullUrl);
  }

  return attachments;
}

async function fetchText(
  url: string,
  init: RequestInit = {},
): Promise<string> {
  const response = await fetchWithRetry(url, init);

  if (!response.ok) {
    throw new Error(`Fetch failed ${response.status} ${response.statusText}: ${url}`);
  }

  const bytes = await response.arrayBuffer();
  const contentType = response.headers.get("content-type") || "";
  const firstPass = decodeBytes(bytes, charsetFromContentType(contentType));
  const metaCharset = charsetFromHtml(firstPass);
  if (metaCharset && metaCharset !== charsetFromContentType(contentType)) {
    return decodeBytes(bytes, metaCharset);
  }
  return firstPass;
}

async function fetchWithRetry(
  url: string,
  init: RequestInit = {},
  attempts = 3,
): Promise<Response> {
  let lastError: unknown;

  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      const response = await fetch(url, {
        ...init,
        headers: {
          ...DEFAULT_HEADERS,
          ...init.headers,
        },
      });

      if (response.ok || response.status < 500 || attempt === attempts) {
        return response;
      }

      lastError = new Error(
        `Fetch failed ${response.status} ${response.statusText}: ${url}`,
      );
    } catch (error) {
      lastError = error;
      if (attempt === attempts) {
        throw error;
      }
    }

    await sleep(500 * attempt);
  }

  throw lastError instanceof Error ? lastError : new Error(String(lastError));
}

function sleep(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

async function storeAttachment(
  bucket: R2Bucket,
  report: LatestReport,
  attachment: Attachment,
  order: number,
): Promise<StoredAttachmentResult> {
  const response = await fetchWithRetry(attachment.url, {
    headers: {
      Referer: report.url,
    },
  });

  if (!response.ok || !response.body) {
    throw new Error(
      `Attachment fetch failed ${response.status} ${response.statusText}: ${attachment.url}`,
    );
  }

  const content = isSpreadsheetAttachment(attachment)
    ? await response.clone().arrayBuffer()
    : undefined;
  const filename = `${report.pageTitle}_${order}${attachment.extension}`;
  const r2Key = `downloads/${filename}`;
  await bucket.put(r2Key, response.body, {
    httpMetadata: {
      contentType:
        response.headers.get("content-type") ||
        contentTypeForExtension(attachment.extension),
    },
    customMetadata: {
      reportDate: report.reportDate,
      sourceUrl: report.url,
      attachmentUrl: attachment.url,
    },
  });

  const stored = await bucket.head(r2Key);
  return {
    stored: {
      filename,
      r2Key,
      fileSize: stored?.size ?? Number(response.headers.get("content-length") || 0),
    },
    content,
  };
}

function isSpreadsheetAttachment(attachment: Attachment): boolean {
  return attachment.extension === ".xls" || attachment.extension === ".xlsx";
}

function parseLotteryWorkbook(bytes: ArrayBuffer): ParsedLotteryData {
  const workbook = XLSX.read(bytes, {
    type: "array",
    cellDates: false,
    cellText: false,
  });
  const monthlySheetName = workbook.SheetNames.find((name) =>
    name.includes("类型"),
  );
  const regionSheetName = workbook.SheetNames.find((name) =>
    name.includes("地区"),
  );

  return {
    monthly: monthlySheetName
      ? parseMonthlySheet(workbook.Sheets[monthlySheetName])
      : [],
    region: regionSheetName
      ? parseRegionSheet(workbook.Sheets[regionSheetName])
      : [],
  };
}

function parseMonthlySheet(sheet: XLSX.WorkSheet): ParsedMonthlyRecord[] {
  const rows = sheetRows(sheet);
  const records: ParsedMonthlyRecord[] = [];
  let lotteryGroup: string | null = null;

  for (const row of rows) {
    const label = cellText(row[0]);
    if (label.includes("福利彩票")) {
      lotteryGroup = "福彩";
      continue;
    }
    if (label.includes("体育彩票")) {
      lotteryGroup = "体彩";
      continue;
    }
    if (label.includes("合计")) {
      lotteryGroup = null;
      continue;
    }
    if (!lotteryGroup) {
      continue;
    }

    const typeMatch = label.match(/^[（(][^）)]*[）)]\s*(.+)$/u);
    const monthlySales = numericCell(row[1]);
    if (!typeMatch?.[1] || monthlySales === null) {
      continue;
    }

    records.push({
      lotteryGroup,
      lotteryType: cleanText(typeMatch[1]),
      monthlySales: roundNumber(monthlySales * BILLION_YUAN_TO_WAN, 6),
    });
  }

  return records;
}

function parseRegionSheet(sheet: XLSX.WorkSheet): ParsedRegionRecord[] {
  const rows = sheetRows(sheet);
  const records: ParsedRegionRecord[] = [];

  for (const row of rows.slice(7)) {
    const regionName = cellText(row[0]);
    if (!regionName) {
      continue;
    }

    const welfareSales = numericCell(row[1]);
    if (welfareSales !== null) {
      records.push({
        regionName,
        lotteryGroup: "福彩",
        monthlySales: roundNumber(welfareSales, 6),
      });
    }

    const sportsSales = numericCell(row[5]);
    if (sportsSales !== null) {
      records.push({
        regionName,
        lotteryGroup: "体彩",
        monthlySales: roundNumber(sportsSales, 6),
      });
    }
  }

  return records;
}

function sheetRows(sheet: XLSX.WorkSheet): unknown[][] {
  return XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: null,
    raw: true,
  });
}

function cellText(value: unknown): string {
  return typeof value === "string" ? cleanText(value) : "";
}

function numericCell(value: unknown): number | null {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value !== "string" || !value.trim()) {
    return null;
  }

  const parsed = Number(value.replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : null;
}

async function upsertLotteryData(
  db: D1Database,
  reportDate: string,
  data: ParsedLotteryData,
): Promise<void> {
  const statements: D1PreparedStatement[] = [];

  for (const row of data.monthly) {
    statements.push(
      db
        .prepare(
          `
          INSERT INTO lottery_monthly_summary (
            report_date,
            lottery_group,
            lottery_type,
            monthly_sales,
            created_at,
            updated_at
          )
          VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
          ON CONFLICT (report_date, lottery_group, lottery_type)
          DO UPDATE SET
            monthly_sales = excluded.monthly_sales,
            updated_at = CURRENT_TIMESTAMP
          `,
        )
        .bind(
          reportDate,
          row.lotteryGroup,
          row.lotteryType,
          row.monthlySales,
        ),
    );
  }

  for (const row of data.region) {
    statements.push(
      db
        .prepare(
          `
          INSERT INTO lottery_region_summary (
            report_date,
            region_name,
            lottery_group,
            monthly_sales,
            created_at,
            updated_at
          )
          VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
          ON CONFLICT (report_date, region_name, lottery_group)
          DO UPDATE SET
            monthly_sales = excluded.monthly_sales,
            updated_at = CURRENT_TIMESTAMP
          `,
        )
        .bind(
          reportDate,
          row.regionName,
          row.lotteryGroup,
          row.monthlySales,
        ),
    );
  }

  if (statements.length > 0) {
    await db.batch(statements);
  }
}

async function getLatestStoredReportDate(db: D1Database): Promise<string | null> {
  const summaryLatest = await db
    .prepare("SELECT MAX(report_date) AS report_date FROM lottery_monthly_summary")
    .first<{ report_date: string | null }>();
  return summaryLatest?.report_date ?? null;
}

async function upsertAttachmentMetadata(
  db: D1Database,
  report: LatestReport,
  attachment: Attachment,
  storedAttachment: StoredAttachment,
): Promise<void> {
  await db
    .prepare(
      `
      INSERT INTO report_attachments (
        report_date,
        page_title,
        source_url,
        attachment_url,
        original_name,
        filename,
        r2_key,
        extension,
        file_size,
        created_at,
        updated_at
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
      ON CONFLICT (report_date, r2_key)
      DO UPDATE SET
        page_title = excluded.page_title,
        source_url = excluded.source_url,
        attachment_url = excluded.attachment_url,
        original_name = excluded.original_name,
        filename = excluded.filename,
        extension = excluded.extension,
        file_size = excluded.file_size,
        updated_at = CURRENT_TIMESTAMP
      `,
    )
    .bind(
      report.reportDate,
      report.pageTitle,
      report.url,
      attachment.url,
      attachment.originalName,
      storedAttachment.filename,
      storedAttachment.r2Key,
      attachment.extension,
      storedAttachment.fileSize,
    )
    .run();
}

async function getCloudState(env: Env): Promise<DatabaseState> {
  const [monthly, region, attachments] = await Promise.all([
    getTableState(env.DB, "lottery_monthly_summary"),
    getTableState(env.DB, "lottery_region_summary"),
    getTableState(env.DB, "report_attachments"),
  ]);
  const r2 = await getR2State(env.DOWNLOADS);

  return { monthly, region, attachments, r2 };
}

async function getUsage(env: Env): Promise<UsageResponse> {
  const [d1StorageBytes, r2] = await Promise.all([
    getD1StorageBytes(env.DB),
    getR2State(env.DOWNLOADS),
  ]);

  const quotas: UsageQuota[] = [
    usageQuota({
      key: "workers.requests",
      resource: "Workers",
      name: "Workers 请求数",
      period: "daily",
      unit: "requests",
      freeLimit: FREE_LIMITS.workersRequestsPerDay,
      currentUsage: null,
      source: "cloudflare_analytics_required",
      notes: "Worker 绑定无法读取账号当天真实请求数；需在 Cloudflare Dashboard 或 Analytics API 查看。",
    }),
    usageQuota({
      key: "workers.cpu_time",
      resource: "Workers",
      name: "Workers CPU 时间",
      period: "per_invocation",
      unit: "ms",
      freeLimit: FREE_LIMITS.workersCpuMsPerInvocation,
      currentUsage: null,
      source: "cloudflare_analytics_required",
      notes: "单次调用的实时 CPU 用量不从业务接口统计；可在 Workers Metrics 或 Tail Logs 查看。",
    }),
    usageQuota({
      key: "workers.memory",
      resource: "Workers",
      name: "Workers 内存",
      period: "per_invocation",
      unit: "bytes",
      freeLimit: FREE_LIMITS.workersMemoryBytes,
      currentUsage: null,
      source: "cloudflare_analytics_required",
      notes: "内存超限在 Workers Metrics 中体现；业务接口无法读到账号级运行时峰值。",
    }),
    usageQuota({
      key: "workers.external_subrequests",
      resource: "Workers",
      name: "外部 subrequests",
      period: "per_invocation",
      unit: "subrequests",
      freeLimit: FREE_LIMITS.workersExternalSubrequestsPerInvocation,
      currentUsage: null,
      source: "worker_runtime",
      notes: "该项目的同步任务通常会请求财政部搜索页、文章页和新增附件；真实次数随远端响应和附件数量变化。",
    }),
    usageQuota({
      key: "workers.internal_subrequests",
      resource: "Workers",
      name: "内部服务 subrequests",
      period: "per_invocation",
      unit: "subrequests",
      freeLimit: FREE_LIMITS.workersInternalSubrequestsPerInvocation,
      currentUsage: null,
      source: "worker_runtime",
      notes: "D1/R2 调用属于内部服务 subrequests；接口无法读取当前调用累计值。",
    }),
    usageQuota({
      key: "workers.count",
      resource: "Workers",
      name: "账号 Workers 数量",
      period: "account",
      unit: "workers",
      freeLimit: FREE_LIMITS.workersPerAccount,
      currentUsage: 1,
      source: "worker_config",
    }),
    usageQuota({
      key: "cron.triggers",
      resource: "Cron",
      name: "账号 Cron Triggers 数量",
      period: "account",
      unit: "triggers",
      freeLimit: FREE_LIMITS.cronTriggersPerAccount,
      currentUsage: 1,
      source: "worker_config",
    }),
    usageQuota({
      key: "custom_domains.count",
      resource: "Custom Domain",
      name: "zone Custom Domains 数量",
      period: "zone",
      unit: "domains",
      freeLimit: FREE_LIMITS.customDomainsPerZone,
      currentUsage: 1,
      source: "worker_config",
    }),
    usageQuota({
      key: "routes.count",
      resource: "Custom Domain",
      name: "zone Routes 数量",
      period: "zone",
      unit: "routes",
      freeLimit: FREE_LIMITS.routesPerZone,
      currentUsage: 1,
      source: "worker_config",
      notes: "当前配置使用 1 个 Custom Domain route。",
    }),
    usageQuota({
      key: "d1.rows_read",
      resource: "D1",
      name: "D1 rows read",
      period: "daily",
      unit: "rows",
      freeLimit: FREE_LIMITS.d1RowsReadPerDay,
      currentUsage: null,
      source: "cloudflare_analytics_required",
      notes: "D1 绑定无法读取今日累计 rows read；Dashboard、Wrangler 查询也会计入用量。",
    }),
    usageQuota({
      key: "d1.rows_written",
      resource: "D1",
      name: "D1 rows written",
      period: "daily",
      unit: "rows",
      freeLimit: FREE_LIMITS.d1RowsWrittenPerDay,
      currentUsage: null,
      source: "cloudflare_analytics_required",
      notes: "D1 绑定无法读取今日累计 rows written；新增报告和 migration 都会计入写入。",
    }),
    usageQuota({
      key: "d1.storage",
      resource: "D1",
      name: "D1 存储",
      period: "total",
      unit: "bytes",
      freeLimit: FREE_LIMITS.d1StorageBytes,
      currentUsage: d1StorageBytes,
      source: "d1_meta",
      notes: "使用 D1 查询 meta.size_after 估算当前数据库大小。",
    }),
    usageQuota({
      key: "r2.storage",
      resource: "R2",
      name: "R2 Standard 存储",
      period: "monthly",
      unit: "bytes-month",
      freeLimit: FREE_LIMITS.r2StorageBytesMonth,
      currentUsage: r2.totalSizeBytes,
      source: "r2_list",
      notes: "这里用当前对象总大小近似对比 10 GB-month；实际账单按月内每日峰值平均计算。",
    }),
    usageQuota({
      key: "r2.class_a_operations",
      resource: "R2",
      name: "R2 Class A 操作",
      period: "monthly",
      unit: "operations",
      freeLimit: FREE_LIMITS.r2ClassAOperationsPerMonth,
      currentUsage: null,
      source: "cloudflare_analytics_required",
      notes: "例如 PutObject、ListObjects；Worker 绑定无法读取本月累计操作数。",
    }),
    usageQuota({
      key: "r2.class_b_operations",
      resource: "R2",
      name: "R2 Class B 操作",
      period: "monthly",
      unit: "operations",
      freeLimit: FREE_LIMITS.r2ClassBOperationsPerMonth,
      currentUsage: null,
      source: "cloudflare_analytics_required",
      notes: "例如 GetObject、HeadObject；Worker 绑定无法读取本月累计操作数。",
    }),
  ];

  return {
    checkedAt: new Date().toISOString(),
    baseUrl: "https://lottery.aback.fun",
    note:
      "该接口返回项目可直接计算的免费额度占比；Cloudflare 账号级真实用量需通过 Dashboard 或 Analytics API 查看。",
    periodSummary: summarizeQuotaPeriods(quotas),
    resources: {
      worker: {
        name: "lottery-latest-crawler",
        customDomain: "lottery.aback.fun",
        cron: "0 2 * * *",
      },
      d1: {
        databaseName: "lottery-data",
        databaseId: "900a0db3-19df-4d0f-9690-8ce5951035a1",
        estimatedStorageBytes: d1StorageBytes,
        estimatedStorageHuman: d1StorageBytes === null ? null : formatBytes(d1StorageBytes),
      },
      r2,
    },
    quotas,
  };
}

async function getAvailableMonths(
  db: D1Database,
  url: URL,
): Promise<ApiRowsResponse<AvailableMonthRow>> {
  const page = getPageRequest(url);
  const total = await getAvailableMonthCount(db);
  const result = await db
    .prepare(
      `
      SELECT
        report_date AS reportDate,
        SUM(monthly_rows) AS monthlyRows,
        SUM(region_rows) AS regionRows,
        SUM(attachment_rows) AS attachmentRows
      FROM (
        SELECT report_date, COUNT(*) AS monthly_rows, 0 AS region_rows, 0 AS attachment_rows
        FROM lottery_monthly_summary
        GROUP BY report_date
        UNION ALL
        SELECT report_date, 0 AS monthly_rows, COUNT(*) AS region_rows, 0 AS attachment_rows
        FROM lottery_region_summary
        GROUP BY report_date
        UNION ALL
        SELECT report_date, 0 AS monthly_rows, 0 AS region_rows, COUNT(*) AS attachment_rows
        FROM report_attachments
        GROUP BY report_date
      ) AS grouped_months
      GROUP BY report_date
      ORDER BY report_date DESC
      LIMIT ? OFFSET ?
      `,
    )
    .bind(page.limit, page.offset)
    .all<AvailableMonthRow>();

  return rowsResponse(page, total, result.results ?? []);
}

async function getMonthlySummary(
  db: D1Database,
  url: URL,
): Promise<ApiRowsResponse<MonthlySummaryRow>> {
  const page = getPageRequest(url);
  const date = await resolveReportDate(
    db,
    "lottery_monthly_summary",
    getReportDateParam(url),
  );
  const lotteryGroup = textParam(url, "lottery_group", "lotteryGroup");
  const lotteryType = textParam(url, "lottery_type", "lotteryType");

  if (!date) {
    return rowsResponse(page, 0, [], null);
  }

  const { whereClause, values } = buildWhereClause([
    ["report_date", date],
    ["lottery_group", lotteryGroup],
    ["lottery_type", lotteryType],
  ]);
  const total = await getFilteredCount(db, "lottery_monthly_summary", whereClause, values);
  const result = await db
    .prepare(
      `
      SELECT
        report_date AS reportDate,
        lottery_group AS lotteryGroup,
        lottery_type AS lotteryType,
        monthly_sales AS monthlySales
      FROM lottery_monthly_summary
      WHERE ${whereClause}
      ORDER BY lottery_group, lottery_type
      LIMIT ? OFFSET ?
      `,
    )
    .bind(...values, page.limit, page.offset)
    .all<MonthlySummaryRow>();

  return rowsResponse(page, total, result.results ?? [], date, {
    lotteryGroup,
    lotteryType,
  });
}

async function getRegionSummary(
  db: D1Database,
  url: URL,
): Promise<ApiRowsResponse<RegionSummaryRow>> {
  const page = getPageRequest(url);
  const date = await resolveReportDate(
    db,
    "lottery_region_summary",
    getReportDateParam(url),
  );
  const region = textParam(url, "region", "region_name", "regionName");
  const lotteryGroup = textParam(url, "lottery_group", "lotteryGroup");

  if (!date) {
    return rowsResponse(page, 0, [], null);
  }

  const { whereClause, values } = buildWhereClause([
    ["report_date", date],
    ["region_name", region],
    ["lottery_group", lotteryGroup],
  ]);
  const total = await getFilteredCount(db, "lottery_region_summary", whereClause, values);
  const result = await db
    .prepare(
      `
      SELECT
        report_date AS reportDate,
        region_name AS regionName,
        lottery_group AS lotteryGroup,
        monthly_sales AS monthlySales
      FROM lottery_region_summary
      WHERE ${whereClause}
      ORDER BY region_name, lottery_group
      LIMIT ? OFFSET ?
      `,
    )
    .bind(...values, page.limit, page.offset)
    .all<RegionSummaryRow>();

  return rowsResponse(page, total, result.results ?? [], date, {
    region,
    lotteryGroup,
  });
}

async function getAttachments(
  db: D1Database,
  url: URL,
): Promise<ApiRowsResponse<AttachmentRow>> {
  const page = getPageRequest(url);
  const date = await resolveReportDate(
    db,
    "report_attachments",
    getReportDateParam(url),
  );
  const extension = normalizeExtension(textParam(url, "extension", "ext"));

  if (!date) {
    return rowsResponse(page, 0, [], null);
  }

  const { whereClause, values } = buildWhereClause([
    ["report_date", date],
    ["extension", extension],
  ]);
  const total = await getFilteredCount(db, "report_attachments", whereClause, values);
  const result = await db
    .prepare(
      `
      SELECT
        report_date AS reportDate,
        page_title AS pageTitle,
        source_url AS sourceUrl,
        attachment_url AS attachmentUrl,
        original_name AS originalName,
        filename,
        r2_key AS r2Key,
        extension,
        file_size AS fileSize,
        created_at AS createdAt,
        updated_at AS updatedAt
      FROM report_attachments
      WHERE ${whereClause}
      ORDER BY report_date DESC, filename
      LIMIT ? OFFSET ?
      `,
    )
    .bind(...values, page.limit, page.offset)
    .all<AttachmentRow>();

  return rowsResponse(page, total, result.results ?? [], date, { extension });
}

async function getDashboardSnapshot(db: D1Database): Promise<DashboardSnapshot> {
  const cached = await readDashboardSnapshot(db);
  if (cached && !dashboardSnapshotNeedsRefresh(cached)) {
    return cached;
  }
  return refreshDashboardSnapshot(db);
}

function dashboardSnapshotNeedsRefresh(snapshot: DashboardSnapshot): boolean {
  return snapshot.typeTrend.length > 0 && snapshot.typeTrend.some((point) => !point.lotteryGroup);
}

async function refreshDashboardSnapshot(db: D1Database): Promise<DashboardSnapshot> {
  const snapshot = await buildDashboardSnapshot(db);
  await db
    .prepare(
      `
      INSERT INTO dashboard_cache (
        cache_key,
        report_date,
        payload,
        created_at,
        updated_at
      )
      VALUES (?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
      ON CONFLICT (cache_key)
      DO UPDATE SET
        report_date = excluded.report_date,
        payload = excluded.payload,
        updated_at = CURRENT_TIMESTAMP
      `,
    )
    .bind(DASHBOARD_CACHE_KEY, snapshot.latest.reportDate, JSON.stringify(snapshot))
    .run();
  return snapshot;
}

async function readDashboardSnapshot(db: D1Database): Promise<DashboardSnapshot | null> {
  const row = await db
    .prepare(
      "SELECT report_date, payload, updated_at FROM dashboard_cache WHERE cache_key = ?",
    )
    .bind(DASHBOARD_CACHE_KEY)
    .first<DashboardCacheRow>();

  if (!row?.payload) {
    return null;
  }

  try {
    return JSON.parse(row.payload) as DashboardSnapshot;
  } catch (error) {
    console.error("dashboard snapshot parse failed", serializeError(error));
    return null;
  }
}

async function buildDashboardSnapshot(db: D1Database): Promise<DashboardSnapshot> {
  const [monthly, region, attachments, monthlyResult, regionResult] =
    await Promise.all([
      getTableState(db, "lottery_monthly_summary"),
      getTableState(db, "lottery_region_summary"),
      getTableState(db, "report_attachments"),
      db
        .prepare(
          `
          SELECT
            report_date AS reportDate,
            lottery_group AS lotteryGroup,
            lottery_type AS lotteryType,
            monthly_sales AS monthlySales
          FROM lottery_monthly_summary
          ORDER BY report_date ASC, lottery_group, lottery_type
          `,
        )
        .all<MonthlySummaryRow>(),
      db
        .prepare(
          `
          SELECT
            report_date AS reportDate,
            region_name AS regionName,
            lottery_group AS lotteryGroup,
            monthly_sales AS monthlySales
          FROM lottery_region_summary
          ORDER BY report_date ASC, region_name, lottery_group
          `,
        )
        .all<RegionSummaryRow>(),
    ]);

  const monthlyRows = (monthlyResult.results ?? []).map((row) => ({
    ...row,
    monthlySales: numberOrZero(row.monthlySales),
  }));
  const regionRows = (regionResult.results ?? []).map((row) => ({
    ...row,
    monthlySales: numberOrZero(row.monthlySales),
  }));

  const monthlyTrend = buildMonthlyTrend(monthlyRows);
  const annualTrend = buildAnnualTrend(monthlyTrend);
  const groupTrend = buildGroupTrend(monthlyRows);
  const typeTrend = buildTypeTrend(monthlyRows);
  const latestTrend = monthlyTrend.at(-1);
  const latestReportDate = latestTrend?.reportDate ?? null;
  const latestGroupBreakdown = buildLatestGroupBreakdown(
    groupTrend,
    latestReportDate,
    latestTrend?.totalSales ?? 0,
  );
  const latestTypeBreakdown = buildLatestTypeBreakdown(
    monthlyRows,
    latestReportDate,
    latestTrend?.totalSales ?? 0,
  );
  const topType =
    latestTypeBreakdown.length === 0
      ? null
      : latestTypeBreakdown.reduce((top, row) =>
          row.monthlySales > top.monthlySales ? row : top,
        );
  const regionSnapshot = buildRegionSnapshot(regionRows);

  return {
    generatedAt: new Date().toISOString(),
    source: {
      unit: "万元",
      salesDisplayUnit: "亿元",
      note: "仪表盘数据来自 D1 预计算快照；销售额原始单位按财政部附件口径保存为万元，前端展示时换算为亿元。上年同月销售额低于 1 亿元时，同比视为低基数不可比，不参与图表坐标轴。",
    },
    coverage: {
      monthly,
      region,
      attachments,
    },
    latest: {
      reportDate: latestReportDate,
      totalSales: latestTrend?.totalSales ?? 0,
      welfareSales: latestTrend?.welfareSales ?? 0,
      sportsSales: latestTrend?.sportsSales ?? 0,
      monthOverMonthPercent: latestTrend?.monthOverMonthPercent ?? null,
      yearOverYearPercent: latestTrend?.yearOverYearPercent ?? null,
      topType:
        topType && typeof topType !== "string"
          ? `${topType.lotteryGroup}${topType.lotteryType}`
          : null,
      topTypeSales: topType && typeof topType !== "string" ? topType.monthlySales : 0,
      topTypeSharePercent:
        topType && typeof topType !== "string" ? topType.sharePercent : 0,
      regionReportDate: regionSnapshot.regionLatest.reportDate,
    },
    monthlyTrend,
    annualTrend,
    groupTrend,
    typeTrend,
    latestGroupBreakdown,
    latestTypeBreakdown,
    growthBars: monthlyTrend.slice(-36).map((point) => ({
      reportDate: point.reportDate,
      monthOverMonthPercent: point.monthOverMonthPercent,
      yearOverYearPercent: point.yearOverYearPercent,
    })),
    regionLatest: regionSnapshot.regionLatest,
    regionHeatmap: regionSnapshot.regionHeatmap,
  };
}

function buildMonthlyTrend(rows: Array<MonthlySummaryRow & { monthlySales: number }>): DashboardTrendPoint[] {
  const byDate = new Map<string, DashboardTrendPoint>();

  for (const row of rows) {
    const point = getOrCreate(byDate, row.reportDate, () => ({
      reportDate: row.reportDate,
      totalSales: 0,
      welfareSales: 0,
      sportsSales: 0,
      monthOverMonthPercent: null,
      yearOverYearPercent: null,
    }));
    point.totalSales += row.monthlySales;
    if (row.lotteryGroup === "福彩") {
      point.welfareSales += row.monthlySales;
    } else if (row.lotteryGroup === "体彩") {
      point.sportsSales += row.monthlySales;
    }
  }

  const trend = [...byDate.values()]
    .sort((left, right) => left.reportDate.localeCompare(right.reportDate))
    .map((point) => ({
      ...point,
      totalSales: roundNumber(point.totalSales),
      welfareSales: roundNumber(point.welfareSales),
      sportsSales: roundNumber(point.sportsSales),
    }));
  const trendByDate = new Map(trend.map((point) => [point.reportDate, point]));

  for (const [index, point] of trend.entries()) {
    point.monthOverMonthPercent =
      index === 0 ? null : percentChange(point.totalSales, trend[index - 1].totalSales);
    const lastYear = trendByDate.get(shiftReportDate(point.reportDate, -12));
    point.yearOverYearPercent = lastYear && lastYear.totalSales >= YOY_MINIMUM_BASE_SALES
      ? percentChange(point.totalSales, lastYear.totalSales)
      : null;
  }

  return trend;
}

function buildAnnualTrend(monthlyTrend: DashboardTrendPoint[]): DashboardAnnualPoint[] {
  const byYear = new Map<string, DashboardAnnualPoint>();

  for (const point of monthlyTrend) {
    const year = point.reportDate.slice(0, 4);
    const annual = getOrCreate(byYear, year, () => ({
      year,
      totalSales: 0,
      welfareSales: 0,
      sportsSales: 0,
      welfareSharePercent: 0,
      sportsSharePercent: 0,
      monthCount: 0,
      isPartial: false,
    }));
    annual.totalSales += point.totalSales;
    annual.welfareSales += point.welfareSales;
    annual.sportsSales += point.sportsSales;
    annual.monthCount += 1;
  }

  return [...byYear.values()]
    .sort((left, right) => left.year.localeCompare(right.year))
    .map((point) => {
      const totalSales = roundNumber(point.totalSales);
      const welfareSales = roundNumber(point.welfareSales);
      const sportsSales = roundNumber(point.sportsSales);
      return {
        ...point,
        totalSales,
        welfareSales,
        sportsSales,
        welfareSharePercent: sharePercent(welfareSales, totalSales),
        sportsSharePercent: sharePercent(sportsSales, totalSales),
        isPartial: point.monthCount < 12,
      };
    });
}

function buildGroupTrend(rows: Array<MonthlySummaryRow & { monthlySales: number }>): DashboardGroupPoint[] {
  const byDateAndGroup = new Map<string, DashboardGroupPoint>();

  for (const row of rows) {
    const key = `${row.reportDate}|${row.lotteryGroup}`;
    const point = getOrCreate(byDateAndGroup, key, () => ({
      reportDate: row.reportDate,
      lotteryGroup: row.lotteryGroup,
      monthlySales: 0,
    }));
    point.monthlySales += row.monthlySales;
  }

  return [...byDateAndGroup.values()]
    .sort(
      (left, right) =>
        left.reportDate.localeCompare(right.reportDate) ||
        left.lotteryGroup.localeCompare(right.lotteryGroup),
    )
    .map((point) => ({ ...point, monthlySales: roundNumber(point.monthlySales) }));
}

function buildTypeTrend(rows: Array<MonthlySummaryRow & { monthlySales: number }>): DashboardTypePoint[] {
  const byDateAndType = new Map<string, DashboardTypePoint>();

  for (const row of rows) {
    const key = `${row.reportDate}|${row.lotteryGroup}|${row.lotteryType}`;
    const point = getOrCreate(byDateAndType, key, () => ({
      reportDate: row.reportDate,
      lotteryGroup: row.lotteryGroup,
      lotteryType: row.lotteryType,
      monthlySales: 0,
    }));
    point.monthlySales += row.monthlySales;
  }

  return [...byDateAndType.values()]
    .sort(
      (left, right) =>
        left.reportDate.localeCompare(right.reportDate) ||
        left.lotteryGroup.localeCompare(right.lotteryGroup) ||
        left.lotteryType.localeCompare(right.lotteryType),
    )
    .map((point) => ({ ...point, monthlySales: roundNumber(point.monthlySales) }));
}

function buildLatestGroupBreakdown(
  groupTrend: DashboardGroupPoint[],
  latestReportDate: string | null,
  totalSales: number,
): DashboardBreakdownPoint[] {
  if (!latestReportDate) {
    return [];
  }

  return groupTrend
    .filter((point) => point.reportDate === latestReportDate)
    .map((point) => ({
      name: point.lotteryGroup,
      monthlySales: point.monthlySales,
      sharePercent: sharePercent(point.monthlySales, totalSales),
    }))
    .sort((left, right) => right.monthlySales - left.monthlySales);
}

function buildLatestTypeBreakdown(
  rows: Array<MonthlySummaryRow & { monthlySales: number }>,
  latestReportDate: string | null,
  totalSales: number,
): DashboardTypeBreakdownPoint[] {
  if (!latestReportDate) {
    return [];
  }

  return rows
    .filter((row) => row.reportDate === latestReportDate)
    .map((row) => ({
      lotteryGroup: row.lotteryGroup,
      lotteryType: row.lotteryType,
      monthlySales: roundNumber(row.monthlySales),
      sharePercent: sharePercent(row.monthlySales, totalSales),
    }))
    .sort((left, right) => right.monthlySales - left.monthlySales);
}

function buildRegionSnapshot(rows: Array<RegionSummaryRow & { monthlySales: number }>): {
  regionLatest: DashboardSnapshot["regionLatest"];
  regionHeatmap: DashboardSnapshot["regionHeatmap"];
} {
  const regionDates = [
    ...new Set(rows.filter((row) => row.regionName !== "总计").map((row) => row.reportDate)),
  ].sort();
  const latestRegionDate = regionDates.at(-1) ?? null;
  const latestByRegion = new Map<string, DashboardRegionPoint>();

  for (const row of rows) {
    if (row.reportDate !== latestRegionDate || row.regionName === "总计") {
      continue;
    }
    const point = getOrCreate(latestByRegion, row.regionName, () => ({
      regionName: row.regionName,
      welfareSales: 0,
      sportsSales: 0,
      totalSales: 0,
    }));
    if (row.lotteryGroup === "福彩") {
      point.welfareSales += row.monthlySales;
    } else if (row.lotteryGroup === "体彩") {
      point.sportsSales += row.monthlySales;
    }
    point.totalSales += row.monthlySales;
  }

  const latestRows = [...latestByRegion.values()]
    .map((row) => ({
      ...row,
      welfareSales: roundNumber(row.welfareSales),
      sportsSales: roundNumber(row.sportsSales),
      totalSales: roundNumber(row.totalSales),
    }))
    .sort((left, right) => right.totalSales - left.totalSales);
  const heatmapMonths = regionDates.slice(-24);
  const heatmapRegions = latestRows.slice(0, 20).map((row) => row.regionName);
  const heatmapMonthSet = new Set(heatmapMonths);
  const heatmapRegionSet = new Set(heatmapRegions);
  const totals = new Map<string, number>();

  for (const row of rows) {
    if (
      row.regionName === "总计" ||
      !heatmapMonthSet.has(row.reportDate) ||
      !heatmapRegionSet.has(row.regionName)
    ) {
      continue;
    }
    const key = `${row.reportDate}|${row.regionName}`;
    totals.set(key, (totals.get(key) ?? 0) + row.monthlySales);
  }

  const values: Array<[number, number, number]> = [];
  for (const [monthIndex, month] of heatmapMonths.entries()) {
    for (const [regionIndex, regionName] of heatmapRegions.entries()) {
      values.push([
        monthIndex,
        regionIndex,
        roundNumber(totals.get(`${month}|${regionName}`) ?? 0),
      ]);
    }
  }

  return {
    regionLatest: {
      reportDate: latestRegionDate,
      rows: latestRows,
    },
    regionHeatmap: {
      months: heatmapMonths,
      regions: heatmapRegions,
      values,
    },
  };
}

async function getTableState(
  db: D1Database,
  tableName: string,
): Promise<TableState> {
  const row = await db
    .prepare(
      `SELECT COUNT(*) AS count, MIN(report_date) AS minReportDate, MAX(report_date) AS maxReportDate FROM ${tableName}`,
    )
    .first<{
      count: number;
      minReportDate: string | null;
      maxReportDate: string | null;
    }>();

  return {
    count: row?.count ?? 0,
    minReportDate: row?.minReportDate ?? null,
    maxReportDate: row?.maxReportDate ?? null,
  };
}

async function getD1StorageBytes(db: D1Database): Promise<number | null> {
  const result = await db.prepare("SELECT 1 AS health_check").all();
  return typeof result.meta.size_after === "number" ? result.meta.size_after : null;
}

async function getLatestReportDateForTable(
  db: D1Database,
  tableName: ReportTable,
): Promise<string | null> {
  const row = await db
    .prepare(`SELECT MAX(report_date) AS report_date FROM ${tableName}`)
    .first<{ report_date: string | null }>();
  return row?.report_date ?? null;
}

async function resolveReportDate(
  db: D1Database,
  tableName: ReportTable,
  requestedDate: string | null,
): Promise<string | null> {
  return requestedDate ?? getLatestReportDateForTable(db, tableName);
}

async function getAvailableMonthCount(db: D1Database): Promise<number> {
  const row = await db
    .prepare(
      `
      SELECT COUNT(*) AS total
      FROM (
        SELECT report_date FROM lottery_monthly_summary
        UNION
        SELECT report_date FROM lottery_region_summary
        UNION
        SELECT report_date FROM report_attachments
      ) AS all_report_dates
      `,
    )
    .first<{ total: number }>();
  return Number(row?.total ?? 0);
}

async function getFilteredCount(
  db: D1Database,
  tableName: ReportTable,
  whereClause: string,
  values: QueryValue[],
): Promise<number> {
  const row = await db
    .prepare(`SELECT COUNT(*) AS total FROM ${tableName} WHERE ${whereClause}`)
    .bind(...values)
    .first<{ total: number }>();
  return Number(row?.total ?? 0);
}

function buildWhereClause(
  filters: Array<[column: string, value: string | null]>,
): { whereClause: string; values: QueryValue[] } {
  const clauses: string[] = [];
  const values: QueryValue[] = [];

  for (const [column, value] of filters) {
    if (!value) {
      continue;
    }
    clauses.push(`${column} = ?`);
    values.push(value);
  }

  if (clauses.length === 0) {
    return { whereClause: "1 = 1", values };
  }
  return { whereClause: clauses.join(" AND "), values };
}

function rowsResponse<Row>(
  page: PageRequest,
  total: number,
  rows: Row[],
  date?: string | null,
  rawFilters: Record<string, string | null> = {},
): ApiRowsResponse<Row> {
  const response: ApiRowsResponse<Row> = {
    pagination: {
      limit: page.limit,
      offset: page.offset,
      total,
    },
    rows,
  };

  if (date !== undefined) {
    response.date = date;
  }

  const filters = compactFilters(rawFilters);
  if (Object.keys(filters).length > 0) {
    response.filters = filters;
  }

  return response;
}

function compactFilters(filters: Record<string, string | null>): Record<string, string> {
  const compacted: Record<string, string> = {};
  for (const [key, value] of Object.entries(filters)) {
    if (value) {
      compacted[key] = value;
    }
  }
  return compacted;
}

function getPageRequest(url: URL): PageRequest {
  return {
    limit: integerParam(url, "limit", DEFAULT_PAGE_LIMIT, 1, MAX_PAGE_LIMIT),
    offset: integerParam(url, "offset", 0, 0, MAX_PAGE_OFFSET),
  };
}

function integerParam(
  url: URL,
  name: string,
  defaultValue: number,
  min: number,
  max: number,
): number {
  const rawValue = url.searchParams.get(name);
  if (!rawValue) {
    return defaultValue;
  }

  const value = Number(rawValue);
  if (!Number.isInteger(value) || value < min) {
    throw new HttpError(400, `Invalid ${name}; expected an integer >= ${min}.`);
  }

  return Math.min(value, max);
}

function getReportDateParam(url: URL): string | null {
  const value = textParam(url, "date", "report_date", "reportDate");
  if (!value || value.toLowerCase() === "latest") {
    return null;
  }
  return normalizeReportDate(value);
}

function normalizeReportDate(value: string): string {
  const match = value.match(/^(?<year>\d{4})-(?<month>\d{1,2})(?:-(?<day>\d{1,2}))?$/);
  if (!match?.groups) {
    throw new HttpError(400, "Invalid date; expected YYYY-MM or YYYY-MM-01.");
  }

  const year = Number(match.groups.year);
  const month = Number(match.groups.month);
  const day = match.groups.day ? Number(match.groups.day) : 1;
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day) ||
    month < 1 ||
    month > 12 ||
    day !== 1
  ) {
    throw new HttpError(400, "Invalid date; expected YYYY-MM or YYYY-MM-01.");
  }

  return `${year}-${String(month).padStart(2, "0")}-01`;
}

function textParam(url: URL, ...names: string[]): string | null {
  for (const name of names) {
    const value = url.searchParams.get(name)?.trim();
    if (value) {
      return value;
    }
  }
  return null;
}

function normalizeExtension(value: string | null): string | null {
  if (!value) {
    return null;
  }
  const normalized = value.startsWith(".") ? value.toLowerCase() : `.${value.toLowerCase()}`;
  if (!ATTACHMENT_EXTENSIONS.has(normalized)) {
    throw new HttpError(400, "Invalid extension filter.");
  }
  return normalized;
}

async function getR2State(bucket: R2Bucket): Promise<R2State> {
  let objectCount = 0;
  let totalSizeBytes = 0;
  let cursor: string | undefined;

  do {
    const page = await bucket.list({
      prefix: "downloads/",
      cursor,
      limit: 1000,
    });
    objectCount += page.objects.length;
    totalSizeBytes += page.objects.reduce((sum, object) => sum + object.size, 0);
    cursor = page.truncated ? page.cursor : undefined;
  } while (cursor);

  return {
    prefix: "downloads/",
    objectCount,
    totalSizeBytes,
    totalSizeHuman: formatBytes(totalSizeBytes),
  };
}

function usageQuota(input: Omit<UsageQuota, "remaining" | "percentOfFreeLimit" | "periodLabel" | "status">): UsageQuota {
  const remaining =
    input.currentUsage === null ? null : input.freeLimit - input.currentUsage;
  return {
    ...input,
    periodLabel: periodLabel(input.period),
    remaining,
    percentOfFreeLimit:
      input.currentUsage === null ? null : percentOf(input.currentUsage, input.freeLimit),
    status: input.currentUsage === null ? "unavailable" : "available",
  };
}

function summarizeQuotaPeriods(quotas: UsageQuota[]): Record<QuotaPeriod, string[]> {
  const summary: Record<QuotaPeriod, string[]> = {
    daily: [],
    weekly: [],
    monthly: [],
    per_invocation: [],
    account: [],
    zone: [],
    total: [],
  };

  for (const quota of quotas) {
    summary[quota.period].push(quota.key);
  }

  return summary;
}

function periodLabel(period: QuotaPeriod): string {
  const labels: Record<QuotaPeriod, string> = {
    daily: "日限",
    weekly: "周限",
    monthly: "月限",
    per_invocation: "单次调用限制",
    account: "账号总量限制",
    zone: "Zone 总量限制",
    total: "总量限制",
  };
  return labels[period];
}

function percentOf(currentUsage: number, freeLimit: number): number {
  return Math.round((currentUsage / freeLimit) * 10_000) / 100;
}

function numberOrZero(value: number | null): number {
  const numberValue = Number(value ?? 0);
  return Number.isFinite(numberValue) ? numberValue : 0;
}

function roundNumber(value: number, digits = 4): number {
  const scale = 10 ** digits;
  return Math.round(value * scale) / scale;
}

function percentChange(currentValue: number, previousValue: number): number | null {
  if (!Number.isFinite(currentValue) || !Number.isFinite(previousValue) || previousValue === 0) {
    return null;
  }
  return roundNumber(((currentValue - previousValue) / previousValue) * 100, 2);
}

function sharePercent(value: number, total: number): number {
  if (!Number.isFinite(value) || !Number.isFinite(total) || total === 0) {
    return 0;
  }
  return roundNumber((value / total) * 100, 2);
}

function shiftReportDate(reportDate: string, monthOffset: number): string {
  const [yearText, monthText] = reportDate.split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  const zeroBasedMonth = year * 12 + month - 1 + monthOffset;
  const shiftedYear = Math.floor(zeroBasedMonth / 12);
  const shiftedMonth = zeroBasedMonth - shiftedYear * 12 + 1;
  return `${shiftedYear}-${String(shiftedMonth).padStart(2, "0")}-01`;
}

function getOrCreate<Key, Value>(
  map: Map<Key, Value>,
  key: Key,
  createValue: () => Value,
): Value {
  const existing = map.get(key);
  if (existing) {
    return existing;
  }
  const value = createValue();
  map.set(key, value);
  return value;
}

function formatBytes(bytes: number): string {
  if (bytes < 1000) {
    return `${bytes} B`;
  }

  const units = ["KB", "MB", "GB", "TB"];
  let value = bytes;
  let unitIndex = -1;
  do {
    value /= 1000;
    unitIndex += 1;
  } while (value >= 1000 && unitIndex < units.length - 1);

  return `${Math.round(value * 100) / 100} ${units[unitIndex]}`;
}

function decodeBytes(bytes: ArrayBuffer, charset?: string): string {
  const labels = charset ? [charset, "gb18030", "utf-8"] : ["utf-8", "gb18030"];

  for (const label of labels) {
    try {
      return new TextDecoder(label).decode(bytes);
    } catch {
      // Try the next supported label.
    }
  }

  return new TextDecoder().decode(bytes);
}

function charsetFromContentType(contentType: string): string | undefined {
  const match = contentType.match(/charset=([^;\s]+)/i);
  return normalizeCharset(match?.[1]);
}

function charsetFromHtml(html: string): string | undefined {
  const match = html.match(/<meta\b[^>]*charset=(["']?)([^"'\s/>]+)\1/i);
  return normalizeCharset(match?.[2]);
}

function normalizeCharset(value?: string): string | undefined {
  if (!value) {
    return undefined;
  }
  const normalized = value.trim().toLowerCase();
  if (normalized === "gb2312" || normalized === "gbk") {
    return "gb18030";
  }
  return normalized;
}

function normalizeMofUrl(href: string, baseUrl: string): string | undefined {
  if (!href || /^(javascript:|#|mailto:)/i.test(href)) {
    return undefined;
  }

  const resolved = href.startsWith("//")
    ? new URL(`https:${href}`)
    : new URL(href, baseUrl);

  if (!["http:", "https:"].includes(resolved.protocol)) {
    return undefined;
  }
  if (!resolved.hostname.endsWith("mof.gov.cn")) {
    return undefined;
  }
  if (resolved.pathname.includes("/web/search")) {
    return undefined;
  }

  resolved.protocol = "https:";
  resolved.hash = "";
  return resolved.toString();
}

function inferExtension(url: string, linkText: string): string | undefined {
  const pathname = new URL(url).pathname.toLowerCase();
  const fromPath = extensionFromText(pathname);
  if (fromPath) {
    return fromPath;
  }
  return extensionFromText(linkText.toLowerCase());
}

function extensionFromText(text: string): string | undefined {
  const match = text.match(/(\.pdf|\.docx?|\.xlsx?|\.zip|\.rar)(?:$|[?#\s])/i);
  const extension = match?.[1]?.toLowerCase();
  return extension && ATTACHMENT_EXTENSIONS.has(extension) ? extension : undefined;
}

function extractReportDate(text: string): string | undefined {
  const match = text.match(/(?<year>\d{4})\D+(?<month>\d{1,2})/);
  if (!match?.groups) {
    return undefined;
  }

  const year = Number(match.groups.year);
  const month = Number(match.groups.month);
  if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) {
    return undefined;
  }

  return `${year}-${String(month).padStart(2, "0")}-01`;
}

function stripTags(html: string): string {
  return decodeHtml(html.replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]+>/g, " "));
}

function cleanText(text: string): string {
  return decodeHtml(text).replace(/\s+/g, " ").trim();
}

function sanitizeTitle(text: string): string {
  return cleanText(text).replace(/[<>:"/\\|?*]+/g, "_").slice(0, 150) || "caipiao";
}

function decodeHtml(text: string): string {
  return text
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, "\"")
    .replace(/&#39;/g, "'")
    .replace(/&#(\d+);/g, (_match, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_match, code) =>
      String.fromCodePoint(Number.parseInt(code, 16)),
    );
}

function basenameFromUrl(url: string): string {
  const pathname = new URL(url).pathname;
  return decodeURIComponent(pathname.split("/").pop() || "附件");
}

function contentTypeForExtension(extension: string): string {
  if (extension === ".xlsx") {
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }
  if (extension === ".xls") {
    return "application/vnd.ms-excel";
  }
  if (extension === ".pdf") {
    return "application/pdf";
  }
  return "application/octet-stream";
}

const DASHBOARD_HTML = `
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>全国彩票销售数据仪表盘</title>
  <script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/echarts/5.6.0/echarts.min.js"></script>
  <style>
    :root {
      color-scheme: light;
      --bg: #f3f7fc;
      --surface: #ffffff;
      --surface-soft: #f8fbff;
      --ink: #11275b;
      --muted: #7585a5;
      --line: #e1e9f4;
      --blue: #2d83e6;
      --blue-deep: #1554bd;
      --cyan: #2ac8bc;
      --orange: #ff8b28;
      --red: #f04a54;
      --purple: #8659e7;
      --green: #26a871;
      --shadow: 0 10px 30px rgba(41, 88, 150, 0.08);
    }

    * { box-sizing: border-box; }

    [hidden] { display: none !important; }

    html, body {
      margin: 0;
      min-height: 100%;
      background:
        radial-gradient(circle at 12% 0%, rgba(66, 145, 255, 0.08), transparent 28%),
        var(--bg);
      color: var(--ink);
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont,
        "Segoe UI", "Microsoft YaHei", sans-serif;
    }

    button, select { font: inherit; }
    button { cursor: pointer; }

    .shell {
      width: min(1440px, calc(100% - 44px));
      margin: 0 auto;
      padding: 24px 0 42px;
    }

    .topbar {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 24px;
      padding: 8px 8px 20px;
    }

    .brand {
      display: flex;
      align-items: center;
      gap: 13px;
    }

    .brand-mark {
      position: relative;
      width: 42px;
      height: 42px;
      flex: 0 0 auto;
    }

    .brand-mark span {
      position: absolute;
      display: block;
      border-radius: 2px;
      transform: rotate(18deg);
    }

    .mark-blue { width: 9px; height: 25px; left: 9px; top: 0; background: #175ac4; }
    .mark-red { width: 22px; height: 7px; right: 0; top: 11px; background: #ee3e4b; }
    .mark-yellow { width: 10px; height: 21px; left: 5px; bottom: 0; background: #f7ba2c; transform: rotate(-40deg) !important; }
    .mark-cyan { width: 20px; height: 8px; right: 4px; bottom: 5px; background: #2bc5b4; transform: rotate(42deg) !important; }

    .brand-name {
      font-size: 26px;
      line-height: 1.2;
      font-weight: 820;
      letter-spacing: 0.5px;
    }

    .brand-sub {
      margin-top: 6px;
      color: var(--muted);
      font-size: 12px;
      letter-spacing: 1px;
    }

    .top-meta {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.8;
      text-align: right;
      white-space: nowrap;
    }

    .top-meta strong { color: var(--ink); }

    .hero {
      display: flex;
      align-items: flex-end;
      justify-content: space-between;
      gap: 24px;
      padding: 20px 8px 18px;
    }

    .eyebrow {
      color: var(--blue);
      font-size: 12px;
      font-weight: 760;
      letter-spacing: 1.8px;
      text-transform: uppercase;
    }

    .hero h1 {
      margin: 5px 0 7px;
      font-size: 32px;
      line-height: 1.15;
      letter-spacing: 0.5px;
    }

    .hero p {
      max-width: 720px;
      margin: 0;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.6;
    }

    .page-switch {
      display: inline-flex;
      gap: 5px;
      padding: 5px;
      border: 1px solid var(--line);
      border-radius: 12px;
      background: rgba(255, 255, 255, 0.78);
      box-shadow: var(--shadow);
    }

    .page-switch button,
    .view-switch button {
      border: 0;
      border-radius: 8px;
      background: transparent;
      color: var(--muted);
      padding: 10px 16px;
      font-size: 13px;
      font-weight: 720;
    }

    .page-switch button.is-active,
    .view-switch button.is-active {
      background: var(--blue);
      color: #fff;
      box-shadow: 0 5px 14px rgba(45, 131, 230, 0.22);
    }

    .filters {
      display: flex;
      align-items: flex-end;
      flex-wrap: wrap;
      gap: 14px;
      margin: 0 0 18px;
      padding: 16px 18px;
      border: 1px solid var(--line);
      border-radius: 12px;
      background: rgba(255, 255, 255, 0.9);
      box-shadow: var(--shadow);
    }

    .filter-field {
      min-width: 150px;
    }

    .filter-field.wide { min-width: 230px; }

    .filter-field label {
      display: block;
      margin: 0 0 6px;
      color: var(--muted);
      font-size: 11px;
      font-weight: 720;
    }

    .filter-control,
    .filter-field select {
      width: 100%;
      height: 38px;
      border: 1px solid #d6e1ef;
      border-radius: 8px;
      background: #fff;
      color: var(--ink);
      padding: 0 11px;
      outline: none;
    }

    .filter-control:focus,
    .filter-field select:focus {
      border-color: var(--blue);
      box-shadow: 0 0 0 3px rgba(45, 131, 230, 0.12);
    }

    .range-pair {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .range-pair select { min-width: 108px; }
    .range-pair span { color: var(--muted); }

    .filter-actions {
      display: flex;
      gap: 8px;
      margin-left: auto;
    }

    .action {
      height: 38px;
      border: 1px solid #d6e1ef;
      border-radius: 8px;
      background: #fff;
      color: var(--ink);
      padding: 0 15px;
      font-size: 13px;
      font-weight: 720;
      white-space: nowrap;
    }

    .action.primary {
      border-color: var(--blue);
      background: var(--blue);
      color: #fff;
      box-shadow: 0 6px 16px rgba(45, 131, 230, 0.22);
    }

    .action:hover { transform: translateY(-1px); }

    .status {
      display: none;
      align-items: center;
      min-height: 42px;
      margin-bottom: 16px;
      padding: 10px 14px;
      border: 1px solid #f1d9a6;
      border-radius: 9px;
      background: #fffaf0;
      color: #8a6919;
      font-size: 13px;
    }

    .status.is-visible { display: flex; }

    .metrics {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
      margin-bottom: 16px;
    }

    .metric {
      position: relative;
      min-height: 124px;
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: 12px;
      background: var(--surface);
      padding: 18px 18px 15px 68px;
      box-shadow: var(--shadow);
    }

    .metric::after {
      position: absolute;
      right: 15px;
      bottom: 12px;
      width: 45px;
      height: 28px;
      border-bottom: 3px solid rgba(45, 131, 230, 0.25);
      border-right: 3px solid rgba(45, 131, 230, 0.16);
      content: "";
      transform: skewY(-24deg);
    }

    .metric-icon {
      position: absolute;
      left: 17px;
      top: 18px;
      display: grid;
      width: 38px;
      height: 38px;
      place-items: center;
      border-radius: 50%;
      background: #e8f2ff;
      color: var(--blue);
      font-size: 20px;
      font-weight: 820;
    }

    .metric:nth-child(2) .metric-icon { background: #e5fbf8; color: var(--cyan); }
    .metric:nth-child(3) .metric-icon { background: #fff1e6; color: var(--orange); }
    .metric:nth-child(4) .metric-icon { background: #f0eaff; color: var(--purple); }

    .metric-label {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.3;
    }

    .metric-value {
      margin-top: 8px;
      color: var(--ink);
      font-size: 24px;
      line-height: 1.1;
      font-weight: 820;
    }

    .metric-value.positive { color: var(--green); }
    .metric-value.negative { color: var(--red); }

    .metric-sub {
      margin-top: 8px;
      color: var(--muted);
      font-size: 12px;
    }

    .content-grid {
      display: grid;
      grid-template-columns: repeat(12, minmax(0, 1fr));
      gap: 16px;
    }

    .panel {
      min-width: 0;
      border: 1px solid var(--line);
      border-radius: 12px;
      background: var(--surface);
      padding: 17px;
      box-shadow: var(--shadow);
    }

    .panel-wide { grid-column: span 7; }
    .panel-side { grid-column: span 5; }
    .panel-half { grid-column: span 6; }
    .panel-full { grid-column: 1 / -1; }

    .panel-header {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 8px;
    }

    .panel-title {
      margin: 0;
      color: var(--ink);
      font-size: 17px;
      line-height: 1.3;
      font-weight: 820;
    }

    .panel-meta {
      color: var(--muted);
      font-size: 11px;
      white-space: nowrap;
    }

    .chart { width: 100%; height: 326px; }
    .chart-tall { height: 374px; }
    .chart-compact { height: 260px; }
    .donut-host { position: relative; }

    .donut-center {
      position: absolute;
      z-index: 5;
      top: 47%;
      left: 50%;
      width: 92%;
      transform: translate(-50%, -50%);
      pointer-events: none;
      text-align: center;
      line-height: 1.1;
      white-space: nowrap;
    }

    .donut-center-label,
    .donut-center-unit {
      display: block;
      color: #7b8ba8;
      font-size: 10px;
    }

    .donut-center-value {
      display: block;
      margin: 4px 0 2px;
      color: #17356e;
      font-size: 17px;
      font-weight: 800;
      letter-spacing: -0.35px;
    }

    .donut-center-value.is-long { font-size: 14px; }
    .donut-center-value.is-very-long { font-size: 12px; }

    .structure-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 9px;
    }

    .structure-card {
      min-width: 0;
      border-left: 1px solid #edf1f7;
      padding-left: 9px;
    }

    .structure-card:first-child {
      border-left: 0;
      padding-left: 0;
    }

    .structure-label {
      min-height: 32px;
      color: var(--muted);
      font-size: 11px;
      line-height: 1.4;
      text-align: center;
    }

    .dual-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
      grid-column: 1 / -1;
    }

    .view-switch {
      display: inline-flex;
      gap: 2px;
      padding: 3px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-soft);
    }

    .view-switch button {
      padding: 6px 12px;
      font-size: 11px;
    }

    .table-wrap {
      overflow-x: auto;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
    }

    th, td {
      border-bottom: 1px solid #e8eef6;
      padding: 10px 12px;
      text-align: right;
      white-space: nowrap;
    }

    th:first-child, td:first-child { text-align: left; }
    th {
      background: #f3f7fd;
      color: var(--muted);
      font-weight: 720;
    }

    td { color: #31517f; }
    td strong { color: var(--ink); }
    .positive-text { color: var(--green) !important; font-weight: 760; }
    .negative-text { color: var(--red) !important; font-weight: 760; }

    .deferred {
      grid-column: 1 / -1;
      border: 1px dashed #cbd9eb;
      border-radius: 12px;
      background: linear-gradient(135deg, #f7fbff, #ffffff);
      padding: 18px 20px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.7;
    }

    .deferred strong { color: var(--ink); }

    .footer {
      display: flex;
      justify-content: space-between;
      gap: 16px;
      margin-top: 22px;
      padding: 0 8px;
      color: var(--muted);
      font-size: 11px;
      line-height: 1.6;
    }

    @media (max-width: 1120px) {
      .metrics { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .panel-wide, .panel-side, .panel-half { grid-column: 1 / -1; }
    }

    @media (max-width: 720px) {
      .shell { width: min(100% - 20px, 1440px); padding-top: 10px; }
      .topbar, .hero { align-items: flex-start; flex-direction: column; }
      .top-meta { text-align: left; }
      .brand-name { font-size: 21px; }
      .hero h1 { font-size: 26px; }
      .page-switch { width: 100%; }
      .page-switch button { flex: 1; }
      .filter-actions { width: 100%; margin-left: 0; }
      .filter-actions .action { flex: 1; }
      .metrics { grid-template-columns: 1fr; }
      .structure-grid, .dual-grid { grid-template-columns: 1fr; }
      .structure-card { border-left: 0; border-top: 1px solid #edf1f7; padding: 8px 0 0; }
      .structure-card:first-child { border-top: 0; padding-top: 0; }
      .chart, .chart-tall { height: 300px; }
      .chart-compact { height: 250px; }
      .footer { flex-direction: column; }
    }
  </style>
</head>
<body>
  <main class="shell">
    <header class="topbar">
      <div class="brand">
        <div class="brand-mark" aria-hidden="true">
          <span class="mark-blue"></span>
          <span class="mark-red"></span>
          <span class="mark-yellow"></span>
          <span class="mark-cyan"></span>
        </div>
        <div>
          <div class="brand-name">全国彩票销售数据仪表盘</div>
          <div class="brand-sub">公益体彩 · 乐善人生 · 数据见证公益的力量</div>
        </div>
      </div>
      <div class="top-meta">
        <div>当前页面：<strong id="pageMeta">年度数据分析</strong></div>
        <div>数据更新时间：<span id="updatedAt">读取中</span></div>
      </div>
    </header>

    <section class="hero">
      <div>
        <div class="eyebrow">DATA ANALYTICS</div>
        <h1 id="pageTitle">年度数据分析</h1>
        <p id="pageSubtitle">支持连续年份区间对比分析，洞察彩票市场年度发展趋势与结构变化。</p>
      </div>
      <div class="page-switch" aria-label="分析页面">
        <button type="button" data-page="annual" class="is-active">年度数据分析</button>
        <button type="button" data-page="monthly">月度数据分析</button>
      </div>
    </section>

    <section id="annualControls" class="filters">
      <div class="filter-field wide">
        <label>年份区间（最多10年）</label>
        <div class="range-pair">
          <select id="annualStart" aria-label="年度起始年份"></select>
          <span>—</span>
          <select id="annualEnd" aria-label="年度结束年份"></select>
        </div>
      </div>
      <div class="filter-field">
        <label for="annualType">彩票类型</label>
        <select id="annualType"></select>
      </div>
      <div class="filter-actions">
        <button type="button" id="annualRefresh" class="action primary">↻ 刷新分析</button>
        <button type="button" id="annualExport" class="action">⇩ 导出数据</button>
      </div>
    </section>

    <section id="monthlyControls" class="filters" hidden>
      <div class="filter-field">
        <label for="monthlyYear">年份</label>
        <select id="monthlyYear"></select>
      </div>
      <div class="filter-field">
        <label for="monthlyMonth">月份</label>
        <select id="monthlyMonth"></select>
      </div>
      <div class="filter-field wide">
        <label for="monthlyType">彩票类型</label>
        <select id="monthlyType"></select>
      </div>
      <div class="filter-field">
        <label>趋势粒度</label>
        <div class="view-switch">
          <button type="button" data-granularity="month" class="is-active">按月</button>
          <button type="button" data-granularity="quarter">按季度</button>
        </div>
      </div>
      <div class="filter-actions">
        <button type="button" id="monthlyRefresh" class="action primary">↻ 刷新数据</button>
        <button type="button" id="monthlyExport" class="action">⇩ 导出数据</button>
      </div>
    </section>

    <div id="status" class="status"></div>

    <section id="annualView">
      <section class="metrics">
        <article class="metric">
          <div class="metric-icon">¥</div>
          <div class="metric-label">年度销售额</div>
          <div id="annualTotal" class="metric-value">-</div>
          <div id="annualTotalSub" class="metric-sub">-</div>
        </article>
        <article class="metric">
          <div class="metric-icon">◉</div>
          <div class="metric-label">区间累计销售额</div>
          <div id="annualRangeTotal" class="metric-value">-</div>
          <div id="annualRangeSub" class="metric-sub">-</div>
        </article>
        <article class="metric">
          <div class="metric-icon">↗</div>
          <div class="metric-label">区间年复合增长率</div>
          <div id="annualCagr" class="metric-value">-</div>
          <div id="annualCagrSub" class="metric-sub">-</div>
        </article>
        <article class="metric">
          <div class="metric-icon">♛</div>
          <div class="metric-label">最高年份</div>
          <div id="annualTopYear" class="metric-value">-</div>
          <div id="annualTopYearSub" class="metric-sub">-</div>
        </article>
      </section>

      <section class="content-grid">
        <article class="panel panel-wide">
          <div class="panel-header">
            <h2 class="panel-title">年度销售趋势</h2>
            <span class="panel-meta">销售额 / 同比增长率</span>
          </div>
          <div id="annualTrendChart" class="chart"></div>
        </article>

        <article class="panel panel-side">
          <div class="panel-header">
            <h2 class="panel-title">年度结构分析</h2>
            <span id="annualStructureMeta" class="panel-meta">-</span>
          </div>
          <div class="structure-grid">
            <div class="structure-card">
              <div class="structure-label">体彩 / 福彩年度占比</div>
              <div id="annualGroupChart" class="chart chart-compact"></div>
            </div>
            <div class="structure-card">
              <div class="structure-label">主要彩种年度占比</div>
              <div id="annualTypeChart" class="chart chart-compact"></div>
            </div>
            <div class="structure-card">
              <div class="structure-label">最新年度机构占比</div>
              <div id="annualLatestGroupChart" class="chart chart-compact"></div>
            </div>
          </div>
        </article>

        <article class="panel panel-full">
          <div class="panel-header">
            <h2 class="panel-title">年度对比分析</h2>
            <span class="panel-meta">主要彩种销售额 / 全国与机构同比</span>
          </div>
          <div id="annualComparisonChart" class="chart chart-tall"></div>
        </article>

        <article class="panel panel-full">
          <div class="panel-header">
            <h2 class="panel-title">年度核心数据明细</h2>
            <span class="panel-meta">公益金字段待数据源补齐</span>
          </div>
          <div id="annualTable" class="table-wrap"></div>
        </article>

        <div class="deferred">
          <strong>地区分析暂不展示：</strong>当前项目的地区数据最新到 2015-02，且尚未配置省级地图边界，无法生成参考图中的 2024/最新地区地图和 TOP10。全国销售额、机构和彩票类型分析不受此限制。
        </div>
      </section>
    </section>

    <section id="monthlyView" hidden>
      <section class="metrics">
        <article class="metric">
          <div class="metric-icon">¥</div>
          <div class="metric-label">最新月份销售额</div>
          <div id="monthlyTotal" class="metric-value">-</div>
          <div id="monthlyTotalSub" class="metric-sub">-</div>
        </article>
        <article class="metric">
          <div class="metric-icon">↕</div>
          <div class="metric-label">环比</div>
          <div id="monthlyMom" class="metric-value">-</div>
          <div class="metric-sub">较上月</div>
        </article>
        <article class="metric">
          <div class="metric-icon">↗</div>
          <div class="metric-label">同比</div>
          <div id="monthlyYoy" class="metric-value">-</div>
          <div class="metric-sub">较去年同月</div>
        </article>
        <article class="metric">
          <div class="metric-icon">♛</div>
          <div class="metric-label">最大类型</div>
          <div id="monthlyTopType" class="metric-value">-</div>
          <div id="monthlyTopTypeSub" class="metric-sub">-</div>
        </article>
      </section>

      <section class="content-grid">
        <article class="panel panel-full">
          <div class="panel-header">
            <h2 class="panel-title">全国销售趋势</h2>
            <span id="monthlyTrendMeta" class="panel-meta">-</span>
          </div>
          <div id="monthlyTrendChart" class="chart"></div>
        </article>

        <article class="panel panel-full">
          <div class="panel-header">
            <h2 class="panel-title">最新结构</h2>
            <span id="monthlyStructureMeta" class="panel-meta">-</span>
          </div>
          <div class="structure-grid">
            <div class="structure-card">
              <div class="structure-label">体彩 / 福彩占比</div>
              <div id="monthlyGroupChart" class="chart chart-compact"></div>
            </div>
            <div class="structure-card">
              <div class="structure-label">体彩类型占比</div>
              <div id="monthlySportsChart" class="chart chart-compact"></div>
            </div>
            <div class="structure-card">
              <div class="structure-label">福彩类型占比</div>
              <div id="monthlyWelfareChart" class="chart chart-compact"></div>
            </div>
          </div>
        </article>

        <div class="dual-grid">
          <article class="panel">
            <div class="panel-header">
              <h2 class="panel-title">福彩 / 体彩趋势</h2>
              <span class="panel-meta">亿元</span>
            </div>
            <div id="monthlyGroupTrendChart" class="chart"></div>
          </article>
          <article class="panel">
            <div class="panel-header">
              <h2 class="panel-title">年度销售与机构占比</h2>
              <span class="panel-meta">年度累计 / 占比</span>
            </div>
            <div id="monthlyAnnualChart" class="chart"></div>
          </article>
        </div>

        <article class="panel panel-full">
          <div class="panel-header">
            <h2 class="panel-title">增长变化</h2>
            <span class="panel-meta">环比 / 同比</span>
          </div>
          <div id="monthlyGrowthChart" class="chart"></div>
        </article>

        <div class="deferred">
          <strong>地区销售分析暂不展示：</strong>当前地区数据只覆盖到 2015-02，且缺少省级地图边界。待地区数据补齐后，可在此加入地图和地区销售额 TOP10。
        </div>
      </section>
    </section>

    <footer class="footer">
      <span id="coverageText">数据来自 D1 预计算快照，销售额原始单位为万元，页面展示单位为亿元。</span>
      <span>公益体彩 · 乐善人生</span>
    </footer>
  </main>

  <script>
    (function () {
      var charts = {};
      var snapshot = null;
      var state = {
        page: 'annual',
        annualStart: '',
        annualEnd: '',
        annualType: 'all',
        monthlyYear: '',
        monthlyMonth: 'all',
        monthlyType: 'all',
        granularity: 'month'
      };
      var palette = ['#2d83e6', '#2ac8bc', '#ffb52b', '#8659e7', '#f04a54', '#26a871', '#ff8b28'];

      function el(id) {
        return document.getElementById(id);
      }

      function escapeHtml(value) {
        return String(value === null || value === undefined ? '' : value)
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;')
          .replace(/'/g, '&#39;');
      }

      function numberValue(value) {
        var result = Number(value);
        return Number.isFinite(result) ? result : 0;
      }

      function round(value) {
        return Math.round(numberValue(value) * 100) / 100;
      }

      function toYi(value) {
        return round(numberValue(value) / 10000);
      }

      function formatYi(value) {
        return toYi(value).toLocaleString('zh-CN', { maximumFractionDigits: 2 }) + ' 亿元';
      }

      function formatPercent(value) {
        if (value === null || value === undefined || !Number.isFinite(Number(value))) {
          return '-';
        }
        var number = Number(value);
        return (number > 0 ? '+' : '') + number.toFixed(1) + '%';
      }

      function percentChange(current, previous) {
        if (!previous || !Number.isFinite(current) || !Number.isFinite(previous)) {
          return null;
        }
        return round((current - previous) / previous * 100);
      }

      function dateYear(reportDate) {
        return String(reportDate || '').slice(0, 4);
      }

      function dateMonth(reportDate) {
        return String(reportDate || '').slice(5, 7);
      }

      function shiftMonths(reportDate, amount) {
        var year = Number(String(reportDate).slice(0, 4));
        var month = Number(String(reportDate).slice(5, 7));
        var date = new Date(Date.UTC(year, month - 1 + amount, 1));
        return date.getUTCFullYear() + '-' + String(date.getUTCMonth() + 1).padStart(2, '0') + '-01';
      }

      function displayMonth(reportDate) {
        if (!reportDate) {
          return '-';
        }
        return dateYear(reportDate) + '年' + Number(dateMonth(reportDate)) + '月';
      }

      function displayUpdated(value) {
        if (!value) {
          return '-';
        }
        return new Date(value).toLocaleString('zh-CN', {
          year: 'numeric',
          month: '2-digit',
          day: '2-digit',
          hour: '2-digit',
          minute: '2-digit'
        });
      }

      function setText(id, value) {
        var node = el(id);
        if (node) {
          node.textContent = value === null || value === undefined ? '-' : String(value);
        }
      }

      function setPercentClass(id, value) {
        var node = el(id);
        if (!node) {
          return;
        }
        node.classList.remove('positive', 'negative');
        if (Number(value) > 0) {
          node.classList.add('positive');
        } else if (Number(value) < 0) {
          node.classList.add('negative');
        }
      }

      function showStatus(message, visible) {
        var node = el('status');
        node.textContent = message || '';
        node.classList.toggle('is-visible', Boolean(visible));
      }

      function clearCharts() {
        Object.keys(charts).forEach(function (key) {
          charts[key].dispose();
        });
        charts = {};
      }

      function makeChart(id, option) {
        var node = el(id);
        if (!node || !window.echarts) {
          return;
        }
        if (charts[id]) {
          charts[id].dispose();
        }
        var chart = window.echarts.init(node, null, { renderer: 'canvas' });
        chart.setOption(option);
        charts[id] = chart;
      }

      function emptyChart(id, message) {
        makeChart(id, {
          graphic: {
            type: 'text',
            left: 'center',
            top: 'middle',
            style: { text: message || '暂无数据', fill: '#94a3bd', fontSize: 14 }
          }
        });
      }

      function commonGrid(extra) {
        var base = {
          left: 48,
          right: 34,
          top: 46,
          bottom: 42,
          containLabel: true
        };
        return Object.assign(base, extra || {});
      }

      function setOptions(id, options, selected) {
        var node = el(id);
        if (!node) {
          return;
        }
        var current = selected;
        node.innerHTML = '';
        options.forEach(function (option) {
          var item = document.createElement('option');
          item.value = option.value;
          item.textContent = option.label;
          node.appendChild(item);
        });
        if (options.some(function (option) { return option.value === current; })) {
          node.value = current;
        } else if (options.length) {
          node.value = options[0].value;
        }
      }

      function yearsFromSnapshot(data) {
        return (data.annualTrend || []).map(function (row) { return row.year; })
          .filter(function (year, index, rows) { return rows.indexOf(year) === index; })
          .sort();
      }

      function typeOptions(data) {
        var options = [{ value: 'all', label: '全部' }];
        var seen = {};
        ['福彩', '体彩'].forEach(function (group) {
          var hasGroup = (data.typeTrend || []).some(function (row) {
            return row.lotteryGroup === group;
          });
          if (hasGroup) {
            options.push({ value: group, label: group });
          }
        });
        (data.typeTrend || []).forEach(function (row) {
          var value = row.lotteryGroup + '|' + row.lotteryType;
          if (!seen[value]) {
            seen[value] = true;
            options.push({ value: value, label: row.lotteryGroup + ' · ' + row.lotteryType });
          }
        });
        return options;
      }

      function matchesType(row, filter) {
        if (!filter || filter === 'all') {
          return true;
        }
        if (filter === '福彩' || filter === '体彩') {
          return row.lotteryGroup === filter;
        }
        var split = filter.split('|');
        return row.lotteryGroup === split[0] && row.lotteryType === split.slice(1).join('|');
      }

      function buildMonthlySeries(data, filter) {
        var byDate = {};
        (data.typeTrend || []).forEach(function (row) {
          if (!matchesType(row, filter)) {
            return;
          }
          if (!byDate[row.reportDate]) {
            byDate[row.reportDate] = {
              reportDate: row.reportDate,
              totalSales: 0,
              welfareSales: 0,
              sportsSales: 0
            };
          }
          byDate[row.reportDate].totalSales += numberValue(row.monthlySales);
          if (row.lotteryGroup === '福彩') {
            byDate[row.reportDate].welfareSales += numberValue(row.monthlySales);
          }
          if (row.lotteryGroup === '体彩') {
            byDate[row.reportDate].sportsSales += numberValue(row.monthlySales);
          }
        });
        var rows = Object.keys(byDate).sort().map(function (date) {
          var row = byDate[date];
          row.totalSales = round(row.totalSales);
          row.welfareSales = round(row.welfareSales);
          row.sportsSales = round(row.sportsSales);
          return row;
        });
        var byDateRows = {};
        rows.forEach(function (row) { byDateRows[row.reportDate] = row; });
        rows.forEach(function (row, index) {
          var previous = rows[index - 1];
          var previousYear = byDateRows[shiftMonths(row.reportDate, -12)];
          row.monthOverMonthPercent = previous
            ? percentChange(row.totalSales, previous.totalSales)
            : null;
          row.yearOverYearPercent = previousYear && previousYear.totalSales >= 10000
            ? percentChange(row.totalSales, previousYear.totalSales)
            : null;
        });
        return rows;
      }

      function buildAnnualSeries(data, filter) {
        var monthly = buildMonthlySeries(data, filter);
        var byYear = {};
        monthly.forEach(function (row) {
          var year = dateYear(row.reportDate);
          if (!byYear[year]) {
            byYear[year] = {
              year: year,
              totalSales: 0,
              welfareSales: 0,
              sportsSales: 0,
              months: {}
            };
          }
          byYear[year].totalSales += row.totalSales;
          byYear[year].welfareSales += row.welfareSales;
          byYear[year].sportsSales += row.sportsSales;
          byYear[year].months[row.reportDate] = true;
        });
        var rows = Object.keys(byYear).sort().map(function (year) {
          var row = byYear[year];
          row.totalSales = round(row.totalSales);
          row.welfareSales = round(row.welfareSales);
          row.sportsSales = round(row.sportsSales);
          row.monthCount = Object.keys(row.months).length;
          row.welfareSharePercent = row.totalSales ? round(row.welfareSales / row.totalSales * 100) : 0;
          row.sportsSharePercent = row.totalSales ? round(row.sportsSales / row.totalSales * 100) : 0;
          row.isPartial = row.monthCount < 12;
          delete row.months;
          return row;
        });
        var byYearRows = {};
        rows.forEach(function (row) { byYearRows[row.year] = row; });
        rows.forEach(function (row) {
          var previous = byYearRows[String(Number(row.year) - 1)];
          row.yearOverYearPercent = previous && previous.totalSales >= 10000
            ? percentChange(row.totalSales, previous.totalSales)
            : null;
          row.sportsYearOverYearPercent = previous && previous.sportsSales >= 10000
            ? percentChange(row.sportsSales, previous.sportsSales)
            : null;
          row.welfareYearOverYearPercent = previous && previous.welfareSales >= 10000
            ? percentChange(row.welfareSales, previous.welfareSales)
            : null;
        });
        return rows;
      }

      function filterYears(rows, start, end) {
        return rows.filter(function (row) {
          return (!start || row.year >= start) && (!end || row.year <= end);
        });
      }

      function aggregateQuarters(rows) {
        var byQuarter = {};
        rows.forEach(function (row) {
          var month = Number(dateMonth(row.reportDate));
          var quarter = Math.floor((month - 1) / 3) + 1;
          var key = dateYear(row.reportDate) + '-Q' + quarter;
          if (!byQuarter[key]) {
            byQuarter[key] = {
              key: key,
              reportDate: key,
              totalSales: 0,
              welfareSales: 0,
              sportsSales: 0
            };
          }
          byQuarter[key].totalSales += row.totalSales;
          byQuarter[key].welfareSales += row.welfareSales;
          byQuarter[key].sportsSales += row.sportsSales;
        });
        var result = Object.keys(byQuarter).sort().map(function (key) {
          var row = byQuarter[key];
          row.totalSales = round(row.totalSales);
          row.welfareSales = round(row.welfareSales);
          row.sportsSales = round(row.sportsSales);
          return row;
        });
        result.forEach(function (row, index) {
          var previous = result[index - 1];
          var previousYear = result[index - 4];
          row.monthOverMonthPercent = previous ? percentChange(row.totalSales, previous.totalSales) : null;
          row.yearOverYearPercent = previousYear && previousYear.totalSales >= 10000
            ? percentChange(row.totalSales, previousYear.totalSales)
            : null;
        });
        return result;
      }

      function typeTotals(data, filter, startYear, endYear) {
        var totals = {};
        (data.typeTrend || []).forEach(function (row) {
          var year = dateYear(row.reportDate);
          if (!matchesType(row, filter) || (startYear && year < startYear) || (endYear && year > endYear)) {
            return;
          }
          var key = row.lotteryGroup + '|' + row.lotteryType;
          if (!totals[key]) {
            totals[key] = {
              key: key,
              name: row.lotteryType,
              group: row.lotteryGroup,
              value: 0
            };
          }
          totals[key].value += numberValue(row.monthlySales);
        });
        return Object.keys(totals).map(function (key) {
          totals[key].value = round(totals[key].value);
          return totals[key];
        }).sort(function (left, right) {
          return right.value - left.value;
        });
      }

      function latestBreakdowns(data, reportDate, filter) {
        var group = {};
        var sports = {};
        var welfare = {};
        (data.typeTrend || []).forEach(function (row) {
          if (row.reportDate !== reportDate || !matchesType(row, filter)) {
            return;
          }
          group[row.lotteryGroup] = (group[row.lotteryGroup] || 0) + numberValue(row.monthlySales);
          var target = row.lotteryGroup === '体彩' ? sports : welfare;
          var key = row.lotteryType;
          target[key] = (target[key] || 0) + numberValue(row.monthlySales);
        });
        function mapBreakdown(source) {
          var total = Object.keys(source).reduce(function (sum, key) { return sum + source[key]; }, 0);
          return Object.keys(source).map(function (key) {
            return {
              name: key,
              value: round(source[key]),
              share: total ? round(source[key] / total * 100) : 0
            };
          }).sort(function (left, right) { return right.value - left.value; });
        }
        return {
          group: mapBreakdown(group),
          sports: mapBreakdown(sports),
          welfare: mapBreakdown(welfare)
        };
      }

      function donutOption(rows, colors, centerValue, centerLabel) {
        if (!rows.length) {
          return {
            graphic: {
              type: 'text',
              left: 'center',
              top: 'middle',
              style: { text: '暂无数据', fill: '#94a3bd', fontSize: 13 }
            }
          };
        }
        return {
          color: colors,
          tooltip: {
            trigger: 'item',
            formatter: function (item) {
              return item.name + '<br>' + formatYi(item.value * 10000) + ' (' + Number(item.percent).toFixed(1) + '%)';
            }
          },
          legend: {
            bottom: 0,
            type: 'scroll',
            itemWidth: 10,
            itemHeight: 8,
            textStyle: { color: '#607497', fontSize: 10 }
          },
          series: [{
            type: 'pie',
            radius: ['44%', '70%'],
            center: ['50%', '47%'],
            avoidLabelOverlap: true,
            itemStyle: { borderColor: '#fff', borderWidth: 2 },
            label: { formatter: '{b}\\n{d}%', fontSize: 10 },
            labelLine: { length: 8, length2: 5 },
            data: rows.map(function (row) { return { name: row.name, value: toYi(row.value) }; })
          }]
        };
      }

      function renderDonutCenter(id, centerValue, centerLabel) {
        var node = el(id);
        if (!node) {
          return;
        }
        node.classList.add('donut-host');
        var valueText = formatYi(centerValue * 10000).replace(' 亿元', '');
        var valueClass = valueText.length >= 11
          ? 'donut-center-value is-very-long'
          : valueText.length >= 9 ? 'donut-center-value is-long' : 'donut-center-value';
        var center = node.querySelector('.donut-center');
        if (!center) {
          center = document.createElement('div');
          center.className = 'donut-center';
          node.appendChild(center);
        }
        center.innerHTML = '<span class="donut-center-label">' + escapeHtml(centerLabel || '总销售额') + '</span>' +
          '<strong class="' + valueClass + '">' + escapeHtml(valueText) + '</strong>' +
          '<span class="donut-center-unit">亿元</span>';
      }

      function renderDonut(id, rows, colors, centerValue, centerLabel) {
        makeChart(id, donutOption(rows, colors, centerValue, centerLabel));
        if (rows.length) {
          renderDonutCenter(id, centerValue, centerLabel);
        }
      }

      function renderAnnual(data) {
        var annual = filterYears(buildAnnualSeries(data, state.annualType), state.annualStart, state.annualEnd);
        if (!annual.length) {
          ['annualTrendChart', 'annualComparisonChart', 'annualTable'].forEach(function (id) { emptyChart(id, '当前筛选没有数据'); });
          return;
        }
        var latest = annual[annual.length - 1];
        var rangeTotal = annual.reduce(function (sum, row) { return sum + row.totalSales; }, 0);
        var first = annual[0];
        var years = Math.max(1, Number(latest.year) - Number(first.year));
        var cagr = first.totalSales > 0 && latest.totalSales > 0
          ? (Math.pow(latest.totalSales / first.totalSales, 1 / years) - 1) * 100
          : null;
        var top = annual.reduce(function (left, right) {
          return right.totalSales > left.totalSales ? right : left;
        });

        setText('annualTotal', formatYi(latest.totalSales));
        setText('annualTotalSub', latest.year + '年' + (latest.isPartial ? '（部分月份）' : ''));
        setText('annualRangeTotal', formatYi(rangeTotal));
        setText('annualRangeSub', first.year + '—' + latest.year + '年累计');
        setText('annualCagr', formatPercent(cagr));
        setPercentClass('annualCagr', cagr);
        setText('annualCagrSub', first.year + '至' + latest.year);
        setText('annualTopYear', top.year + '年');
        setText('annualTopYearSub', formatYi(top.totalSales) + (top.isPartial ? ' · 部分月份' : ''));
        setText('annualStructureMeta', first.year + '—' + latest.year);

        makeChart('annualTrendChart', {
          color: ['#2d83e6', '#ff8b28'],
          tooltip: { trigger: 'axis' },
          legend: { top: 5, data: ['全国销售额', '同比增长率'], textStyle: { color: '#607497' } },
          grid: commonGrid(),
          xAxis: { type: 'category', data: annual.map(function (row) { return row.year + (row.isPartial ? '*' : ''); }) },
          yAxis: [
            { type: 'value', name: '亿元', splitLine: { lineStyle: { color: '#edf2f8' } } },
            { type: 'value', name: '%', splitLine: { show: false } }
          ],
          series: [
            {
              name: '全国销售额',
              type: 'bar',
              barMaxWidth: 38,
              itemStyle: { borderRadius: [5, 5, 0, 0] },
              label: { show: true, position: 'top', color: '#294b80', formatter: function (item) { return toYi(item.value).toFixed(2); } },
              data: annual.map(function (row) { return toYi(row.totalSales); })
            },
            {
              name: '同比增长率',
              type: 'line',
              yAxisIndex: 1,
              smooth: true,
              symbolSize: 8,
              lineStyle: { width: 2, color: '#ff8b28' },
              itemStyle: { color: '#ff8b28' },
              data: annual.map(function (row) { return row.yearOverYearPercent; })
            }
          ]
        });

        var groupRows = [
          { name: '中国体育彩票', value: annual.reduce(function (sum, row) { return sum + row.sportsSales; }, 0) },
          { name: '中国福利彩票', value: annual.reduce(function (sum, row) { return sum + row.welfareSales; }, 0) }
        ];
        var annualTypes = typeTotals(data, state.annualType, state.annualStart, state.annualEnd).slice(0, 5)
          .map(function (row) { return { name: row.group + row.name, value: row.value }; });
        var latestGroupRows = [
          { name: '体育彩票', value: latest.sportsSales },
          { name: '福利彩票', value: latest.welfareSales }
        ];
        renderDonut('annualGroupChart', groupRows, ['#2d83e6', '#8659e7'], toYi(rangeTotal), '区间销售额');
        renderDonut('annualTypeChart', annualTypes, palette, toYi(annualTypes.reduce(function (sum, row) { return sum + row.value; }, 0)), '主要彩种');
        renderDonut('annualLatestGroupChart', latestGroupRows, ['#2d83e6', '#2ac8bc'], toYi(latest.totalSales), '最新年度');

        var typeRows = typeTotals(data, state.annualType, state.annualStart, state.annualEnd).slice(0, 5);
        var typeYears = annual.map(function (row) { return row.year; });
        var typeSeries = typeRows.map(function (typeRow) {
          return {
            name: typeRow.group + typeRow.name,
            type: 'bar',
            barMaxWidth: 18,
            data: typeYears.map(function (year) {
              return toYi((data.typeTrend || []).filter(function (row) {
                return row.reportDate.slice(0, 4) === year &&
                  row.lotteryGroup === typeRow.group &&
                  row.lotteryType === typeRow.name;
              }).reduce(function (sum, row) { return sum + numberValue(row.monthlySales); }, 0));
            })
          };
        });
        makeChart('annualComparisonChart', {
          tooltip: { trigger: 'axis' },
          legend: { top: 4, type: 'scroll', textStyle: { color: '#607497' } },
          grid: { left: 50, right: 52, top: 58, bottom: 46, containLabel: true },
          axisPointer: { link: [{ xAxisIndex: 'all' }] },
          xAxis: [
            { type: 'category', data: typeYears, gridIndex: 0 },
            { type: 'category', data: typeYears, gridIndex: 1 }
          ],
          yAxis: [
            { type: 'value', name: '彩种销售额（亿元）', gridIndex: 0, splitLine: { lineStyle: { color: '#edf2f8' } } },
            { type: 'value', name: '同比（%）', gridIndex: 1, splitLine: { lineStyle: { color: '#edf2f8' } } }
          ],
          grid: [
            { left: 54, right: '52%', top: 58, bottom: 48, containLabel: true },
            { left: '54%', right: 42, top: 58, bottom: 48, containLabel: true }
          ],
          series: typeSeries.concat([
            {
              name: '全国同比',
              type: 'bar',
              xAxisIndex: 1,
              yAxisIndex: 1,
              data: annual.map(function (row) { return row.yearOverYearPercent; }),
              itemStyle: { color: '#2d83e6' }
            },
            {
              name: '体彩同比',
              type: 'bar',
              xAxisIndex: 1,
              yAxisIndex: 1,
              data: annual.map(function (row) { return row.sportsYearOverYearPercent; }),
              itemStyle: { color: '#f04a54' }
            },
            {
              name: '福彩同比',
              type: 'bar',
              xAxisIndex: 1,
              yAxisIndex: 1,
              data: annual.map(function (row) { return row.welfareYearOverYearPercent; }),
              itemStyle: { color: '#26a871' }
            }
          ])
        });

        el('annualTable').innerHTML = '<table><thead><tr>' +
          '<th>年份</th><th>全国销售额（亿元）</th><th>同比增长</th>' +
          '<th>体彩销售额（亿元）</th><th>福彩销售额（亿元）</th><th>已入库月份</th>' +
          '</tr></thead><tbody>' +
          annual.slice().reverse().map(function (row) {
            var yoyClass = Number(row.yearOverYearPercent) > 0 ? 'positive-text' : Number(row.yearOverYearPercent) < 0 ? 'negative-text' : '';
            return '<tr><td><strong>' + escapeHtml(row.year + (row.isPartial ? '*' : '')) + '</strong></td>' +
              '<td>' + toYi(row.totalSales).toFixed(2) + '</td>' +
              '<td class="' + yoyClass + '">' + escapeHtml(formatPercent(row.yearOverYearPercent)) + '</td>' +
              '<td>' + toYi(row.sportsSales).toFixed(2) + '</td>' +
              '<td>' + toYi(row.welfareSales).toFixed(2) + '</td>' +
              '<td>' + row.monthCount + ' / 12</td></tr>';
          }).join('') + '</tbody></table>';
      }

      function renderMonthly(data) {
        var full = buildMonthlySeries(data, state.monthlyType);
        var yearRows = full.filter(function (row) { return dateYear(row.reportDate) === state.monthlyYear; });
        var chartRows = state.granularity === 'quarter' ? aggregateQuarters(yearRows) : yearRows;
        var selectedDate = state.monthlyMonth === 'all'
          ? (yearRows.length ? yearRows[yearRows.length - 1].reportDate : null)
          : state.monthlyYear + '-' + state.monthlyMonth + '-01';
        var selected = full.find(function (row) { return row.reportDate === selectedDate; }) || yearRows[yearRows.length - 1];
        if (!selected) {
          ['monthlyTrendChart', 'monthlyGroupTrendChart', 'monthlyAnnualChart', 'monthlyGrowthChart'].forEach(function (id) { emptyChart(id, '当前筛选没有数据'); });
          return;
        }
        var breakdowns = latestBreakdowns(data, selected.reportDate, state.monthlyType);
        var selectedTypes = (data.typeTrend || []).filter(function (row) {
          return row.reportDate === selected.reportDate && matchesType(row, state.monthlyType);
        });
        var topType = selectedTypes.reduce(function (top, row) {
          return !top || numberValue(row.monthlySales) > numberValue(top.monthlySales) ? row : top;
        }, null);
        setText('monthlyTotal', formatYi(selected.totalSales));
        setText('monthlyTotalSub', displayMonth(selected.reportDate));
        setText('monthlyMom', formatPercent(selected.monthOverMonthPercent));
        setPercentClass('monthlyMom', selected.monthOverMonthPercent);
        setText('monthlyYoy', formatPercent(selected.yearOverYearPercent));
        setPercentClass('monthlyYoy', selected.yearOverYearPercent);
        setText('monthlyTopType', topType ? topType.lotteryGroup + topType.lotteryType : '-');
        setText('monthlyTopTypeSub', topType ? formatYi(topType.monthlySales) : '-');
        setText('monthlyTrendMeta', state.granularity === 'quarter' ? state.monthlyYear + '年季度视图' : state.monthlyYear + '年月度视图');
        setText('monthlyStructureMeta', displayMonth(selected.reportDate));

        makeChart('monthlyTrendChart', {
          color: ['#2d83e6', '#ff8b28'],
          tooltip: { trigger: 'axis' },
          legend: { top: 5, data: ['月度销售额', '同比增长率'], textStyle: { color: '#607497' } },
          grid: commonGrid({ right: 54 }),
          xAxis: { type: 'category', data: chartRows.map(function (row) { return state.granularity === 'quarter' ? row.key : Number(dateMonth(row.reportDate)) + '月'; }) },
          yAxis: [
            { type: 'value', name: '亿元', splitLine: { lineStyle: { color: '#edf2f8' } } },
            { type: 'value', name: '%', splitLine: { show: false } }
          ],
          series: [
            {
              name: '月度销售额',
              type: 'bar',
              barMaxWidth: 34,
              itemStyle: { borderRadius: [4, 4, 0, 0] },
              data: chartRows.map(function (row) { return toYi(row.totalSales); })
            },
            {
              name: '同比增长率',
              type: 'line',
              yAxisIndex: 1,
              smooth: true,
              symbolSize: 7,
              itemStyle: { color: '#ff8b28' },
              lineStyle: { color: '#ff8b28', width: 2 },
              data: chartRows.map(function (row) { return row.yearOverYearPercent; })
            }
          ]
        });

        renderDonut('monthlyGroupChart', breakdowns.group, ['#2d83e6', '#2ac8bc'], toYi(selected.totalSales), '总销售额');
        renderDonut('monthlySportsChart', breakdowns.sports, ['#2d83e6', '#5ab1ef', '#26a871', '#ffb52b', '#8659e7'], toYi(breakdowns.sports.reduce(function (sum, row) { return sum + row.value; }, 0)), '体彩');
        renderDonut('monthlyWelfareChart', breakdowns.welfare, ['#f04a54', '#5ab1ef', '#ffb52b', '#ff8b28', '#8659e7'], toYi(breakdowns.welfare.reduce(function (sum, row) { return sum + row.value; }, 0)), '福彩');

        makeChart('monthlyGroupTrendChart', {
          color: ['#2d83e6', '#f04a54'],
          tooltip: { trigger: 'axis' },
          legend: { top: 5, data: ['中国体育彩票', '中国福利彩票'], textStyle: { color: '#607497' } },
          grid: commonGrid(),
          xAxis: { type: 'category', data: chartRows.map(function (row) { return state.granularity === 'quarter' ? row.key : Number(dateMonth(row.reportDate)) + '月'; }) },
          yAxis: { type: 'value', name: '亿元', splitLine: { lineStyle: { color: '#edf2f8' } } },
          series: [
            { name: '中国体育彩票', type: 'line', smooth: true, symbolSize: 6, data: chartRows.map(function (row) { return toYi(row.sportsSales); }) },
            { name: '中国福利彩票', type: 'line', smooth: true, symbolSize: 6, data: chartRows.map(function (row) { return toYi(row.welfareSales); }) }
          ]
        });

        var annual = buildAnnualSeries(data, state.monthlyType).slice(-5);
        makeChart('monthlyAnnualChart', {
          color: ['#2d83e6', '#f04a54'],
          tooltip: { trigger: 'axis' },
          legend: { top: 5, data: ['体彩销售额', '福彩销售额'], textStyle: { color: '#607497' } },
          grid: commonGrid({ right: 56 }),
          xAxis: { type: 'category', data: annual.map(function (row) { return row.year; }) },
          yAxis: { type: 'value', name: '亿元', splitLine: { lineStyle: { color: '#edf2f8' } } },
          series: [
            { name: '体彩销售额', type: 'bar', stack: 'annual', data: annual.map(function (row) { return toYi(row.sportsSales); }) },
            { name: '福彩销售额', type: 'bar', stack: 'annual', data: annual.map(function (row) { return toYi(row.welfareSales); }) }
          ]
        });

        makeChart('monthlyGrowthChart', {
          color: ['#2d83e6', '#f04a54'],
          tooltip: { trigger: 'axis' },
          legend: { top: 5, data: ['环比', '同比'], textStyle: { color: '#607497' } },
          grid: commonGrid({ top: 48 }),
          xAxis: { type: 'category', data: chartRows.map(function (row) { return state.granularity === 'quarter' ? row.key : Number(dateMonth(row.reportDate)) + '月'; }) },
          yAxis: { type: 'value', name: '%', splitLine: { lineStyle: { color: '#edf2f8' } } },
          series: [
            { name: '环比', type: 'bar', barMaxWidth: 20, data: chartRows.map(function (row) { return row.monthOverMonthPercent; }) },
            { name: '同比', type: 'bar', barMaxWidth: 20, data: chartRows.map(function (row) { return row.yearOverYearPercent; }) }
          ]
        });
      }

      function render() {
        if (!snapshot) {
          return;
        }
        clearCharts();
        if (state.page === 'annual') {
          renderAnnual(snapshot);
        } else {
          renderMonthly(snapshot);
        }
        setText('updatedAt', displayUpdated(snapshot.generatedAt));
        setText('coverageText', '数据来自 D1 预计算快照，销售额原始单位为万元，页面展示单位为亿元；月度覆盖 ' +
          (snapshot.coverage.monthly.minReportDate || '-') + ' 至 ' + (snapshot.coverage.monthly.maxReportDate || '-') + '。');
      }

      function populateControls(data) {
        var years = yearsFromSnapshot(data);
        if (!years.length) {
          return;
        }
        var latestYear = years[years.length - 1];
        if (!state.annualEnd || years.indexOf(state.annualEnd) < 0) {
          state.annualEnd = latestYear;
        }
        if (!state.annualStart || years.indexOf(state.annualStart) < 0) {
          state.annualStart = years[Math.max(0, years.length - 7)];
        }
        if (Number(state.annualEnd) - Number(state.annualStart) > 9) {
          state.annualStart = String(Number(state.annualEnd) - 9);
        }
        if (!state.monthlyYear || years.indexOf(state.monthlyYear) < 0) {
          state.monthlyYear = latestYear;
        }
        var monthlyOptions = (data.monthlyTrend || []).filter(function (row) {
          return dateYear(row.reportDate) === state.monthlyYear;
        }).map(function (row) {
          return { value: dateMonth(row.reportDate), label: Number(dateMonth(row.reportDate)) + '月' };
        });
        setOptions('annualStart', years.map(function (year) { return { value: year, label: year + '年' }; }), state.annualStart);
        setOptions('annualEnd', years.map(function (year) { return { value: year, label: year + '年' }; }), state.annualEnd);
        setOptions('monthlyYear', years.map(function (year) { return { value: year, label: year + '年' }; }), state.monthlyYear);
        setOptions('annualType', typeOptions(data), state.annualType);
        setOptions('monthlyType', typeOptions(data), state.monthlyType);
        setOptions('monthlyMonth', [{ value: 'all', label: '全年' }].concat(monthlyOptions), state.monthlyMonth);
      }

      function switchPage(page) {
        state.page = page;
        el('annualView').hidden = page !== 'annual';
        el('monthlyView').hidden = page !== 'monthly';
        el('annualControls').hidden = page !== 'annual';
        el('monthlyControls').hidden = page !== 'monthly';
        document.querySelectorAll('[data-page]').forEach(function (button) {
          button.classList.toggle('is-active', button.getAttribute('data-page') === page);
        });
        setText('pageMeta', page === 'annual' ? '年度数据分析' : '月度数据分析');
        setText('pageTitle', page === 'annual' ? '年度数据分析' : '月度数据分析');
        setText('pageSubtitle', page === 'annual'
          ? '支持连续年份区间对比分析，洞察彩票市场年度发展趋势与结构变化。'
          : '按月追踪全国销售走势、机构结构和增长变化，快速定位最新数据。');
        render();
      }

      function csvCell(value) {
        var text = String(value === null || value === undefined ? '' : value);
        return '"' + text.replace(/"/g, '""') + '"';
      }

      function downloadCsv(filename, rows) {
        var content = '\ufeff' + rows.map(function (row) {
          return row.map(csvCell).join(',');
        }).join('\\r\\n');
        var blob = new Blob([content], { type: 'text/csv;charset=utf-8' });
        var url = URL.createObjectURL(blob);
        var link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.click();
        URL.revokeObjectURL(url);
      }

      function exportCurrent() {
        if (!snapshot) {
          return;
        }
        if (state.page === 'annual') {
          var annual = filterYears(buildAnnualSeries(snapshot, state.annualType), state.annualStart, state.annualEnd);
          downloadCsv('lottery-annual-analysis.csv', [
            ['年份', '全国销售额（亿元）', '同比增长', '体彩销售额（亿元）', '福彩销售额（亿元）', '已入库月份'],
          ].concat(annual.map(function (row) {
            return [row.year, toYi(row.totalSales).toFixed(2), formatPercent(row.yearOverYearPercent),
              toYi(row.sportsSales).toFixed(2), toYi(row.welfareSales).toFixed(2), row.monthCount + ' / 12'];
          })));
        } else {
          var monthly = buildMonthlySeries(snapshot, state.monthlyType).filter(function (row) {
            return dateYear(row.reportDate) === state.monthlyYear;
          });
          downloadCsv('lottery-monthly-analysis.csv', [
            ['月份', '销售额（亿元）', '体彩销售额（亿元）', '福彩销售额（亿元）', '环比', '同比']
          ].concat(monthly.map(function (row) {
            return [displayMonth(row.reportDate), toYi(row.totalSales).toFixed(2), toYi(row.sportsSales).toFixed(2),
              toYi(row.welfareSales).toFixed(2), formatPercent(row.monthOverMonthPercent), formatPercent(row.yearOverYearPercent)];
          })));
        }
        showStatus('数据已导出为 CSV。', true);
      }

      function loadDashboard(rebuild) {
        showStatus(rebuild ? '正在刷新 D1 仪表盘快照…' : '正在读取仪表盘快照…', true);
        var path = rebuild ? '/api/dashboard/rebuild' : '/api/dashboard';
        return fetch(path + '?ts=' + Date.now(), { cache: 'no-store' })
          .then(function (response) {
            if (!response.ok) {
              throw new Error('HTTP ' + response.status);
            }
            return response.json();
          })
          .then(function (data) {
            snapshot = data;
            populateControls(data);
            switchPage(state.page);
            showStatus(rebuild ? '数据刷新完成。' : '', rebuild);
          })
          .catch(function (error) {
            showStatus('仪表盘读取失败：' + error.message, true);
          });
      }

      document.querySelectorAll('[data-page]').forEach(function (button) {
        button.addEventListener('click', function () {
          switchPage(button.getAttribute('data-page'));
        });
      });

      ['annualStart', 'annualEnd', 'annualType'].forEach(function (id) {
        el(id).addEventListener('change', function () {
          state[id] = el(id).value;
          if (id === 'annualStart' && Number(state.annualStart) > Number(state.annualEnd)) {
            state.annualEnd = state.annualStart;
            el('annualEnd').value = state.annualEnd;
          }
          if (id === 'annualEnd' && Number(state.annualEnd) < Number(state.annualStart)) {
            state.annualStart = state.annualEnd;
            el('annualStart').value = state.annualStart;
          }
          render();
        });
      });

      el('monthlyYear').addEventListener('change', function () {
        state.monthlyYear = el('monthlyYear').value;
        state.monthlyMonth = 'all';
        populateControls(snapshot);
        render();
      });
      el('monthlyMonth').addEventListener('change', function () {
        state.monthlyMonth = el('monthlyMonth').value;
        render();
      });
      el('monthlyType').addEventListener('change', function () {
        state.monthlyType = el('monthlyType').value;
        render();
      });
      document.querySelectorAll('[data-granularity]').forEach(function (button) {
        button.addEventListener('click', function () {
          state.granularity = button.getAttribute('data-granularity');
          document.querySelectorAll('[data-granularity]').forEach(function (item) {
            item.classList.toggle('is-active', item === button);
          });
          render();
        });
      });
      el('annualRefresh').addEventListener('click', function () { loadDashboard(true); });
      el('monthlyRefresh').addEventListener('click', function () { loadDashboard(true); });
      el('annualExport').addEventListener('click', exportCurrent);
      el('monthlyExport').addEventListener('click', exportCurrent);
      window.addEventListener('resize', function () {
        Object.keys(charts).forEach(function (key) { charts[key].resize(); });
      });

      loadDashboard(false);
    })();
  </script>
</body>
</html>`;

function htmlResponse(body: string, status = 200): Response {
  return new Response(body, {
    status,
    headers: {
      ...CORS_HEADERS,
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "no-store, max-age=0",
    },
  });
}

function jsonResponse(
  body: unknown,
  status = 200,
  headers: Record<string, string> = {},
): Response {
  return new Response(JSON.stringify(body, null, 2), {
    status,
    headers: {
      ...CORS_HEADERS,
      "Content-Type": "application/json; charset=utf-8",
      ...headers,
    },
  });
}

class HttpError extends Error {
  status: number;

  constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}

function statusForError(error: unknown): number {
  return error instanceof HttpError ? error.status : 500;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function serializeError(error: unknown): Record<string, string> {
  if (error instanceof Error) {
    return { name: error.name, message: error.message };
  }
  return { message: String(error) };
}
