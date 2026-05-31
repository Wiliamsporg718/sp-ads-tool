/**
 * Bulk Sheet Generator — Amazon Sponsored Products Manual Campaign
 *
 * Migrated from VBA: AD Detail(6)(1).xlsm
 * Supports two modes:
 *   - keyword: AD Detail Keywords → keyword entity rows
 *   - asin:    AD Detail ASIN    → Product targeting entity rows
 *
 * Architecture:
 *   AdDetailRow[] (flat input) → generateBulkSheet() → BulkSheetRow[] (hierarchical output)
 */

// ---------------------------------------------------------------------------
// Input types — maps to "AD Detail Keywords" sheet columns
// ---------------------------------------------------------------------------

export interface AdDetailRow {
  /** Column A: Campaign Name */
  campaignName: string;
  /** Column B: Ad Group Name */
  adGroupName: string;
  /** Column C: Daily Budget */
  budget: number;
  /** Column D: SKU (may be empty) */
  sku: string;
  /** Column E: ASIN (may be empty) */
  asin: string;
  /** Column F: Keyword Bid (may be empty) */
  bid: number | null;
  /** Column G: Keyword Text (may be empty) */
  keywordText: string;
  /** Column H: Match Type — exact / phrase / broad (may be empty) */
  matchType: string;
  /** Column I: Bidding Strategy (may be empty) */
  biddingStrategy: string;
  /** Column J: Campaign Start Date — YYYYMMDD (optional, overrides options.startDate) */
  startDate: string;
  /** Whether this is a negative targeting entry (optional, defaults to false) */
  isNegative?: boolean;
}

// ---------------------------------------------------------------------------
// Output types — Amazon Bulk Sheet columns (Sponsored Products)
// ---------------------------------------------------------------------------

export type BulkEntityType =
  | "Campaign"
  | "Bidding adjustment"
  | "Ad group"
  | "Product ad"
  | "keyword"
  | "Product targeting"
  | "negative keyword"
  | "negative product targeting";

export type PlacementType =
  | "Placement top"
  | "Placement product page"
  | "Placement rest of search"
  | "Placement Amazon Business";

export interface BulkSheetRow {
  /** Col A: Product */
  product: string;
  /** Col B: Entity */
  entity: BulkEntityType;
  /** Col C: Operation */
  operation: string;
  /** Col D: Campaign Id (campaign name as placeholder) */
  campaignId: string;
  /** Col E: Ad Group Id (ad group name as placeholder) */
  adGroupId: string;
  /** Col J: Campaign Name */
  campaignName: string;
  /** Col K: Ad Group Name */
  adGroupName: string;
  /** Col L: Campaign Name (Informational only) */
  campaignNameInfo: string;
  /** Col M: Ad Group Name (Informational only) */
  adGroupNameInfo: string;
  /** Col O: Start Date (YYYYMMDD) */
  startDate: string;
  /** Col Q: Targeting Type */
  targetingType: string;
  /** Col R: Campaign Status */
  campaignStatus: string;
  /** Col S: Campaign Status (secondary) */
  campaignStatusSecondary: string;
  /** Col T: Ad Group Status */
  adGroupStatus: string;
  /** Col U: Budget */
  budget: string;
  /** Col V: SKU */
  sku: string;
  /** Col W: ASIN (Informational only) */
  asinInfo: string;
  /** Col X: Eligibility Status (Informational only) */
  eligibilityStatus: string;
  /** Col Z: Ad Group Default Bid */
  adGroupDefaultBid: string;
  /** Col AA: Ad Group Default Bid (Informational only) */
  adGroupDefaultBidInfo: string;
  /** Col AB: Bid */
  keywordBid: string;
  /** Col AC: Keyword Text */
  keywordText: string;
  /** Col AF: Match Type */
  matchType: string;
  /** Col AG: Bidding Strategy */
  biddingStrategy: string;
  /** Col AH: Placement */
  placement: string;
  /** Col AI: Percentage */
  percentage: string;
  /** Col AJ: Product Targeting Expression (ASIN mode only) */
  productTargetingExpression: string;
  /** Col AK: Resolved Product Targeting Expression (ASIN mode only) */
  resolvedProductTargetingExpression: string;
}

// ---------------------------------------------------------------------------
// Generator options
// ---------------------------------------------------------------------------

export interface BulkSheetOptions {
  /** Default ad group bid (VBA hardcoded "0.50") */
  defaultAdGroupBid?: string;
  /** Default bidding adjustment percentage (VBA hardcoded "0.00") */
  defaultPlacementPercentage?: string;
  /** Start date override (defaults to today YYYYMMDD) */
  startDate?: string;
  /** Targeting type — "manual" or "auto" */
  targetingType?: "manual" | "auto";
  /** Targeting mode — "keyword" for keyword targeting, "asin" for product targeting */
  mode?: "keyword" | "asin";
}

// ---------------------------------------------------------------------------
// Bulk Sheet Header — Amazon official column headers
// ---------------------------------------------------------------------------

export const BULK_SHEET_HEADERS = [
  "Product",                          // A (0)
  "Entity",                           // B (1)
  "Operation",                        // C (2)
  "Campaign Id",                      // D (3)
  "Ad Group Id",                      // E (4)
  "Portfolio Id",                     // F (5)
  "Ad Id",                            // G (6)
  "Keyword Id",                       // H (7)
  "Product Targeting Id",             // I (8)
  "Campaign Name",                    // J (9)
  "Ad Group Name",                    // K (10)
  "Campaign Name (Informational only)", // L (11)
  "Ad Group Name (Informational only)", // M (12)
  "Portfolio Name (Informational only)", // N (13)
  "Start Date",                       // O (14)
  "End Date",                         // P (15)
  "Targeting Type",                   // Q (16)
  "State",                            // R (17)
  "Campaign State (Informational only)", // S (18)
  "Ad Group State (Informational only)", // T (19)
  "Daily Budget",                     // U (20)
  "SKU",                              // V (21)
  "ASIN (Informational only)",        // W (22)
  "Eligibility Status (Informational only)", // X (23)
  "Reason for ineligibility (Informational only)", // Y (24)
  "Ad Group Default Bid",             // Z (25)
  "Ad Group Default Bid (Informational only)", // AA (26)
  "Bid",                              // AB (27)
  "Keyword Text",                     // AC (28)
  "Native Language Keyword (Informational only)", // AD (29)
  "Native Language Negative Keyword (Informational only)", // AE (30)
  "Match Type",                       // AF (31)
  "Bidding Strategy",                 // AG (32)
  "Placement",                        // AH (33)
  "Percentage",                       // AI (34)
  "Product Targeting Expression",     // AJ (35)
  "Resolved Product Targeting Expression (Informational only)", // AK (36)
] as const;

const PLACEMENT_TYPES: PlacementType[] = [
  "Placement top",
  "Placement product page",
  "Placement rest of search",
  "Placement Amazon Business",
];

// ---------------------------------------------------------------------------
// Core generator — pure function, no side effects
// ---------------------------------------------------------------------------

export function generateBulkSheet(
  input: AdDetailRow[],
  options: BulkSheetOptions = {},
): BulkSheetRow[] {
  const {
    defaultAdGroupBid = "0.50",
    defaultPlacementPercentage = "0.00",
    startDate = formatDate(new Date()),
    targetingType = "manual",
    mode = "keyword",
  } = options;

  if (input.length === 0) return [];

  const result: BulkSheetRow[] = [];
  let i = 0;
  let prevCampaign = "";
  let currentBiddingStrategy = "";

  while (i < input.length) {
    const row = input[i];
    if (!row.campaignName) { i++; continue; }

    const campaignName = row.campaignName;
    const adGroupName = row.adGroupName;

    // --- Collect all rows belonging to this campaign + ad group ---
    const groupRows: AdDetailRow[] = [];
    let j = i;
    while (j < input.length && input[j].campaignName === campaignName && input[j].adGroupName === adGroupName) {
      groupRows.push(input[j]);
      j++;
    }

    // --- Determine bidding strategy from group ---
    const strategies = new Set<string>();
    for (const r of groupRows) {
      if (r.biddingStrategy?.trim()) strategies.add(r.biddingStrategy.trim());
    }
    currentBiddingStrategy = [...strategies].join(", ");

    // --- Determine start date: prefer per-row value, then options ---
    const rowStartDate = row.startDate || startDate;

    // --- Campaign row (only when campaign changes) ---
    if (campaignName !== prevCampaign) {
      // Campaign row
      result.push(createRow({
        entity: "Campaign",
        campaignId: campaignName,
        campaignName: campaignName,
        campaignNameInfo: campaignName,
        startDate: rowStartDate,
        targetingType,
        campaignStatus: "enabled",
        campaignStatusSecondary: "enabled",
        budget: formatBudget(row.budget),
        biddingStrategy: currentBiddingStrategy,
      }));

      // 4 Bidding adjustment rows
      for (const placement of PLACEMENT_TYPES) {
        result.push(createRow({
          entity: "Bidding adjustment",
          campaignId: campaignName,
          campaignNameInfo: campaignName,
          campaignStatusSecondary: "enabled",
          biddingStrategy: currentBiddingStrategy,
          placement,
          percentage: defaultPlacementPercentage,
        }));
      }

      prevCampaign = campaignName;
    }

    // --- Ad group row ---
    result.push(createRow({
      entity: "Ad group",
      campaignId: campaignName,
      adGroupId: adGroupName,
      adGroupName: adGroupName,
      campaignNameInfo: campaignName,
      adGroupNameInfo: adGroupName,
      campaignStatus: "enabled",
      campaignStatusSecondary: "enabled",
      adGroupStatus: "enabled",
      adGroupDefaultBid: defaultAdGroupBid,
    }));

    // --- Product ad rows (one per SKU) ---
    const skuRows = groupRows.filter(r => r.sku);
    for (const skuRow of skuRows) {
      result.push(createRow({
        entity: "Product ad",
        campaignId: campaignName,
        adGroupId: adGroupName,
        campaignNameInfo: campaignName,
        adGroupNameInfo: adGroupName,
        campaignStatus: "enabled",
        campaignStatusSecondary: "enabled",
        adGroupStatus: "enabled",
        sku: skuRow.sku,
        asinInfo: skuRow.asin || "",
        eligibilityStatus: "Eligible",
      }));
    }

    // --- Keyword / Product targeting rows ---
    if (mode === "keyword") {
      // Positive keyword rows
      const keywordRows = groupRows.filter(r => r.keywordText && !r.isNegative);
      for (const kwRow of keywordRows) {
        result.push(createRow({
          entity: "keyword",
          campaignId: campaignName,
          adGroupId: adGroupName,
          campaignNameInfo: campaignName,
          adGroupNameInfo: adGroupName,
          campaignStatus: "enabled",
          campaignStatusSecondary: "enabled",
          adGroupStatus: "enabled",
          adGroupDefaultBidInfo: defaultAdGroupBid,
          keywordBid: kwRow.bid != null ? kwRow.bid.toFixed(2) : "",
          keywordText: kwRow.keywordText,
          matchType: kwRow.matchType,
        }));
      }

      // Negative keyword rows
      const negKeywordRows = groupRows.filter(r => r.keywordText && r.isNegative);
      for (const nkRow of negKeywordRows) {
        result.push(createRow({
          entity: "negative keyword",
          campaignId: campaignName,
          adGroupId: adGroupName,
          campaignNameInfo: campaignName,
          adGroupNameInfo: adGroupName,
          campaignStatus: "enabled",
          campaignStatusSecondary: "enabled",
          adGroupStatus: "enabled",
          keywordText: nkRow.keywordText,
          matchType: nkRow.matchType || "negative exact",
        }));
      }
    } else {
      // Positive ASIN targeting rows
      const targetingRows = groupRows.filter(r => r.keywordText && !r.isNegative);
      for (const tRow of targetingRows) {
        const expr = `asin="${tRow.keywordText}"`;
        result.push(createRow({
          entity: "Product targeting",
          campaignId: campaignName,
          adGroupId: adGroupName,
          campaignNameInfo: campaignName,
          adGroupNameInfo: adGroupName,
          campaignStatus: "enabled",
          campaignStatusSecondary: "enabled",
          adGroupStatus: "enabled",
          adGroupDefaultBidInfo: defaultAdGroupBid,
          keywordBid: tRow.bid != null ? tRow.bid.toFixed(2) : "",
          productTargetingExpression: expr,
          resolvedProductTargetingExpression: expr,
        }));
      }

      // Negative ASIN targeting rows
      const negTargetingRows = groupRows.filter(r => r.keywordText && r.isNegative);
      for (const ntRow of negTargetingRows) {
        const expr = `asin="${ntRow.keywordText}"`;
        result.push(createRow({
          entity: "negative product targeting",
          campaignId: campaignName,
          adGroupId: adGroupName,
          campaignNameInfo: campaignName,
          adGroupNameInfo: adGroupName,
          campaignStatus: "enabled",
          campaignStatusSecondary: "enabled",
          adGroupStatus: "enabled",
          productTargetingExpression: expr,
          resolvedProductTargetingExpression: expr,
        }));
      }
    }

    i = j;
  }

  return result;
}

// ---------------------------------------------------------------------------
// Row-to-array conversion (for XLSX writing)
// ---------------------------------------------------------------------------

export function bulkRowToArray(row: BulkSheetRow): (string | number)[] {
  return [
    row.product,              // A
    row.entity,               // B
    row.operation,            // C
    row.campaignId,           // D
    row.adGroupId,            // E
    "",                       // F: Portfolio Id
    "",                       // G: Ad Id
    "",                       // H: Keyword Id
    "",                       // I: Product Targeting Id
    row.campaignName,         // J
    row.adGroupName,          // K
    row.campaignNameInfo,     // L
    row.adGroupNameInfo,      // M
    "",                       // N: Portfolio Name
    row.startDate,            // O
    "",                       // P: End Date
    row.targetingType,        // Q
    row.campaignStatus,       // R
    row.campaignStatusSecondary, // S
    row.adGroupStatus,        // T
    row.budget,               // U
    row.sku,                  // V
    row.asinInfo,             // W
    row.eligibilityStatus,    // X
    "",                       // Y: Reason for ineligibility
    row.adGroupDefaultBid,    // Z
    row.adGroupDefaultBidInfo, // AA
    row.keywordBid,           // AB
    row.keywordText,          // AC
    "",                       // AD: Native Language Keyword
    "",                       // AE: Native Language Negative Keyword
    row.matchType,            // AF
    row.biddingStrategy,      // AG
    row.placement,            // AH
    row.percentage,           // AI
    row.productTargetingExpression, // AJ
    row.resolvedProductTargetingExpression, // AK
  ];
}

// ---------------------------------------------------------------------------
// Statistics
// ---------------------------------------------------------------------------

export interface BulkSheetStats {
  totalRows: number;
  campaignCount: number;
  adGroupCount: number;
  productAdCount: number;
  keywordCount: number;
  productTargetingCount: number;
  negativeKeywordCount: number;
  negativeProductTargetingCount: number;
  biddingAdjustmentCount: number;
}

export function getBulkSheetStats(rows: BulkSheetRow[]): BulkSheetStats {
  return {
    totalRows: rows.length,
    campaignCount: rows.filter(r => r.entity === "Campaign").length,
    adGroupCount: rows.filter(r => r.entity === "Ad group").length,
    productAdCount: rows.filter(r => r.entity === "Product ad").length,
    keywordCount: rows.filter(r => r.entity === "keyword").length,
    productTargetingCount: rows.filter(r => r.entity === "Product targeting").length,
    negativeKeywordCount: rows.filter(r => r.entity === "negative keyword").length,
    negativeProductTargetingCount: rows.filter(r => r.entity === "negative product targeting").length,
    biddingAdjustmentCount: rows.filter(r => r.entity === "Bidding adjustment").length,
  };
}

// ---------------------------------------------------------------------------
// AD Detail parser — reads flat rows from uploaded Excel
// ---------------------------------------------------------------------------

export function parseAdDetailRows(sheetData: unknown[][]): AdDetailRow[] {
  // Skip header row (index 0), parse from row 1 onward
  const rows: AdDetailRow[] = [];
  for (let i = 1; i < sheetData.length; i++) {
    const r = sheetData[i];
    if (!r || !r[0]) continue; // skip empty rows
    rows.push({
      campaignName: String(r[0] ?? "").trim(),
      adGroupName: String(r[1] ?? "").trim(),
      budget: parseFloat(String(r[2] ?? "0")) || 0,
      sku: String(r[3] ?? "").trim(),
      asin: String(r[4] ?? "").trim(),
      bid: r[5] != null && String(r[5]).trim() !== "" ? parseFloat(String(r[5])) : null,
      keywordText: String(r[6] ?? "").trim(),
      matchType: String(r[7] ?? "").trim(),
      biddingStrategy: String(r[8] ?? "").trim(),
      startDate: r[9] != null ? String(r[9]).trim() : "",
    });
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

export interface ValidationError {
  row: number;
  field: string;
  message: string;
}

/** Amazon ASIN format: starts with B0 and is 10 characters */
const ASIN_PATTERN = /^B[A-Z0-9]{9}$/i;

export function validateAdDetailRows(
  rows: AdDetailRow[],
  options?: { mode?: "keyword" | "asin" },
): ValidationError[] {
  const mode = options?.mode ?? "keyword";
  const errors: ValidationError[] = [];
  const seenCombos = new Set<string>();

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rowNum = i + 2; // 1-indexed + header

    // --- Required fields ---
    if (!r.campaignName) {
      errors.push({ row: rowNum, field: "campaignName", message: "广告活动名称不能为空" });
    }
    if (!r.adGroupName) {
      errors.push({ row: rowNum, field: "adGroupName", message: "广告组名称不能为空" });
    }

    // --- Length limits ---
    if (r.campaignName && r.campaignName.length > 128) {
      errors.push({ row: rowNum, field: "campaignName", message: "广告活动名称不能超过128字符" });
    }
    if (r.adGroupName && r.adGroupName.length > 255) {
      errors.push({ row: rowNum, field: "adGroupName", message: "广告组名称不能超过255字符" });
    }

    // --- Budget ---
    if (r.budget && r.budget < 1) {
      errors.push({ row: rowNum, field: "budget", message: "日预算不能低于 $1" });
    }
    if (r.budget && r.budget > 1000000) {
      errors.push({ row: rowNum, field: "budget", message: "日预算不能超过 $1,000,000" });
    }

    // --- Bid ---
    if (r.bid != null && (r.bid < 0.02 || r.bid > 1000)) {
      errors.push({ row: rowNum, field: "bid", message: "竞价必须在 $0.02 - $1000 之间" });
    }

    // --- ASIN format ---
    if (r.asin && !ASIN_PATTERN.test(r.asin)) {
      errors.push({ row: rowNum, field: "asin", message: `ASIN 格式无效: "${r.asin}"，应为 B0 开头的10位编码` });
    }

    // --- Mode-specific validation ---
    if (mode === "keyword") {
      // Match type validation
      if (r.matchType && !["exact", "phrase", "broad"].includes(r.matchType.toLowerCase())) {
        errors.push({ row: rowNum, field: "matchType", message: `匹配类型无效: "${r.matchType}"，应为 exact/phrase/broad` });
      }
    } else {
      // ASIN mode: keywordText field holds competitor ASIN
      if (r.keywordText && !ASIN_PATTERN.test(r.keywordText)) {
        errors.push({ row: rowNum, field: "competitorAsin", message: `竞品ASIN 格式无效: "${r.keywordText}"，应为 B0 开头的10位编码` });
      }

      // Duplicate detection: same ad group + same competitor ASIN
      if (r.adGroupName && r.keywordText) {
        const key = `${r.adGroupName}||${r.keywordText}`;
        if (seenCombos.has(key)) {
          errors.push({ row: rowNum, field: "competitorAsin", message: `重复投放: 同一广告组 "${r.adGroupName}" 已包含 ASIN "${r.keywordText}"` });
        }
        seenCombos.add(key);
      }
    }
  }
  return errors;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function createRow(overrides: Partial<BulkSheetRow>): BulkSheetRow {
  return {
    product: "Sponsored Products",
    entity: "Campaign",
    operation: "Create",
    campaignId: "",
    adGroupId: "",
    campaignName: "",
    adGroupName: "",
    campaignNameInfo: "",
    adGroupNameInfo: "",
    startDate: "",
    targetingType: "",
    campaignStatus: "",
    campaignStatusSecondary: "",
    adGroupStatus: "",
    budget: "",
    sku: "",
    asinInfo: "",
    eligibilityStatus: "",
    adGroupDefaultBid: "",
    adGroupDefaultBidInfo: "",
    keywordBid: "",
    keywordText: "",
    matchType: "",
    biddingStrategy: "",
    placement: "",
    percentage: "",
    productTargetingExpression: "",
    resolvedProductTargetingExpression: "",
    ...overrides,
  };
}

function formatDate(date: Date): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}${m}${d}`;
}

function formatBudget(value: number): string {
  return value ? value.toFixed(2) : "";
}
