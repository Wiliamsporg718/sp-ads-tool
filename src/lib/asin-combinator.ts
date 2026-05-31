/**
 * ASIN Combinator — Data Preparation for ASIN Targeting Campaigns
 *
 * Migrated from VBA: AD Detail(6)(1).xlsm → CommandButton1 + CommandButton2
 *
 * Two functions:
 *   1. generateAsinPermutations() — Creates all N×(N-1) product targeting combinations
 *   2. fixAdGroupNames()          — Replaces RES(xxx) in ad group names with current SKU
 */

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface AsinProduct {
  /** Column A: SKU */
  sku: string;
  /** Column B: ASIN */
  asin: string;
}

export interface AsinCombinationRow {
  /** Column D: Source SKU (only filled on first row of each group) */
  sku: string;
  /** Column E: Source ASIN (only filled on first row of each group) */
  asin: string;
  /** Column G: Competitor ASIN to target */
  competitorAsin: string;
}

export interface AdDetailAsinRow {
  /** Column A: Campaign Name */
  campaignName: string;
  /** Column B: Ad Group Name */
  adGroupName: string;
  /** Column C: Daily Budget */
  budget: number;
  /** Column D: SKU */
  sku: string;
  /** Column E: ASIN */
  asin: string;
  /** Column F: Bid */
  bid: number;
  /** Column G: Competitor ASIN */
  competitorAsin: string;
  /** Column H: Bidding Strategy */
  biddingStrategy: string;
  /** Column I: Start Date (YYYYMMDD) */
  startDate: string;
}

// ---------------------------------------------------------------------------
// CommandButton1 — ASIN Permutation Generator
// ---------------------------------------------------------------------------

/**
 * Generate all N×(N-1) permutations from a product list.
 * Each product targets every OTHER product's ASIN (excludes self).
 * D/E columns only filled on first row of each source product group.
 */
export function generateAsinPermutations(products: AsinProduct[]): AsinCombinationRow[] {
  // Filter valid products (both SKU and ASIN non-empty)
  const valid = products.filter(p => p.sku.trim() !== "" && p.asin.trim() !== "");

  if (valid.length < 2) return [];

  const result: AsinCombinationRow[] = [];

  for (let i = 0; i < valid.length; i++) {
    let isFirst = true;
    for (let j = 0; j < valid.length; j++) {
      if (i === j) continue; // skip self

      result.push({
        sku: isFirst ? valid[i].sku : "",
        asin: isFirst ? valid[i].asin : "",
        competitorAsin: valid[j].asin,
      });
      isFirst = false;
    }
  }

  return result;
}

// ---------------------------------------------------------------------------
// P0: Offensive ASIN Targeting — target competitor ASINs
// ---------------------------------------------------------------------------

export interface CompetitorAsin {
  /** Competitor ASIN to target */
  asin: string;
  /** Optional label for identification */
  label?: string;
}

/**
 * Generate offensive targeting combinations:
 * Each own product targets every competitor ASIN.
 * Result: ownProducts.length × competitors.length combinations.
 * D/E columns only filled on first row of each own-product group.
 */
export function generateOffensiveTargeting(
  ownProducts: AsinProduct[],
  competitors: CompetitorAsin[],
): AsinCombinationRow[] {
  const validOwn = ownProducts.filter(p => p.sku.trim() !== "" && p.asin.trim() !== "");
  const validComp = competitors.filter(c => c.asin.trim() !== "");

  if (validOwn.length === 0 || validComp.length === 0) return [];

  const result: AsinCombinationRow[] = [];

  for (let i = 0; i < validOwn.length; i++) {
    let isFirst = true;
    for (let j = 0; j < validComp.length; j++) {
      // Skip if competitor ASIN is same as own ASIN
      if (validComp[j].asin === validOwn[i].asin) continue;

      result.push({
        sku: isFirst ? validOwn[i].sku : "",
        asin: isFirst ? validOwn[i].asin : "",
        competitorAsin: validComp[j].asin,
      });
      isFirst = false;
    }
  }

  return result;
}

// ---------------------------------------------------------------------------
// P1: Grouped ASIN Permutations — only cross-target within groups
// ---------------------------------------------------------------------------

export interface GroupedAsinProduct extends AsinProduct {
  /** Group identifier (e.g., parent ASIN, category, gender) */
  group: string;
}

/**
 * Generate permutations only within the same group.
 * Products in different groups will NOT target each other.
 * Useful for avoiding irrelevant cross-targeting (e.g., men's XS → women's shorts).
 */
export function generateGroupedPermutations(
  products: GroupedAsinProduct[],
): AsinCombinationRow[] {
  const valid = products.filter(p => p.sku.trim() !== "" && p.asin.trim() !== "" && p.group.trim() !== "");

  // Group by group identifier
  const groups = new Map<string, GroupedAsinProduct[]>();
  for (const p of valid) {
    const existing = groups.get(p.group) || [];
    existing.push(p);
    groups.set(p.group, existing);
  }

  const result: AsinCombinationRow[] = [];

  for (const [, groupProducts] of groups) {
    if (groupProducts.length < 2) continue;

    // Generate N×(N-1) within this group only
    for (let i = 0; i < groupProducts.length; i++) {
      let isFirst = true;
      for (let j = 0; j < groupProducts.length; j++) {
        if (i === j) continue;

        result.push({
          sku: isFirst ? groupProducts[i].sku : "",
          asin: isFirst ? groupProducts[i].asin : "",
          competitorAsin: groupProducts[j].asin,
        });
        isFirst = false;
      }
    }
  }

  return result;
}

/**
 * Build complete AD Detail ASIN rows by combining permutations with defaults.
 * If naming config is provided, auto-generates campaign and ad group names.
 */
export function buildAdDetailAsinRows(
  combinations: AsinCombinationRow[],
  defaults: {
    budget: number;
    bid: number;
    biddingStrategy: string;
    startDate?: string;
  },
  naming?: CampaignNamingConfig,
): AdDetailAsinRow[] {
  const startDate = defaults.startDate || formatDate(new Date());
  const dateStr = naming?.date || formatDateDotted(new Date());

  return combinations.map(c => {
    let campaignName = "";
    let adGroupName = "";

    if (naming && c.sku) {
      campaignName = generateCampaignName(c.sku, naming, dateStr);
      adGroupName = generateAdGroupName(c.sku, naming, dateStr);
    }

    return {
      campaignName,
      adGroupName,
      budget: defaults.budget,
      sku: c.sku,
      asin: c.asin,
      bid: defaults.bid,
      competitorAsin: c.competitorAsin,
      biddingStrategy: defaults.biddingStrategy,
      startDate,
    };
  });
}

// ---------------------------------------------------------------------------
// Campaign Naming Generator
// ---------------------------------------------------------------------------

export interface CampaignNamingConfig {
  /** Brand name (e.g., "AGU") */
  brand: string;
  /** Product category (e.g., "Jersey") */
  category: string;
  /** Campaign type: DEF (defense) / OFF (offense) */
  campaignType?: "DEF" | "OFF" | string;
  /** Date string override (e.g., "2026.05.25"), defaults to today */
  date?: string;
}

/**
 * Generate campaign name from template.
 * Pattern: {Brand} - {Category} - ASIN {Type} - RES({SKU}) - {Date}
 */
export function generateCampaignName(
  sku: string,
  config: CampaignNamingConfig,
  date?: string,
): string {
  const type = config.campaignType || "DEF";
  const d = date || config.date || formatDateDotted(new Date());
  return `${config.brand} - ${config.category} - ASIN ${type} - RES(${sku}) - ${d}`;
}

/**
 * Generate ad group name from template.
 * Pattern: {Brand} - {Category} - ASIN {Type} - RES({SKU}) - {Date}
 */
export function generateAdGroupName(
  sku: string,
  config: CampaignNamingConfig,
  date?: string,
): string {
  const type = config.campaignType || "DEF";
  const d = date || config.date || formatDateDotted(new Date());
  return `${config.brand} - ${config.category} - ASIN ${type} - RES(${sku}) - ${d}`;
}

// ---------------------------------------------------------------------------
// CommandButton2 — Ad Group Name Fixer
// ---------------------------------------------------------------------------

/**
 * Replace `RES(xxx)` portion in ad group names with `RES(currentSKU)`.
 * Tracks the current SKU value — updates whenever column D has a value.
 * Returns a new array (does not mutate input).
 */
export function fixAdGroupNames(rows: AdDetailAsinRow[]): AdDetailAsinRow[] {
  let currentSku = "";

  return rows.map(row => {
    // Update current SKU when D column has a value
    if (row.sku.trim() !== "") {
      currentSku = row.sku;
    }

    // Only replace if ad group name exists and current SKU is set
    if (row.adGroupName.trim() === "" || currentSku === "") {
      return row;
    }

    const newAdGroupName = replaceResInName(row.adGroupName, currentSku);

    return newAdGroupName !== row.adGroupName
      ? { ...row, adGroupName: newAdGroupName }
      : row;
  });
}

/**
 * Replace `RES(xxx)` in a string with `RES(newValue)`.
 */
function replaceResInName(name: string, newValue: string): string {
  const startPos = name.indexOf("RES(");
  if (startPos === -1) return name;

  const endPos = name.indexOf(")", startPos);
  if (endPos === -1 || endPos <= startPos) return name;

  const oldRes = name.substring(startPos, endPos + 1);
  const newRes = `RES(${newValue})`;

  return name.replace(oldRes, newRes);
}

// ---------------------------------------------------------------------------
// XLSX parsing helpers
// ---------------------------------------------------------------------------

/**
 * Parse ASIN Data sheet (columns A: SKU, B: ASIN) into AsinProduct[].
 */
export function parseAsinDataSheet(sheetData: unknown[][]): AsinProduct[] {
  const products: AsinProduct[] = [];

  for (let i = 0; i < sheetData.length; i++) {
    const r = sheetData[i];
    if (!r) continue;

    const sku = String(r[0] ?? "").trim();
    const asin = String(r[1] ?? "").trim();

    if (sku !== "" && asin !== "") {
      products.push({ sku, asin });
    }
  }

  return products;
}

/**
 * Parse AD Detail ASIN sheet (9 columns) into AdDetailAsinRow[].
 */
export function parseAdDetailAsinSheet(sheetData: unknown[][]): AdDetailAsinRow[] {
  const rows: AdDetailAsinRow[] = [];

  // Skip header row (index 0)
  for (let i = 1; i < sheetData.length; i++) {
    const r = sheetData[i];
    if (!r || !r[0]) continue;

    rows.push({
      campaignName: String(r[0] ?? "").trim(),
      adGroupName: String(r[1] ?? "").trim(),
      budget: parseFloat(String(r[2] ?? "0")) || 0,
      sku: String(r[3] ?? "").trim(),
      asin: String(r[4] ?? "").trim(),
      bid: parseFloat(String(r[5] ?? "0")) || 0,
      competitorAsin: String(r[6] ?? "").trim(),
      biddingStrategy: String(r[7] ?? "").trim(),
      startDate: r[8] != null ? String(r[8]).trim() : "",
    });
  }

  return rows;
}

// ---------------------------------------------------------------------------
// Statistics
// ---------------------------------------------------------------------------

export interface AsinCombinatorStats {
  /** Number of unique source products */
  productCount: number;
  /** Total combinations generated */
  combinationCount: number;
  /** Expected: N × (N-1) */
  expectedCombinations: number;
  /** Warning if product count exceeds threshold */
  warning?: string;
}

/** Thresholds for combination count warnings */
const WARN_THRESHOLD = 100;  // > 100 products → 9,900+ combos
const MAX_THRESHOLD = 500;   // > 500 products → 249,500+ combos

export function getPermutationStats(
  products: AsinProduct[],
  combinations: AsinCombinationRow[],
): AsinCombinatorStats {
  const valid = products.filter(p => p.sku.trim() !== "" && p.asin.trim() !== "");
  const stats: AsinCombinatorStats = {
    productCount: valid.length,
    combinationCount: combinations.length,
    expectedCombinations: valid.length * (valid.length - 1),
  };

  if (valid.length > MAX_THRESHOLD) {
    stats.warning = `产品数量 (${valid.length}) 超过 ${MAX_THRESHOLD}，将生成 ${stats.expectedCombinations.toLocaleString()} 个组合，可能超出 Amazon Bulk Upload 限制`;
  } else if (valid.length > WARN_THRESHOLD) {
    stats.warning = `产品数量较多 (${valid.length})，将生成 ${stats.expectedCombinations.toLocaleString()} 个组合，建议分批上传`;
  }

  return stats;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function formatDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}${m}${day}`;
}

function formatDateDotted(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}.${m}.${day}`;
}
