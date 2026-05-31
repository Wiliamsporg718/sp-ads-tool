/**
 * Auto Campaign Generator — Amazon Sponsored Products Auto Campaigns
 *
 * Generates Bulk Sheet data for SP Auto campaigns with 4 targeting groups:
 *   1. Close Match  — Ads appear when search terms closely match your products
 *   2. Loose Match  — Ads appear when search terms loosely relate to your products
 *   3. Substitutes  — Ads appear on detail pages of similar products
 *   4. Complements  — Ads appear on detail pages of complementary products
 *
 * Standard workflow:
 *   1. Create Auto campaigns to discover search terms and ASINs
 *   2. Analyze Search Term Report to find high-performing terms
 *   3. Move winning terms to Manual Exact campaigns (harvest)
 *   4. Add those terms as negatives in Auto to avoid duplicate spend
 */

import type { BulkSheetRow } from "./bulk-sheet-generator";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface AutoCampaignProduct {
  /** Product SKU */
  sku: string;
  /** Product ASIN */
  asin: string;
}

export interface AutoCampaignConfig {
  /** Brand name for campaign naming */
  brand: string;
  /** Product category for campaign naming */
  category: string;
  /** Daily budget per campaign */
  budget: number;
  /** Default bid for all targeting groups */
  defaultBid: number;
  /** Individual bids per targeting group (overrides defaultBid) */
  groupBids?: {
    closeMatch?: number;
    looseMatch?: number;
    substitutes?: number;
    complements?: number;
  };
  /** Bidding strategy */
  biddingStrategy?: string;
  /** Start date (YYYYMMDD), defaults to today */
  startDate?: string;
  /** Campaign date string for naming (e.g., "2026.05.25") */
  namingDate?: string;
  /** Whether to create one campaign per product or one campaign for all */
  structure?: "per-product" | "single";
}

const AUTO_TARGETING_GROUPS = [
  { key: "closeMatch", name: "Close Match", label: "close" },
  { key: "looseMatch", name: "Loose Match", label: "loose" },
  { key: "substitutes", name: "Substitutes", label: "substitutes" },
  { key: "complements", name: "Complements", label: "complements" },
] as const;

type GroupKey = typeof AUTO_TARGETING_GROUPS[number]["key"];

// ---------------------------------------------------------------------------
// Generator
// ---------------------------------------------------------------------------

/**
 * Generate Bulk Sheet rows for Auto campaigns.
 * Creates one campaign per product (default) or one combined campaign.
 */
export function generateAutoCampaigns(
  products: AutoCampaignProduct[],
  config: AutoCampaignConfig,
): BulkSheetRow[] {
  const validProducts = products.filter(p => p.sku.trim() !== "" && p.asin.trim() !== "");
  if (validProducts.length === 0) return [];

  const startDate = config.startDate || formatDate(new Date());
  const namingDate = config.namingDate || formatDateDotted(new Date());
  const strategy = config.biddingStrategy || "Dynamic bids - down only";
  const structure = config.structure || "per-product";

  const result: BulkSheetRow[] = [];

  if (structure === "per-product") {
    // One campaign per product, one ad group per targeting group
    for (const product of validProducts) {
      const campaignName = `${config.brand} - ${config.category} - SP Auto - ${product.sku} - ${namingDate}`;

      // Campaign row
      result.push(createRow({
        entity: "Campaign",
        campaignId: campaignName,
        campaignName: campaignName,
        campaignNameInfo: campaignName,
        startDate,
        targetingType: "auto",
        campaignStatus: "enabled",
        campaignStatusSecondary: "enabled",
        budget: config.budget.toFixed(2),
        biddingStrategy: strategy,
      }));

      // 4 Bidding adjustment rows
      for (const placement of PLACEMENTS) {
        result.push(createRow({
          entity: "Bidding adjustment",
          campaignId: campaignName,
          campaignNameInfo: campaignName,
          campaignStatusSecondary: "enabled",
          biddingStrategy: strategy,
          placement,
          percentage: "0.00",
        }));
      }

      // 4 Ad groups (one per targeting group)
      for (const group of AUTO_TARGETING_GROUPS) {
        const adGroupName = `${campaignName} - ${group.name}`;
        const groupBid = config.groupBids?.[group.key as GroupKey] ?? config.defaultBid;

        // Ad group
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
          adGroupDefaultBid: groupBid.toFixed(2),
        }));

        // Product ad
        result.push(createRow({
          entity: "Product ad",
          campaignId: campaignName,
          adGroupId: adGroupName,
          campaignNameInfo: campaignName,
          adGroupNameInfo: adGroupName,
          campaignStatus: "enabled",
          campaignStatusSecondary: "enabled",
          adGroupStatus: "enabled",
          sku: product.sku,
          asinInfo: product.asin,
          eligibilityStatus: "Eligible",
        }));
      }
    }
  } else {
    // Single campaign for all products
    const campaignName = `${config.brand} - ${config.category} - SP Auto - ALL - ${namingDate}`;

    result.push(createRow({
      entity: "Campaign",
      campaignId: campaignName,
      campaignName: campaignName,
      campaignNameInfo: campaignName,
      startDate,
      targetingType: "auto",
      campaignStatus: "enabled",
      campaignStatusSecondary: "enabled",
      budget: config.budget.toFixed(2),
      biddingStrategy: strategy,
    }));

    for (const placement of PLACEMENTS) {
      result.push(createRow({
        entity: "Bidding adjustment",
        campaignId: campaignName,
        campaignNameInfo: campaignName,
        campaignStatusSecondary: "enabled",
        biddingStrategy: strategy,
        placement,
        percentage: "0.00",
      }));
    }

    // One ad group per targeting group, all products inside
    for (const group of AUTO_TARGETING_GROUPS) {
      const adGroupName = `${campaignName} - ${group.name}`;
      const groupBid = config.groupBids?.[group.key as GroupKey] ?? config.defaultBid;

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
        adGroupDefaultBid: groupBid.toFixed(2),
      }));

      for (const product of validProducts) {
        result.push(createRow({
          entity: "Product ad",
          campaignId: campaignName,
          adGroupId: adGroupName,
          campaignNameInfo: campaignName,
          adGroupNameInfo: adGroupName,
          campaignStatus: "enabled",
          campaignStatusSecondary: "enabled",
          adGroupStatus: "enabled",
          sku: product.sku,
          asinInfo: product.asin,
          eligibilityStatus: "Eligible",
        }));
      }
    }
  }

  return result;
}

// ---------------------------------------------------------------------------
// Statistics
// ---------------------------------------------------------------------------

export interface AutoCampaignStats {
  totalRows: number;
  campaignCount: number;
  adGroupCount: number;
  productAdCount: number;
  biddingAdjustmentCount: number;
  productsPerCampaign: number;
}

export function getAutoCampaignStats(
  rows: BulkSheetRow[],
  productCount: number,
  structure: "per-product" | "single",
): AutoCampaignStats {
  return {
    totalRows: rows.length,
    campaignCount: rows.filter(r => r.entity === "Campaign").length,
    adGroupCount: rows.filter(r => r.entity === "Ad group").length,
    productAdCount: rows.filter(r => r.entity === "Product ad").length,
    biddingAdjustmentCount: rows.filter(r => r.entity === "Bidding adjustment").length,
    productsPerCampaign: structure === "per-product" ? 1 : productCount,
  };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const PLACEMENTS = [
  "Placement top",
  "Placement product page",
  "Placement rest of search",
  "Placement Amazon Business",
] as const;

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
