/**
 * Search Term Harvester — Auto-to-Manual Campaign Optimization (Standalone)
 *
 * Core PPC optimization loop:
 *   1. Auto campaigns discover search terms
 *   2. Search Term Report reveals performance data
 *   3. High-performing terms → harvest to Manual Exact campaigns
 *   4. Harvested terms → add as negatives in Auto campaigns
 *   5. Low-performing terms → add as negatives everywhere
 */

// ---------------------------------------------------------------------------
// Minimal report row type (standalone — no dependency on report-parser)
// ---------------------------------------------------------------------------

export interface SearchTermReportRow {
  searchTerm: string;
  campaignName: string;
  adGroupName: string;
  impressions: number;
  clicks: number;
  spend: number;
  sales: number;
  orders: number;
}

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface HarvestThresholds {
  minClicks?: number;
  maxAcos?: number;
  minOrders?: number;
  negativeAcosThreshold?: number;
  negativeMinClicks?: number;
}

export interface HarvestCandidate {
  term: string;
  type: "keyword" | "asin";
  sourceCampaign: string;
  sourceAdGroup: string;
  clicks: number;
  impressions: number;
  spend: number;
  sales: number;
  orders: number;
  acos: number;
  cvr: number;
  recommendedMatchType: "exact" | "phrase";
  recommendedBid: number;
}

export interface NegativeCandidate {
  term: string;
  type: "keyword" | "asin";
  sourceCampaign: string;
  sourceAdGroup: string;
  reason: "high_acos" | "no_conversion" | "irrelevant";
  clicks: number;
  spend: number;
  sales: number;
  orders: number;
  acos: number;
  recommendedMatchType: "negative exact" | "negative phrase";
}

export interface HarvestResult {
  harvestCandidates: HarvestCandidate[];
  negativeCandidates: NegativeCandidate[];
  stats: HarvestStats;
}

export interface HarvestStats {
  totalTermsAnalyzed: number;
  harvestCount: number;
  negativeCount: number;
  estimatedSavings: number;
  topHarvestTerm: string | null;
  worstNegativeTerm: string | null;
}

// ---------------------------------------------------------------------------
// Core Harvester
// ---------------------------------------------------------------------------

const DEFAULT_THRESHOLDS: Required<HarvestThresholds> = {
  minClicks: 10,
  maxAcos: 0.30,
  minOrders: 1,
  negativeAcosThreshold: 0.50,
  negativeMinClicks: 20,
};

export function analyzeSearchTerms(
  rows: SearchTermReportRow[],
  thresholds: HarvestThresholds = {},
): HarvestResult {
  const t = { ...DEFAULT_THRESHOLDS, ...thresholds };

  const termMap = new Map<string, AggregatedTerm>();

  for (const row of rows) {
    if (!row.searchTerm || row.searchTerm.trim() === "") continue;

    const key = `${row.searchTerm}||${row.campaignName}||${row.adGroupName}`;
    const existing = termMap.get(key);

    if (existing) {
      existing.impressions += row.impressions || 0;
      existing.clicks += row.clicks || 0;
      existing.spend += row.spend || 0;
      existing.sales += row.sales || 0;
      existing.orders += row.orders || 0;
    } else {
      termMap.set(key, {
        term: row.searchTerm.trim(),
        type: isAsin(row.searchTerm) ? "asin" : "keyword",
        campaign: row.campaignName || "",
        adGroup: row.adGroupName || "",
        impressions: row.impressions || 0,
        clicks: row.clicks || 0,
        spend: row.spend || 0,
        sales: row.sales || 0,
        orders: row.orders || 0,
      });
    }
  }

  const harvestCandidates: HarvestCandidate[] = [];
  const negativeCandidates: NegativeCandidate[] = [];

  for (const [, agg] of termMap) {
    const acos = agg.sales > 0 ? agg.spend / agg.sales : Infinity;
    const cvr = agg.clicks > 0 ? agg.orders / agg.clicks : 0;
    const cpc = agg.clicks > 0 ? agg.spend / agg.clicks : 0;

    if (agg.clicks >= t.minClicks && agg.orders >= t.minOrders && acos <= t.maxAcos) {
      harvestCandidates.push({
        term: agg.term,
        type: agg.type,
        sourceCampaign: agg.campaign,
        sourceAdGroup: agg.adGroup,
        clicks: agg.clicks,
        impressions: agg.impressions,
        spend: agg.spend,
        sales: agg.sales,
        orders: agg.orders,
        acos,
        cvr,
        recommendedMatchType: "exact",
        recommendedBid: Math.round(cpc * 1.2 * 100) / 100,
      });
    }

    if (agg.clicks >= t.negativeMinClicks && agg.orders === 0) {
      negativeCandidates.push({
        term: agg.term,
        type: agg.type,
        sourceCampaign: agg.campaign,
        sourceAdGroup: agg.adGroup,
        reason: "no_conversion",
        clicks: agg.clicks,
        spend: agg.spend,
        sales: 0,
        orders: 0,
        acos: Infinity,
        recommendedMatchType: "negative exact",
      });
    } else if (agg.clicks >= t.minClicks && agg.orders > 0 && acos > t.negativeAcosThreshold) {
      negativeCandidates.push({
        term: agg.term,
        type: agg.type,
        sourceCampaign: agg.campaign,
        sourceAdGroup: agg.adGroup,
        reason: "high_acos",
        clicks: agg.clicks,
        spend: agg.spend,
        sales: agg.sales,
        orders: agg.orders,
        acos,
        recommendedMatchType: "negative exact",
      });
    }
  }

  harvestCandidates.sort((a, b) => a.acos - b.acos);
  negativeCandidates.sort((a, b) => b.spend - a.spend);

  const estimatedSavings = negativeCandidates.reduce((sum, n) => sum + n.spend, 0);

  return {
    harvestCandidates,
    negativeCandidates,
    stats: {
      totalTermsAnalyzed: termMap.size,
      harvestCount: harvestCandidates.length,
      negativeCount: negativeCandidates.length,
      estimatedSavings,
      topHarvestTerm: harvestCandidates[0]?.term ?? null,
      worstNegativeTerm: negativeCandidates[0]?.term ?? null,
    },
  };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

interface AggregatedTerm {
  term: string;
  type: "keyword" | "asin";
  campaign: string;
  adGroup: string;
  impressions: number;
  clicks: number;
  spend: number;
  sales: number;
  orders: number;
}

function isAsin(term: string): boolean {
  return /^B[A-Z0-9]{9}$/i.test(term.trim());
}
