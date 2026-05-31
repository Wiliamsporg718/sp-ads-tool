import React, { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  type AdDetailRow,
  type BulkSheetRow,
  type BulkSheetOptions,
  type BulkSheetStats,
  generateBulkSheet,
  getBulkSheetStats,
  parseAdDetailRows,
  validateAdDetailRows,
  bulkRowToArray,
  BULK_SHEET_HEADERS,
} from "./lib/bulk-sheet-generator";
import {
  generateAsinPermutations,
  buildAdDetailAsinRows,
  parseAsinDataSheet,

  type AsinProduct,
  type AsinCombinationRow,
} from "./lib/asin-combinator";
import {
  generateAutoCampaigns,
  getAutoCampaignStats,
  type AutoCampaignProduct,
  type AutoCampaignConfig,
  type AutoCampaignStats,
} from "./lib/auto-campaign-generator";
import {
  analyzeSearchTerms,
  type SearchTermReportRow,
  type HarvestResult,
} from "./lib/search-term-harvester";

// ---------------------------------------------------------------------------
// SVG Icons (inline, no external dependency)
// ---------------------------------------------------------------------------

const Icons = {
  upload: (
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" />
    </svg>
  ),
  download: (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" />
    </svg>
  ),
  zap: (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2" />
    </svg>
  ),
  plus: (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round">
      <line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" />
    </svg>
  ),
  x: (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round">
      <line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" />
    </svg>
  ),
  check: (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="20 6 9 17 4 12" />
    </svg>
  ),
  refresh: (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="23 4 23 10 17 10" /><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10" />
    </svg>
  ),
  file: (
    <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" />
    </svg>
  ),
};

// ---------------------------------------------------------------------------
// Tab config
// ---------------------------------------------------------------------------

type Tab = "manual" | "auto" | "asin" | "harvest";

const TAB_CONFIG: { key: Tab; label: string; desc: string }[] = [
  { key: "manual", label: "手动投放", desc: "创建广告" },
  { key: "auto", label: "Auto Campaign", desc: "自动广告" },
  { key: "asin", label: "ASIN 组合", desc: "防御组合" },
  { key: "harvest", label: "搜索词分析", desc: "收割/否定" },
];

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------

export default function App() {
  const [tab, setTab] = useState<Tab>("manual");

  return (
    <div className="min-h-screen">
      {/* Header */}
      <header className="sticky top-0 z-50" style={{ background: "var(--brand-dark)" }}>
        <div className="max-w-7xl mx-auto px-4 sm:px-6 flex items-center h-14 gap-4">
          <div className="flex items-center gap-2">
            <div className="w-7 h-7 rounded-md flex items-center justify-center text-sm font-bold" style={{ background: "var(--brand-primary)", color: "#fff" }}>SP</div>
            <h1 className="text-white text-base font-semibold tracking-tight hidden sm:block">SP Ads Tool</h1>
          </div>
          <nav className="flex gap-1 ml-6">
            {TAB_CONFIG.map((t) => (
              <button
                key={t.key}
                onClick={() => setTab(t.key)}
                className="px-3 py-1.5 rounded-md text-sm font-medium transition-all duration-150"
                style={{
                  color: tab === t.key ? "#fff" : "rgba(255,255,255,0.6)",
                  background: tab === t.key ? "rgba(255,255,255,0.12)" : "transparent",
                }}
                onMouseEnter={(e) => { if (tab !== t.key) (e.target as HTMLElement).style.color = "rgba(255,255,255,0.85)"; }}
                onMouseLeave={(e) => { if (tab !== t.key) (e.target as HTMLElement).style.color = "rgba(255,255,255,0.6)"; }}
              >
                {t.label}
              </button>
            ))}
          </nav>
        </div>
      </header>

      {/* Content */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 py-6">
        <div className="animate-fadeIn">
          {tab === "manual" && <ManualCampaignTab />}
          {tab === "auto" && <AutoCampaignTab />}
          {tab === "asin" && <AsinPrepTab />}
          {tab === "auto" && <AutoCampaignTab />}
          {tab === "harvest" && <SearchTermHarvesterTab />}
        </div>
      </main>
    </div>
  );
}

// ===========================================================================
// Shared UI Components — Amazon Design System
// ===========================================================================

function Card({ children, className = "" }: { children: React.ReactNode; className?: string }) {
  return (
    <div
      className={`rounded-xl transition-shadow duration-200 ${className}`}
      style={{
        background: "var(--surface-card)",
        border: "1px solid var(--border-default)",
        boxShadow: "var(--shadow-sm)",
      }}
    >
      {children}
    </div>
  );
}

function CardHeader({ title, description, action }: { title: string; description?: string; action?: React.ReactNode }) {
  return (
    <div className="p-5 flex items-start justify-between">
      <div>
        <h2 className="text-lg font-semibold" style={{ color: "var(--text-primary)" }}>{title}</h2>
        {description && <p className="text-sm mt-0.5" style={{ color: "var(--text-tertiary)" }}>{description}</p>}
      </div>
      {action}
    </div>
  );
}

function StatCard({ label, value, color, icon }: { label: string; value: number | string; color?: string; icon?: React.ReactNode }) {
  return (
    <Card className="p-4 hover:shadow-[var(--shadow-md)]">
      <div className="flex items-center justify-between mb-2">
        <span className="text-xs font-medium uppercase tracking-wider" style={{ color: "var(--text-tertiary)" }}>{label}</span>
        {icon && <span style={{ color: color || "var(--text-tertiary)" }}>{icon}</span>}
      </div>
      <p className="text-2xl font-bold" style={{ color: color || "var(--text-primary)" }}>{value}</p>
    </Card>
  );
}

function Btn({ children, variant = "primary", size = "md", disabled, onClick, className = "" }: {
  children: React.ReactNode; variant?: "primary" | "amazon" | "outline" | "ghost" | "danger";
  size?: "sm" | "md" | "lg"; disabled?: boolean; onClick?: () => void; className?: string;
}) {
  const base = "inline-flex items-center justify-center gap-1.5 font-medium transition-all duration-150 active:scale-[0.98] disabled:opacity-50 disabled:pointer-events-none";
  const sizes = { sm: "px-3 py-1.5 text-xs rounded-md", md: "px-4 py-2 text-sm rounded-lg", lg: "px-6 py-2.5 text-sm rounded-lg" };
  const variants: Record<string, string> = {
    primary: "",
    amazon: "",
    outline: "",
    ghost: "",
    danger: "",
  };
  const variantStyles: Record<string, React.CSSProperties> = {
    primary: { background: "var(--brand-primary)", color: "#fff" },
    amazon: { background: "var(--brand-dark)", color: "#fff" },
    outline: { background: "transparent", color: "var(--text-secondary)", border: "1px solid var(--border-default)" },
    ghost: { background: "transparent", color: "var(--text-secondary)" },
    danger: { background: "var(--color-danger)", color: "#fff" },
  };

  return (
    <button
      className={`${base} ${sizes[size]} ${variants[variant]} ${className}`}
      style={variantStyles[variant]}
      disabled={disabled}
      onClick={onClick}
      onMouseEnter={(e) => {
        const el = e.currentTarget;
        if (variant === "primary") el.style.background = "var(--brand-primary-hover)";
        else if (variant === "amazon") el.style.background = "var(--brand-dark-hover)";
        else if (variant === "outline" || variant === "ghost") el.style.background = "var(--surface-hover)";
      }}
      onMouseLeave={(e) => {
        const el = e.currentTarget;
        el.style.background = variantStyles[variant].background as string;
      }}
    >
      {children}
    </button>
  );
}

function Input({ label, hint, ...props }: { label?: string; hint?: string } & React.InputHTMLAttributes<HTMLInputElement>) {
  return (
    <div>
      {label && <label className="block text-sm font-medium mb-1.5" style={{ color: "var(--text-primary)" }}>{label}</label>}
      <input
        {...props}
        className="w-full h-9 px-3 text-sm rounded-md transition-all duration-150 focus:ring-2 focus:ring-[var(--border-focus)] focus:border-[var(--border-focus)]"
        style={{ border: "1px solid var(--border-default)", background: "var(--surface-card)", boxShadow: "var(--shadow-xs)", color: "var(--text-primary)" }}
      />
      {hint && <p className="mt-1 text-xs" style={{ color: "var(--text-tertiary)" }}>{hint}</p>}
    </div>
  );
}

function Select({ label, children, ...props }: { label?: string } & React.SelectHTMLAttributes<HTMLSelectElement>) {
  return (
    <div>
      {label && <label className="block text-sm font-medium mb-1.5" style={{ color: "var(--text-primary)" }}>{label}</label>}
      <select
        {...props}
        className="w-full h-9 px-3 text-sm rounded-md transition-all duration-150 appearance-none bg-no-repeat bg-right pr-8"
        style={{
          border: "1px solid var(--border-default)", background: "var(--surface-card)", boxShadow: "var(--shadow-xs)",
          color: "var(--text-primary)",
          backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%23667085' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E")`,
          backgroundPosition: "right 8px center",
        }}
      >
        {children}
      </select>
    </div>
  );
}

function FileUpload({ onFile, accept = ".xlsx,.xls,.xlsm", label, hint }: {
  onFile: (file: File) => void; accept?: string; label: string; hint: string;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [isDragOver, setIsDragOver] = useState(false);
  return (
    <div
      role="button" tabIndex={0}
      className="flex flex-col items-center justify-center gap-3 rounded-xl p-12 cursor-pointer transition-all duration-200"
      style={{
        border: `2px dashed ${isDragOver ? "var(--brand-primary)" : "var(--border-strong)"}`,
        background: isDragOver ? "var(--brand-primary-light)" : "var(--surface-muted)",
      }}
      onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
      onDragLeave={(e) => { e.preventDefault(); setIsDragOver(false); }}
      onDrop={(e) => { e.preventDefault(); setIsDragOver(false); const f = e.dataTransfer.files[0]; if (f) onFile(f); }}
      onClick={() => inputRef.current?.click()}
      onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") inputRef.current?.click(); }}
    >
      <div className="rounded-full p-3" style={{ background: "var(--color-info-light)", color: "var(--color-info)" }}>
        {Icons.upload}
      </div>
      <div className="text-center">
        <p className="font-medium text-sm" style={{ color: "var(--text-primary)" }}>{label}</p>
        <p className="text-xs mt-1" style={{ color: "var(--text-tertiary)" }}>{hint}</p>
      </div>
      <Btn variant="outline" size="sm" onClick={() => { inputRef.current?.click(); }}>
        选择文件
      </Btn>
      <input ref={inputRef} type="file" accept={accept} className="hidden"
        onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); if (inputRef.current) inputRef.current.value = ""; }} />
    </div>
  );
}

function EntityBadge({ entity }: { entity: string }) {
  const styles: Record<string, { bg: string; fg: string }> = {
    Campaign: { bg: "var(--color-info-light)", fg: "var(--color-info-text)" },
    "Bidding adjustment": { bg: "var(--color-warning-light)", fg: "var(--color-warning-text)" },
    "Ad group": { bg: "var(--color-success-light)", fg: "var(--color-success-text)" },
    "Product ad": { bg: "var(--color-purple-light)", fg: "var(--color-purple)" },
    keyword: { bg: "var(--color-info-light)", fg: "var(--color-info-text)" },
    "Product targeting": { bg: "var(--color-purple-light)", fg: "var(--color-purple)" },
    "negative keyword": { bg: "var(--color-danger-light)", fg: "var(--color-danger-text)" },
    "negative product targeting": { bg: "var(--color-danger-light)", fg: "var(--color-danger-text)" },
  };
  const s = styles[entity] || { bg: "var(--surface-muted)", fg: "var(--text-secondary)" };
  return (
    <span
      className="inline-block px-2 py-0.5 rounded text-[10px] font-medium"
      style={{ background: s.bg, color: s.fg }}
    >
      {entity}
    </span>
  );
}

function Badge({ children, variant = "default" }: { children: React.ReactNode; variant?: "default" | "success" | "warning" | "danger" }) {
  const colors = {
    default: { bg: "var(--surface-muted)", fg: "var(--text-secondary)" },
    success: { bg: "var(--color-success-light)", fg: "var(--color-success-text)" },
    warning: { bg: "var(--color-warning-light)", fg: "var(--color-warning-text)" },
    danger: { bg: "var(--color-danger-light)", fg: "var(--color-danger-text)" },
  };
  const c = colors[variant];
  return <span className="inline-block px-2 py-0.5 rounded-md text-xs font-medium" style={{ background: c.bg, color: c.fg }}>{children}</span>;
}

function PageHeader({ title, description }: { title: string; description: string }) {
  return (
    <div className="mb-6">
      <h2 className="text-xl font-semibold" style={{ color: "var(--text-primary)" }}>{title}</h2>
      <p className="text-sm mt-1" style={{ color: "var(--text-tertiary)" }}>{description}</p>
    </div>
  );
}

function DataTable({ headers, children, maxHeight = "320px" }: { headers: { label: string; align?: "left" | "right" }[]; children: React.ReactNode; maxHeight?: string }) {
  return (
    <div className="overflow-x-auto rounded-lg" style={{ maxHeight, border: "1px solid var(--border-default)" }}>
      <table className="w-full text-xs">
        <thead className="sticky top-0" style={{ background: "var(--surface-muted)" }}>
          <tr>
            {headers.map((h, i) => (
              <th key={i} className={`p-2.5 font-semibold ${h.align === "right" ? "text-right" : "text-left"}`} style={{ color: "var(--text-secondary)", borderBottom: "1px solid var(--border-default)" }}>{h.label}</th>
            ))}
          </tr>
        </thead>
        <tbody>{children}</tbody>
      </table>
    </div>
  );
}

function downloadXlsx(sheets: { name: string; data: unknown[][] }[], fileName: string) {
  const wb = XLSX.utils.book_new();
  for (const s of sheets) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s.data), s.name);
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = fileName;
  document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}

// ===========================================================================
// Manual Campaign (Keyword / ASIN Targeting)
// ===========================================================================

interface ManualKeywordEntry {
  keyword: string;
  matchType: "exact" | "phrase" | "broad";
  bid: string;
}

interface ManualAdGroup {
  name: string;
  sku: string;
  asin: string;
  keywords: ManualKeywordEntry[];
}

function ManualCampaignTab() {
  const [campaignName, setCampaignName] = useState("");
  const [budget, setBudget] = useState("10.00");
  const [defaultBid, setDefaultBid] = useState("0.50");
  const [biddingStrategy, setBiddingStrategy] = useState("Dynamic bids - down only");
  const [mode, setMode] = useState<"keyword" | "asin">("keyword");
  const [adGroups, setAdGroups] = useState<ManualAdGroup[]>([{
    name: "", sku: "", asin: "", keywords: [{ keyword: "", matchType: "exact", bid: "" }],
  }]);
  const [bulkRows, setBulkRows] = useState<BulkSheetRow[]>([]);
  const [stats, setStats] = useState<BulkSheetStats | null>(null);
  const [generated, setGenerated] = useState(false);
  const [showBatchPaste, setShowBatchPaste] = useState<number | null>(null);
  const [batchText, setBatchText] = useState("");
  // File import state
  const [inputSource, setInputSource] = useState<"manual" | "file">("manual");
  const [fileRows, setFileRows] = useState<AdDetailRow[]>([]);
  const [fileName, setFileName] = useState("");
  const [fileErrors, setFileErrors] = useState<{ row: number; field: string; message: string }[]>([]);

  const updateAdGroup = useCallback((idx: number, patch: Partial<ManualAdGroup>) => {
    setAdGroups((prev) => prev.map((g, i) => i === idx ? { ...g, ...patch } : g));
  }, []);

  const addKeyword = useCallback((groupIdx: number) => {
    setAdGroups((prev) => prev.map((g, i) => i === groupIdx ? { ...g, keywords: [...g.keywords, { keyword: "", matchType: "exact", bid: "" }] } : g));
  }, []);

  const updateKeyword = useCallback((groupIdx: number, kwIdx: number, patch: Partial<ManualKeywordEntry>) => {
    setAdGroups((prev) => prev.map((g, gi) => gi === groupIdx ? {
      ...g, keywords: g.keywords.map((k, ki) => ki === kwIdx ? { ...k, ...patch } : k),
    } : g));
  }, []);

  const removeKeyword = useCallback((groupIdx: number, kwIdx: number) => {
    setAdGroups((prev) => prev.map((g, gi) => gi === groupIdx ? {
      ...g, keywords: g.keywords.filter((_, ki) => ki !== kwIdx),
    } : g));
  }, []);

  // Batch paste: one keyword per line
  const handleBatchPaste = useCallback((groupIdx: number) => {
    const lines = batchText.split(/\n/).map((l) => l.trim()).filter((l) => l.length > 0);
    if (lines.length === 0) return;
    const newKeywords: ManualKeywordEntry[] = lines.map((l) => ({ keyword: l, matchType: "exact" as const, bid: "" }));
    setAdGroups((prev) => prev.map((g, i) => i === groupIdx ? {
      ...g, keywords: [...g.keywords.filter((k) => k.keyword.trim()), ...newKeywords],
    } : g));
    setBatchText("");
    setShowBatchPaste(null);
  }, [batchText]);

  // Quick template: create 3 ad groups (Exact/Phrase/Broad) with same keywords
  const handleQuickTemplate = useCallback(() => {
    const baseName = campaignName || "SP_Manual";
    const sku = adGroups[0]?.sku || "";
    const asin = adGroups[0]?.asin || "";
    const existingKws = adGroups[0]?.keywords.filter((k) => k.keyword.trim()) || [];
    const templateKws = existingKws.length > 0 ? existingKws : [{ keyword: "", matchType: "exact" as const, bid: "" }];

    setAdGroups([
      { name: `${baseName}_Exact`, sku, asin, keywords: templateKws.map((k) => ({ ...k, matchType: "exact" as const })) },
      { name: `${baseName}_Phrase`, sku, asin, keywords: templateKws.map((k) => ({ ...k, matchType: "phrase" as const })) },
      { name: `${baseName}_Broad`, sku, asin, keywords: templateKws.map((k) => ({ ...k, matchType: "broad" as const })) },
    ]);
  }, [campaignName, adGroups]);

  // File import handler
  const handleFileImport = useCallback((file: File) => {
    file.arrayBuffer().then((buffer) => {
      const wb = XLSX.read(buffer, { type: "array" });
      const sheetName = wb.SheetNames.find((n) =>
        ["AD Detail Keywords", "AD Detail ASIN", "AD Detail", "Sheet1", "数据"].includes(n),
      ) || wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][];
      const parsed = parseAdDetailRows(raw);
      setFileRows(parsed); setFileName(file.name);
      setFileErrors(validateAdDetailRows(parsed, { mode }));
    });
  }, [mode]);

  const handleGenerate = useCallback(() => {
    let adDetailRows: AdDetailRow[];

    if (inputSource === "file") {
      adDetailRows = fileRows;
    } else {
      adDetailRows = [];
      for (const group of adGroups) {
        if (!group.name.trim() || (!group.sku.trim() && !group.asin.trim())) continue;
        for (const kw of group.keywords) {
          if (!kw.keyword.trim()) continue;
          adDetailRows.push({
            campaignName: campaignName || "Manual Campaign",
            adGroupName: group.name,
            budget: parseFloat(budget) || 10,
            sku: group.sku,
            asin: group.asin,
            bid: kw.bid ? parseFloat(kw.bid) : null,
            keywordText: kw.keyword,
            matchType: kw.matchType,
            biddingStrategy,
            startDate: "",
          });
        }
      }
    }

    if (adDetailRows.length === 0) return;

    const opts: BulkSheetOptions = { defaultAdGroupBid: defaultBid, targetingType: "manual", mode };
    const result = generateBulkSheet(adDetailRows, opts);
    setBulkRows(result); setStats(getBulkSheetStats(result)); setGenerated(true);
  }, [inputSource, fileRows, adGroups, campaignName, budget, defaultBid, biddingStrategy, mode]);

  const handleDownload = useCallback(() => {
    const prefix = mode === "asin" ? "BULK ASIN" : "BULK MANUAL";
    downloadXlsx(
      [{ name: prefix, data: [[...BULK_SHEET_HEADERS], ...bulkRows.map(bulkRowToArray)] }],
      `${prefix} ${new Date().toISOString().slice(0, 10)}.xlsx`,
    );
  }, [bulkRows, mode]);

  const totalKeywords = adGroups.reduce((sum, g) => sum + g.keywords.filter((k) => k.keyword.trim()).length, 0);
  const validGroups = adGroups.filter((g) => g.name.trim() && (g.sku.trim() || g.asin.trim()));

  return (
    <div className="space-y-5">
      <PageHeader title="创建 Manual Campaign" description="输入广告信息 → 生成 Amazon Bulk Sheet → 上传到广告后台即可投放" />

      {!generated ? (
        <>
          {/* Input source toggle */}
          <div className="flex gap-0" style={{ borderBottom: "1px solid var(--border-default)" }}>
            {([
              { key: "manual" as const, label: "手动输入" },
              { key: "file" as const, label: "从 AD Detail 文件导入" },
            ]).map((t) => (
              <button
                key={t.key}
                onClick={() => setInputSource(t.key)}
                className="px-4 py-2.5 text-sm font-medium transition-colors relative"
                style={{ color: inputSource === t.key ? "var(--text-primary)" : "var(--text-tertiary)" }}
              >
                {t.label}
                {inputSource === t.key && <div className="absolute bottom-0 left-0 right-0 h-[2.5px] rounded-full" style={{ background: "var(--border-focus)" }} />}
              </button>
            ))}
          </div>

          {inputSource === "file" ? (
            /* File import mode */
            <>
              {fileRows.length === 0 ? (
                <Card>
                  <div className="p-6">
                    <FileUpload onFile={handleFileImport} label="上传 AD Detail 规划表" hint="包含广告活动名称、广告组、SKU、关键词等列的 Excel 文件" />
                  </div>
                  <div className="px-6 pb-4 text-center">
                    <button className="text-xs underline" style={{ color: "var(--color-info)" }}
                      onClick={() => {
                        downloadXlsx([{
                          name: "AD Detail Keywords",
                          data: [
                            ["广告活动名称", "广告组名称", "预算", "SKU", "ASIN", "出价", "关键词", "匹配类型", "竞价策略", "开始日期"],
                            ["SP_Brand_Exact", "Brand_Exact", 10, "ABC-123", "B0EXAMPLE1", 0.5, "example keyword", "exact", "Dynamic bids - down only", ""],
                          ],
                        }], "AD Detail 模板.xlsx");
                      }}>下载模板文件</button>
                  </div>
                </Card>
              ) : (
                <>
                  <Card>
                    <div className="p-4 flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <div className="rounded-lg p-2" style={{ background: "var(--color-success-light)", color: "var(--color-success)" }}>{Icons.file}</div>
                        <div>
                          <p className="font-medium text-sm" style={{ color: "var(--text-primary)" }}>{fileName}</p>
                          <p className="text-xs" style={{ color: "var(--text-tertiary)" }}>
                            {fileRows.length} 行数据
                            {fileErrors.length > 0 && <span style={{ color: "var(--color-warning)" }}> · {fileErrors.length} 个警告</span>}
                          </p>
                        </div>
                      </div>
                      <Btn variant="ghost" size="sm" onClick={() => { setFileRows([]); setFileName(""); }}>
                        {Icons.refresh} 重新选择
                      </Btn>
                    </div>
                  </Card>
                  <Card>
                    <CardHeader title="生成参数" />
                    <div className="px-5 pb-5 space-y-4">
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <Select label="投放模式" value={mode} onChange={(e) => setMode(e.target.value as "keyword" | "asin")}>
                          <option value="keyword">关键词投放 (Keywords)</option>
                          <option value="asin">商品投放 (ASIN Targeting)</option>
                        </Select>
                        <Input label="默认竞价 ($)" type="number" step="0.01" min="0.02" value={defaultBid} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setDefaultBid(e.target.value)} />
                      </div>
                      <div className="flex justify-end pt-2">
                        <Btn variant="primary" onClick={handleGenerate}>{Icons.zap} 生成 Bulk Sheet</Btn>
                      </div>
                    </div>
                  </Card>
                </>
              )}
            </>
          ) : (
            /* Manual input mode */
            <>
          {/* Campaign config */}
          <Card>
            <CardHeader title="广告活动设置" />
            <div className="px-5 pb-5 space-y-4">
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
                <Input label="广告活动名称" placeholder="SP_Brand_Manual" value={campaignName} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setCampaignName(e.target.value)} />
                <Input label="日预算 ($)" type="number" step="0.01" value={budget} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setBudget(e.target.value)} />
                <Input label="默认竞价 ($)" type="number" step="0.01" value={defaultBid} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setDefaultBid(e.target.value)} />
                <Select label="投放模式" value={mode} onChange={(e) => setMode(e.target.value as "keyword" | "asin")}>
                  <option value="keyword">关键词投放</option>
                  <option value="asin">ASIN 投放</option>
                </Select>
              </div>
              <Select label="竞价策略" value={biddingStrategy} onChange={(e) => setBiddingStrategy(e.target.value)}>
                <option value="Dynamic bids - down only">Dynamic bids - down only（仅降低）</option>
                <option value="Dynamic bids - up and down">Dynamic bids - up and down（双向调整）</option>
                <option value="Fixed bid">Fixed bid（固定竞价）</option>
              </Select>
            </div>
          </Card>

          {/* Quick actions */}
          <div className="flex gap-2 flex-wrap">
            <Btn variant="outline" size="sm" onClick={handleQuickTemplate}>
              {Icons.zap} 快速模板：3 匹配类型
            </Btn>
            <span className="text-xs self-center" style={{ color: "var(--text-tertiary)" }}>
              一键创建 Exact + Phrase + Broad 三个广告组
            </span>
          </div>

          {/* Ad Groups */}
          {adGroups.map((group, gi) => (
            <Card key={gi}>
              <div className="p-5 flex items-start justify-between" style={{ borderBottom: "1px solid var(--border-default)" }}>
                <div className="flex items-center gap-2">
                  <Badge variant="success">广告组 {gi + 1}</Badge>
                  {adGroups.length > 1 && (
                    <button className="p-1 rounded-md" style={{ color: "var(--color-danger)" }}
                      onClick={() => setAdGroups((prev) => prev.filter((_, i) => i !== gi))}>
                      {Icons.x}
                    </button>
                  )}
                </div>
              </div>
              <div className="p-5 space-y-4">
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <Input label="广告组名称" placeholder="e.g. Brand_Exact" value={group.name}
                    onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateAdGroup(gi, { name: e.target.value })} />
                  <Input label="SKU" placeholder="ABC-123" value={group.sku}
                    onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateAdGroup(gi, { sku: e.target.value })} />
                  <Input label="ASIN" placeholder="B0XXXXXXXXX" value={group.asin}
                    onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateAdGroup(gi, { asin: e.target.value })} />
                </div>

                <div>
                  <div className="flex items-center justify-between mb-2">
                    <label className="text-sm font-medium" style={{ color: "var(--text-primary)" }}>
                      {mode === "keyword" ? "关键词" : "投放 ASIN"}
                    </label>
                    <div className="flex gap-1">
                      <Btn variant="ghost" size="sm" onClick={() => { setShowBatchPaste(showBatchPaste === gi ? null : gi); setBatchText(""); }}>
                        批量粘贴
                      </Btn>
                      <Btn variant="outline" size="sm" onClick={() => addKeyword(gi)}>{Icons.plus} 添加</Btn>
                    </div>
                  </div>

                  {/* Batch paste area */}
                  {showBatchPaste === gi && (
                    <div className="mb-3 p-3 rounded-lg" style={{ background: "var(--surface-muted)", border: "1px solid var(--border-default)" }}>
                      <textarea
                        rows={5}
                        className="w-full px-3 py-2 text-sm rounded-md mb-2 resize-none"
                        style={{ border: "1px solid var(--border-default)", background: "var(--surface-card)" }}
                        placeholder={"每行一个关键词，例如：\nwireless headphones\nbluetooth earbuds\nnoise cancelling"}
                        value={batchText}
                        onChange={(e) => setBatchText(e.target.value)}
                      />
                      <div className="flex items-center justify-between">
                        <span className="text-xs" style={{ color: "var(--text-tertiary)" }}>
                          {batchText.split(/\n/).filter((l) => l.trim()).length} 个关键词
                        </span>
                        <div className="flex gap-2">
                          <Btn variant="ghost" size="sm" onClick={() => setShowBatchPaste(null)}>取消</Btn>
                          <Btn variant="primary" size="sm" onClick={() => handleBatchPaste(gi)}>添加全部</Btn>
                        </div>
                      </div>
                    </div>
                  )}
                  <div className="space-y-2">
                    {group.keywords.map((kw, ki) => (
                      <div key={ki} className="flex gap-2 items-center group">
                        <input
                          placeholder={mode === "keyword" ? "输入关键词..." : "B0XXXXXXXXX"}
                          className="flex-1 h-9 px-3 text-sm rounded-md"
                          style={{ border: "1px solid var(--border-default)", background: "var(--surface-card)" }}
                          value={kw.keyword} onChange={(e) => updateKeyword(gi, ki, { keyword: e.target.value })}
                        />
                        {mode === "keyword" && (
                          <select
                            className="h-9 px-2 text-xs rounded-md"
                            style={{ border: "1px solid var(--border-default)", background: "var(--surface-card)", color: "var(--text-primary)", minWidth: "90px" }}
                            value={kw.matchType} onChange={(e) => updateKeyword(gi, ki, { matchType: e.target.value as "exact" | "phrase" | "broad" })}
                          >
                            <option value="exact">Exact</option>
                            <option value="phrase">Phrase</option>
                            <option value="broad">Broad</option>
                          </select>
                        )}
                        <input
                          placeholder="竞价"
                          className="w-20 h-9 px-2 text-sm text-right font-mono rounded-md"
                          style={{ border: "1px solid var(--border-default)", background: "var(--surface-card)" }}
                          value={kw.bid} onChange={(e) => updateKeyword(gi, ki, { bid: e.target.value })}
                        />
                        {group.keywords.length > 1 && (
                          <button className="p-1.5 rounded-md opacity-0 group-hover:opacity-100 transition-opacity"
                            style={{ color: "var(--color-danger)" }} onClick={() => removeKeyword(gi, ki)}>
                            {Icons.x}
                          </button>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </Card>
          ))}

          {/* Add group + Generate */}
          <div className="flex items-center justify-between">
            <Btn variant="outline" onClick={() => setAdGroups((prev) => [...prev, { name: "", sku: "", asin: "", keywords: [{ keyword: "", matchType: "exact", bid: "" }] }])}>
              {Icons.plus} 添加广告组
            </Btn>
            <div className="flex items-center gap-3">
              <span className="text-xs" style={{ color: "var(--text-tertiary)" }}>
                {validGroups.length} 个广告组 · {totalKeywords} 个{mode === "keyword" ? "关键词" : "ASIN"}
              </span>
              <Btn variant="primary" disabled={validGroups.length === 0 || totalKeywords === 0} onClick={handleGenerate}>
                {Icons.zap} 生成 Bulk Sheet
              </Btn>
            </div>
          </div>
            </>
          )}
        </>
      ) : (
        <>
          {stats && (
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <StatCard label="广告活动" value={stats.campaignCount} />
              <StatCard label="广告组" value={stats.adGroupCount} />
              <StatCard label="商品广告" value={stats.productAdCount} />
              <StatCard label="总行数" value={stats.totalRows} color="var(--color-success)" />
              {stats.keywordCount > 0 && <StatCard label="关键词" value={stats.keywordCount} color="var(--color-info)" />}
              {stats.productTargetingCount > 0 && <StatCard label="商品投放" value={stats.productTargetingCount} color="var(--color-purple)" />}
            </div>
          )}
          <Card>
            <div className="p-4 flex items-center justify-between" style={{ borderBottom: "1px solid var(--border-default)" }}>
              <div className="flex items-center gap-2">
                <span style={{ color: "var(--color-success)" }}>{Icons.check}</span>
                <span className="font-medium text-sm" style={{ color: "var(--color-success)" }}>生成完成</span>
              </div>
              <div className="flex gap-2">
                <Btn variant="ghost" size="sm" onClick={() => { setGenerated(false); setBulkRows([]); setStats(null); }}>
                  {Icons.refresh} 返回编辑
                </Btn>
                <Btn variant="primary" size="sm" onClick={handleDownload}>{Icons.download} 下载 XLSX</Btn>
              </div>
            </div>
            <div className="p-4">
              <DataTable headers={[
                { label: "#" }, { label: "Entity" }, { label: "Campaign" }, { label: "Ad Group" }, { label: "SKU" }, { label: "Keyword / ASIN" },
              ]}>
                {bulkRows.slice(0, 50).map((row, idx) => (
                  <tr key={idx} className="hover:bg-[var(--surface-hover)]" style={{ borderBottom: "1px solid var(--border-default)" }}>
                    <td className="p-2.5 text-xs" style={{ color: "var(--text-tertiary)" }}>{idx + 1}</td>
                    <td className="p-2.5"><EntityBadge entity={row.entity} /></td>
                    <td className="p-2.5 text-xs max-w-[180px] truncate">{row.campaignNameInfo}</td>
                    <td className="p-2.5 text-xs max-w-[140px] truncate">{row.adGroupNameInfo}</td>
                    <td className="p-2.5 text-xs font-mono">{row.sku}</td>
                    <td className="p-2.5 text-xs max-w-[140px] truncate">{row.keywordText || row.productTargetingExpression}</td>
                  </tr>
                ))}
              </DataTable>
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

// ===========================================================================
// Tab 3: ASIN Prep
// ===========================================================================

function AsinPrepTab() {
  const [products, setProducts] = useState<AsinProduct[]>([]);
  const [combinations, setCombinations] = useState<AsinCombinationRow[]>([]);
  const [fileName, setFileName] = useState("");
  const [budget, setBudget] = useState("5");
  const [bid, setBid] = useState("0.20");
  const [generated, setGenerated] = useState(false);
  const [inputMode, setInputMode] = useState<"manual" | "file">("manual");
  const [manualProducts, setManualProducts] = useState<AsinProduct[]>([{ sku: "", asin: "" }]);

  const handleFile = useCallback((file: File) => {
    file.arrayBuffer().then((buffer) => {
      const wb = XLSX.read(buffer, { type: "array" });
      const sheetName = wb.SheetNames.find((n) => n.includes("ASIN Data")) || wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][];
      const parsed = parseAsinDataSheet(raw);
      setProducts(parsed);
      setFileName(file.name);
    });
  }, []);

  const handleStartManual = useCallback(() => {
    const valid = manualProducts.filter((p) => p.sku.trim() && p.asin.trim());
    if (valid.length < 2) return;
    setProducts(valid);
    setFileName(`手动输入 ${valid.length} 个商品`);
  }, [manualProducts]);

  const handleGenerate = useCallback(() => {
    setCombinations(generateAsinPermutations(products));
    setGenerated(true);
  }, [products]);

  const handleDownload = useCallback(() => {
    const adRows = buildAdDetailAsinRows(combinations, {
      budget: parseFloat(budget) || 5, bid: parseFloat(bid) || 0.2, biddingStrategy: "Fixed bid",
    });
    const header = ["广告活动名称", "广告组", "预算", "SKU", "投放ASIN", "CLOSE出价", "竞品ASIN", "竞价策略", "开始日期"];
    const data = adRows.map((r) => [r.campaignName, r.adGroupName, r.budget, r.sku, r.asin, r.bid, r.competitorAsin, r.biddingStrategy, r.startDate]);
    downloadXlsx([{ name: "AD Detail ASIN", data: [header, ...data] }], `ASIN Prep ${new Date().toISOString().slice(0, 10)}.xlsx`);
  }, [combinations, budget, bid]);

  const expectedCombos = products.length * (products.length - 1);
  const validManualCount = manualProducts.filter((p) => p.sku.trim() && p.asin.trim()).length;

  return (
    <div className="space-y-5">
      <PageHeader title="ASIN 防御组合" description="将你的商品列表自动交叉排列，生成 N × (N-1) 防御性 ASIN 投放组合" />

      {products.length === 0 ? (
        <>
          {/* Input mode toggle */}
          <div className="flex gap-0" style={{ borderBottom: "1px solid var(--border-default)" }}>
            {([
              { key: "manual" as const, label: "直接输入商品" },
              { key: "file" as const, label: "从文件导入" },
            ]).map((t) => (
              <button
                key={t.key}
                onClick={() => setInputMode(t.key)}
                className="px-4 py-2.5 text-sm font-medium transition-colors relative"
                style={{ color: inputMode === t.key ? "var(--text-primary)" : "var(--text-tertiary)" }}
              >
                {t.label}
                {inputMode === t.key && <div className="absolute bottom-0 left-0 right-0 h-[2.5px] rounded-full" style={{ background: "var(--border-focus)" }} />}
              </button>
            ))}
          </div>

          {inputMode === "manual" ? (
            <Card>
              <CardHeader
                title="商品列表"
                description="输入至少 2 个商品的 SKU 和 ASIN，系统将自动生成所有交叉组合"
                action={
                  <Btn variant="outline" size="sm" onClick={() => setManualProducts((p) => [...p, { sku: "", asin: "" }])}>
                    {Icons.plus} 添加
                  </Btn>
                }
              />
              <div className="px-5 pb-5 space-y-2">
                {manualProducts.map((p, idx) => (
                  <div key={idx} className="flex gap-2 items-center group">
                    <span className="text-xs w-6 text-right font-mono" style={{ color: "var(--text-tertiary)" }}>{idx + 1}.</span>
                    <input placeholder="SKU"
                      className="flex-1 h-9 px-3 text-sm font-mono rounded-md transition-all"
                      style={{ border: "1px solid var(--border-default)", background: "var(--surface-card)" }}
                      value={p.sku} onChange={(e) => setManualProducts((prev) => prev.map((pr, i) => i === idx ? { ...pr, sku: e.target.value } : pr))} />
                    <input placeholder="ASIN (B0XXXXXXXXX)"
                      className="flex-1 h-9 px-3 text-sm font-mono rounded-md transition-all"
                      style={{ border: "1px solid var(--border-default)", background: "var(--surface-card)" }}
                      value={p.asin} onChange={(e) => setManualProducts((prev) => prev.map((pr, i) => i === idx ? { ...pr, asin: e.target.value } : pr))} />
                    {manualProducts.length > 1 && (
                      <button className="p-1.5 rounded-md opacity-0 group-hover:opacity-100 transition-opacity"
                        style={{ color: "var(--color-danger)" }}
                        onClick={() => setManualProducts((prev) => prev.filter((_, i) => i !== idx))}
                      >{Icons.x}</button>
                    )}
                  </div>
                ))}
                <div className="flex items-center justify-between pt-3">
                  <span className="text-xs" style={{ color: "var(--text-tertiary)" }}>
                    {validManualCount} 个有效商品{validManualCount >= 2 ? ` → ${validManualCount * (validManualCount - 1)} 个组合` : "（至少需要 2 个）"}
                  </span>
                  <Btn variant="primary" disabled={validManualCount < 2} onClick={handleStartManual}>
                    {Icons.zap} 开始组合
                  </Btn>
                </div>
              </div>
            </Card>
          ) : (
            <Card>
              <div className="p-6">
                <FileUpload onFile={handleFile} label="上传 ASIN 列表文件" hint="Excel 文件需包含 SKU 和 ASIN 列（适合大批量商品）" />
              </div>
            </Card>
          )}
        </>
      ) : (
        <>
          <Card>
            <div className="p-4 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="rounded-lg p-2" style={{ background: "var(--color-info-light)", color: "var(--color-info)" }}>{Icons.file}</div>
                <div>
                  <p className="font-medium text-sm" style={{ color: "var(--text-primary)" }}>{fileName}</p>
                  <div className="flex items-center gap-2 mt-0.5">
                    <Badge variant="success">{products.length} 个商品</Badge>
                    {products.length > 1 && <span className="text-xs" style={{ color: "var(--text-tertiary)" }}>预计 {expectedCombos} 个组合</span>}
                  </div>
                </div>
              </div>
              <Btn variant="ghost" size="sm" onClick={() => { setProducts([]); setGenerated(false); setCombinations([]); }}>
                {Icons.refresh} 重新开始
              </Btn>
            </div>
          </Card>
          <Card>
            <CardHeader title="组合参数" />
            <div className="px-5 pb-5 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <Input label="预算 ($)" type="number" step="0.01" value={budget} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setBudget(e.target.value)} />
                <Input label="竞价 ($)" type="number" step="0.01" value={bid} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setBid(e.target.value)} />
              </div>
              <div className="flex justify-end gap-2 pt-2">
                {!generated ? (
                  <Btn variant="primary" onClick={handleGenerate}>{Icons.zap} 生成组合</Btn>
                ) : (
                  <div className="flex items-center gap-3">
                    <Badge variant="success">{combinations.length} 个组合已生成</Badge>
                    <Btn variant="primary" size="sm" onClick={handleDownload}>{Icons.download} 下载 XLSX</Btn>
                  </div>
                )}
              </div>
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

// ===========================================================================
// Tab 3: Auto Campaign
// ===========================================================================

function AutoCampaignTab() {
  const [products, setProducts] = useState<AutoCampaignProduct[]>([{ sku: "", asin: "" }]);
  const [brand, setBrand] = useState("");
  const [category, setCategory] = useState("");
  const [budget, setBudget] = useState("10.00");
  const [defaultBid, setDefaultBid] = useState("0.50");
  const [structure, setStructure] = useState<"per-product" | "single">("per-product");
  const [bulkRows, setBulkRows] = useState<BulkSheetRow[]>([]);
  const [stats, setStats] = useState<AutoCampaignStats | null>(null);
  const [generated, setGenerated] = useState(false);

  const validProducts = products.filter((p) => p.sku.trim() && p.asin.trim());

  const handleGenerate = useCallback(() => {
    const config: AutoCampaignConfig = { brand, category, budget: parseFloat(budget) || 10, defaultBid: parseFloat(defaultBid) || 0.5, structure };
    const rows = generateAutoCampaigns(validProducts, config);
    setBulkRows(rows); setStats(getAutoCampaignStats(rows, validProducts.length, structure)); setGenerated(true);
  }, [brand, category, budget, defaultBid, structure, validProducts]);

  const handleDownload = useCallback(() => {
    downloadXlsx(
      [{ name: "BULK AUTO", data: [[...BULK_SHEET_HEADERS], ...bulkRows.map(bulkRowToArray)] }],
      `BULK AUTO ${new Date().toISOString().slice(0, 10)}.xlsx`,
    );
  }, [bulkRows]);

  return (
    <div className="space-y-5">
      <PageHeader title="Auto Campaign 生成器" description="输入商品信息即可一键创建完整的 SP 自动广告活动，包含 4 种匹配类型的独立广告组" />

      <Card>
        <CardHeader
          title="商品列表"
          action={
            <Btn variant="outline" size="sm" onClick={() => setProducts((p) => [...p, { sku: "", asin: "" }])}>
              {Icons.plus} 添加商品
            </Btn>
          }
        />
        <div className="px-5 pb-5 space-y-2">
          {products.map((p, idx) => (
            <div key={idx} className="flex gap-2 items-center group">
              <span className="text-xs w-6 text-right font-mono" style={{ color: "var(--text-tertiary)" }}>{idx + 1}.</span>
              <input placeholder="SKU"
                className="flex-1 h-9 px-3 text-sm font-mono rounded-md transition-all"
                style={{ border: "1px solid var(--border-default)", background: "var(--surface-card)" }}
                value={p.sku} onChange={(e) => setProducts((prev) => prev.map((pr, i) => i === idx ? { ...pr, sku: e.target.value } : pr))} />
              <input placeholder="ASIN (B0XXXXXXXXX)"
                className="flex-1 h-9 px-3 text-sm font-mono rounded-md transition-all"
                style={{ border: "1px solid var(--border-default)", background: "var(--surface-card)" }}
                value={p.asin} onChange={(e) => setProducts((prev) => prev.map((pr, i) => i === idx ? { ...pr, asin: e.target.value } : pr))} />
              {products.length > 1 && (
                <button className="p-1.5 rounded-md opacity-0 group-hover:opacity-100 transition-opacity"
                  style={{ color: "var(--color-danger)" }}
                  onClick={() => setProducts((prev) => prev.filter((_, i) => i !== idx))}
                >{Icons.x}</button>
              )}
            </div>
          ))}
        </div>
      </Card>

      <Card>
        <CardHeader title="广告配置" />
        <div className="px-5 pb-5 space-y-4">
          <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
            <Input label="品牌" placeholder="e.g. AGU" value={brand} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setBrand(e.target.value)} />
            <Input label="品类" placeholder="e.g. Jersey" value={category} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setCategory(e.target.value)} />
            <Input label="日预算 ($)" type="number" step="0.01" value={budget} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setBudget(e.target.value)} />
            <Input label="默认竞价 ($)" type="number" step="0.01" value={defaultBid} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setDefaultBid(e.target.value)} />
            <Select label="结构" value={structure} onChange={(e) => setStructure(e.target.value as "per-product" | "single")}>
              <option value="per-product">每商品独立广告活动</option>
              <option value="single">合并为单个广告活动</option>
            </Select>
          </div>
          <div className="flex justify-end pt-2">
            <Btn variant="primary" disabled={validProducts.length === 0 || !brand || !category} onClick={handleGenerate}>
              {Icons.zap} 生成 Auto Campaign
            </Btn>
          </div>
        </div>
      </Card>

      {generated && stats && (
        <>
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
            <StatCard label="广告活动" value={stats.campaignCount} />
            <StatCard label="广告组" value={stats.adGroupCount} />
            <StatCard label="商品广告" value={stats.productAdCount} />
            <StatCard label="竞价调整" value={stats.biddingAdjustmentCount} color="var(--color-warning)" />
            <StatCard label="总行数" value={stats.totalRows} color="var(--color-success)" />
          </div>
          <Card>
            <div className="p-4 flex items-center justify-between" style={{ borderBottom: "1px solid var(--border-default)" }}>
              <div className="flex items-center gap-2">
                <span style={{ color: "var(--color-success)" }}>{Icons.check}</span>
                <span className="font-medium text-sm" style={{ color: "var(--color-success)" }}>生成完成</span>
              </div>
              <Btn variant="primary" size="sm" onClick={handleDownload}>{Icons.download} 下载 XLSX</Btn>
            </div>
            <div className="p-4">
              <DataTable headers={[
                { label: "#" }, { label: "Entity" }, { label: "Campaign" }, { label: "Ad Group" }, { label: "SKU" },
              ]}>
                {bulkRows.slice(0, 40).map((row, idx) => (
                  <tr key={idx} className="hover:bg-[var(--surface-hover)]" style={{ borderBottom: "1px solid var(--border-default)" }}>
                    <td className="p-2.5 text-xs" style={{ color: "var(--text-tertiary)" }}>{idx + 1}</td>
                    <td className="p-2.5"><EntityBadge entity={row.entity} /></td>
                    <td className="p-2.5 text-xs max-w-[200px] truncate">{row.campaignNameInfo}</td>
                    <td className="p-2.5 text-xs max-w-[200px] truncate">{row.adGroupNameInfo}</td>
                    <td className="p-2.5 text-xs font-mono">{row.sku}</td>
                  </tr>
                ))}
              </DataTable>
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

// ===========================================================================
// Tab 4: Search Term Harvester
// ===========================================================================

function SearchTermHarvesterTab() {
  const [fileName, setFileName] = useState("");
  const [rows, setRows] = useState<SearchTermReportRow[]>([]);
  const [result, setResult] = useState<HarvestResult | null>(null);
  const [minClicks, setMinClicks] = useState("10");
  const [maxAcos, setMaxAcos] = useState("30");
  const [negMinClicks, setNegMinClicks] = useState("20");
  const [subTab, setSubTab] = useState<"harvest" | "negative">("harvest");

  const handleFile = useCallback((file: File) => {
    file.arrayBuffer().then((buffer) => {
      const wb = XLSX.read(buffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws);
      const parsed: SearchTermReportRow[] = raw
        .map((r) => ({
          searchTerm: String(r["Customer Search Term"] || r["搜索词"] || r["Search Term"] || ""),
          campaignName: String(r["Campaign Name"] || r["广告活动名称"] || ""),
          adGroupName: String(r["Ad Group Name"] || r["广告组名称"] || ""),
          impressions: Number(r["Impressions"] || r["展示量"] || 0),
          clicks: Number(r["Clicks"] || r["点击量"] || 0),
          spend: Number(r["Spend"] || r["花费"] || 0),
          sales: Number(r["7 Day Total Sales"] || r["Sales"] || r["销售额"] || 0),
          orders: Number(r["7 Day Total Orders (#)"] || r["Orders"] || r["订单"] || 0),
        }))
        .filter((r) => r.searchTerm.trim() !== "");
      setFileName(file.name); setRows(parsed);
    });
  }, []);

  const handleAnalyze = useCallback(() => {
    setResult(analyzeSearchTerms(rows, {
      minClicks: parseInt(minClicks) || 10,
      maxAcos: (parseFloat(maxAcos) || 30) / 100,
      negativeMinClicks: parseInt(negMinClicks) || 20,
    }));
  }, [rows, minClicks, maxAcos, negMinClicks]);

  const handleDownload = useCallback(() => {
    if (!result) return;
    downloadXlsx([
      { name: "Harvest", data: [["Search Term", "Type", "Clicks", "Orders", "ACOS", "Rec Bid"], ...result.harvestCandidates.map((c) => [c.term, c.type, c.clicks, c.orders, (c.acos * 100).toFixed(1) + "%", "$" + c.recommendedBid.toFixed(2)])] },
      { name: "Negatives", data: [["Search Term", "Reason", "Clicks", "Spend", "ACOS"], ...result.negativeCandidates.map((c) => [c.term, c.reason, c.clicks, c.spend, c.acos === Infinity ? "∞" : (c.acos * 100).toFixed(1) + "%"])] },
    ], `Search Term Analysis ${new Date().toISOString().slice(0, 10)}.xlsx`);
  }, [result]);

  return (
    <div className="space-y-5">
      <PageHeader title="搜索词收割分析" description="上传 Amazon 搜索词报告 → 自动识别高效词（收割到 Manual）和低效词（添加否定）" />

      {rows.length === 0 && (
        <>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-2">
            <Card className="p-4">
              <div className="flex items-center gap-2 mb-2">
                <span className="w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold" style={{ background: "var(--color-success-light)", color: "var(--color-success)" }}>✓</span>
                <span className="text-sm font-semibold" style={{ color: "var(--color-success)" }}>收割高效词</span>
              </div>
              <p className="text-xs" style={{ color: "var(--text-tertiary)" }}>找出点击高、转化好、ACOS 低的搜索词，建议添加到 Manual Exact 广告组精准投放</p>
            </Card>
            <Card className="p-4">
              <div className="flex items-center gap-2 mb-2">
                <span className="w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold" style={{ background: "var(--color-danger-light)", color: "var(--color-danger)" }}>✕</span>
                <span className="text-sm font-semibold" style={{ color: "var(--color-danger)" }}>否定低效词</span>
              </div>
              <p className="text-xs" style={{ color: "var(--text-tertiary)" }}>找出花费高但零转化或 ACOS 过高的搜索词，建议添加为否定关键词节省预算</p>
            </Card>
          </div>
          <Card>
            <div className="p-6">
              <FileUpload onFile={handleFile} accept=".xlsx,.xls,.csv,.xlsm" label="上传搜索词报告" hint="Amazon Seller Central → 广告 → 搜索词报告 → 导出" />
            </div>
            <div className="px-6 pb-5">
              <details className="text-xs" style={{ color: "var(--text-tertiary)" }}>
                <summary className="cursor-pointer hover:underline font-medium mb-2">如何下载搜索词报告？</summary>
                <ol className="list-decimal list-inside space-y-1 ml-1">
                  <li>登录 Amazon Seller Central</li>
                  <li>进入 广告 → 广告活动管理器</li>
                  <li>点击"报告"标签</li>
                  <li>创建报告：类型选"搜索词"，时间范围建议 30-60 天</li>
                  <li>下载 .xlsx 文件并上传到此处</li>
                </ol>
              </details>
            </div>
          </Card>
        </>
      )}

      {rows.length > 0 && !result && (
        <>
          <Card>
            <div className="p-4 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="rounded-lg p-2" style={{ background: "var(--color-info-light)", color: "var(--color-info)" }}>{Icons.file}</div>
                <div>
                  <p className="font-medium text-sm" style={{ color: "var(--text-primary)" }}>{fileName}</p>
                  <Badge variant="success">{rows.length} 条记录</Badge>
                </div>
              </div>
            </div>
          </Card>
          <Card>
            <CardHeader title="分析阈值" description="调整参数控制收割和否定的灵敏度" />
            <div className="px-5 pb-5 space-y-4">
              <div className="grid grid-cols-3 gap-4">
                <Input label="最低点击数" type="number" hint="收割候选最少需要的点击量" value={minClicks} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setMinClicks(e.target.value)} />
                <Input label="最大 ACOS (%)" type="number" hint="低于此 ACOS 视为高效" value={maxAcos} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setMaxAcos(e.target.value)} />
                <Input label="否定最低点击" type="number" hint="否定候选最少需要的点击量" value={negMinClicks} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setNegMinClicks(e.target.value)} />
              </div>
              <div className="flex justify-end pt-2">
                <Btn variant="primary" onClick={handleAnalyze}>{Icons.zap} 开始分析</Btn>
              </div>
            </div>
          </Card>
        </>
      )}

      {result && (
        <>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            <StatCard label="分析词数" value={result.stats.totalTermsAnalyzed} />
            <StatCard label="收割候选" value={result.stats.harvestCount} color="var(--color-success)" />
            <StatCard label="否定候选" value={result.stats.negativeCount} color="var(--color-danger)" />
            <StatCard label="预计节省" value={`$${result.stats.estimatedSavings.toFixed(2)}`} color="var(--brand-primary)" />
          </div>

          {/* Sub-tabs */}
          <div className="flex gap-0" style={{ borderBottom: "1px solid var(--border-default)" }}>
            {([
              { key: "harvest" as const, label: "收割候选", count: result.harvestCandidates.length, color: "var(--color-success)" },
              { key: "negative" as const, label: "否定候选", count: result.negativeCandidates.length, color: "var(--color-danger)" },
            ]).map((t) => (
              <button
                key={t.key}
                onClick={() => setSubTab(t.key)}
                className="px-4 py-2.5 text-sm font-medium transition-colors relative"
                style={{ color: subTab === t.key ? "var(--text-primary)" : "var(--text-tertiary)" }}
              >
                {t.label} <span className="ml-1 text-xs" style={{ color: t.color }}>({t.count})</span>
                {subTab === t.key && <div className="absolute bottom-0 left-0 right-0 h-[2.5px] rounded-full" style={{ background: "var(--border-focus)" }} />}
              </button>
            ))}
          </div>

          {subTab === "harvest" && (
            <Card>
              <div className="p-4 flex justify-between items-center" style={{ borderBottom: "1px solid var(--border-default)" }}>
                <span className="text-sm font-medium" style={{ color: "var(--text-secondary)" }}>高效搜索词 — 建议添加到 Manual Exact</span>
                <Btn variant="primary" size="sm" onClick={handleDownload}>{Icons.download} 导出分析</Btn>
              </div>
              <div className="p-4">
                <DataTable headers={[
                  { label: "#" }, { label: "搜索词" }, { label: "点击", align: "right" }, { label: "订单", align: "right" }, { label: "ACOS", align: "right" }, { label: "建议竞价", align: "right" },
                ]}>
                  {result.harvestCandidates.slice(0, 100).map((c, idx) => (
                    <tr key={idx} className="hover:bg-[var(--color-success-light)]" style={{ borderBottom: "1px solid var(--border-default)", transition: "background 150ms" }}>
                      <td className="p-2.5 text-xs" style={{ color: "var(--text-tertiary)" }}>{idx + 1}</td>
                      <td className="p-2.5 text-xs font-medium max-w-[220px] truncate">{c.term}</td>
                      <td className="p-2.5 text-xs text-right">{c.clicks}</td>
                      <td className="p-2.5 text-xs text-right">{c.orders}</td>
                      <td className="p-2.5 text-xs text-right" style={{ color: "var(--color-success)" }}>{(c.acos * 100).toFixed(1)}%</td>
                      <td className="p-2.5 text-xs text-right font-mono">${c.recommendedBid.toFixed(2)}</td>
                    </tr>
                  ))}
                </DataTable>
                {result.harvestCandidates.length === 0 && <p className="text-center py-8 text-sm" style={{ color: "var(--text-tertiary)" }}>未发现收割候选词</p>}
              </div>
            </Card>
          )}

          {subTab === "negative" && (
            <Card>
              <div className="p-4">
                <DataTable headers={[
                  { label: "#" }, { label: "搜索词" }, { label: "原因" }, { label: "点击", align: "right" }, { label: "花费", align: "right" }, { label: "ACOS", align: "right" },
                ]}>
                  {result.negativeCandidates.slice(0, 100).map((c, idx) => (
                    <tr key={idx} className="hover:bg-[var(--color-danger-light)]" style={{ borderBottom: "1px solid var(--border-default)", transition: "background 150ms" }}>
                      <td className="p-2.5 text-xs" style={{ color: "var(--text-tertiary)" }}>{idx + 1}</td>
                      <td className="p-2.5 text-xs font-medium max-w-[220px] truncate">{c.term}</td>
                      <td className="p-2.5"><Badge variant={c.reason === "no_conversion" ? "danger" : "warning"}>{c.reason === "no_conversion" ? "零转化" : "高ACOS"}</Badge></td>
                      <td className="p-2.5 text-xs text-right">{c.clicks}</td>
                      <td className="p-2.5 text-xs text-right" style={{ color: "var(--color-danger)" }}>${c.spend.toFixed(2)}</td>
                      <td className="p-2.5 text-xs text-right">{c.acos === Infinity ? "∞" : (c.acos * 100).toFixed(1) + "%"}</td>
                    </tr>
                  ))}
                </DataTable>
                {result.negativeCandidates.length === 0 && <p className="text-center py-8 text-sm" style={{ color: "var(--text-tertiary)" }}>未发现否定候选词</p>}
              </div>
            </Card>
          )}

          <div className="flex justify-center">
            <Btn variant="outline" onClick={() => { setRows([]); setFileName(""); setResult(null); }}>
              {Icons.refresh} 重新分析
            </Btn>
          </div>
        </>
      )}
    </div>
  );
}
