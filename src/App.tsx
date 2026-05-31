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
// Tab config
// ---------------------------------------------------------------------------

type Tab = "bulk" | "asin" | "auto" | "harvest";

const TAB_CONFIG: { key: Tab; label: string; icon: string }[] = [
  { key: "bulk", label: "Bulk Sheet 生成", icon: "📋" },
  { key: "asin", label: "ASIN 数据准备", icon: "🔗" },
  { key: "auto", label: "Auto Campaign", icon: "🤖" },
  { key: "harvest", label: "搜索词收割", icon: "🔍" },
];

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------

export default function App() {
  const [tab, setTab] = useState<Tab>("bulk");

  return (
    <div className="p-4 md:p-6">
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-gray-900">SP 广告投放工具</h1>
        <p className="text-sm text-gray-500 mt-1">
          Amazon Sponsored Products 批量广告创建 · 搜索词优化 · 全流程自动化
        </p>
      </div>

      <div className="flex gap-1 mb-6 p-1 bg-gray-100 rounded-lg w-fit flex-wrap">
        {TAB_CONFIG.map((t) => (
          <button
            key={t.key}
            className={`px-3 py-2 rounded-md text-sm font-medium transition-colors ${
              tab === t.key
                ? "bg-white text-gray-900 shadow-sm"
                : "text-gray-500 hover:text-gray-700"
            }`}
            onClick={() => setTab(t.key)}
          >
            <span className="mr-1">{t.icon}</span>{t.label}
          </button>
        ))}
      </div>

      {tab === "bulk" && <BulkSheetTab />}
      {tab === "asin" && <AsinPrepTab />}
      {tab === "auto" && <AutoCampaignTab />}
      {tab === "harvest" && <SearchTermHarvesterTab />}
    </div>
  );
}

// ===========================================================================
// Shared UI
// ===========================================================================

function Card({ children, className = "" }: { children: React.ReactNode; className?: string }) {
  return <div className={`bg-white border border-gray-200 rounded-lg ${className}`}>{children}</div>;
}

function StatCard({ label, value, color }: { label: string; value: number | string; color?: string }) {
  return (
    <Card className="p-4 text-center">
      <p className={`text-2xl font-bold ${color || "text-gray-900"}`}>{value}</p>
      <p className="text-xs text-gray-500">{label}</p>
    </Card>
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
      className={`flex flex-col items-center justify-center gap-3 rounded-lg border-2 border-dashed p-10 cursor-pointer transition-all ${
        isDragOver ? "border-blue-500 bg-blue-50" : "border-gray-300 hover:border-blue-400"
      }`}
      onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
      onDragLeave={(e) => { e.preventDefault(); setIsDragOver(false); }}
      onDrop={(e) => { e.preventDefault(); setIsDragOver(false); const f = e.dataTransfer.files[0]; if (f) onFile(f); }}
      onClick={() => inputRef.current?.click()}
      onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") inputRef.current?.click(); }}
    >
      <span className="text-3xl">📂</span>
      <p className="font-medium text-gray-700">{label}</p>
      <p className="text-sm text-gray-500">{hint}</p>
      <button className="px-3 py-1.5 text-sm border border-gray-300 rounded-md hover:bg-gray-50"
        onClick={(e) => { e.stopPropagation(); inputRef.current?.click(); }}>
        选择文件
      </button>
      <input ref={inputRef} type="file" accept={accept} className="hidden"
        onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); if (inputRef.current) inputRef.current.value = ""; }} />
    </div>
  );
}

function EntityBadge({ entity }: { entity: string }) {
  const colors: Record<string, string> = {
    Campaign: "bg-blue-100 text-blue-800",
    "Bidding adjustment": "bg-amber-100 text-amber-800",
    "Ad group": "bg-green-100 text-green-800",
    "Product ad": "bg-purple-100 text-purple-800",
    keyword: "bg-cyan-100 text-cyan-800",
    "Product targeting": "bg-indigo-100 text-indigo-800",
    "negative keyword": "bg-red-100 text-red-800",
    "negative product targeting": "bg-red-100 text-red-800",
  };
  return (
    <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-medium ${colors[entity] || "bg-gray-100 text-gray-800"}`}>
      {entity}
    </span>
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
// Tab 1: Bulk Sheet
// ===========================================================================

function BulkSheetTab() {
  const [fileName, setFileName] = useState("");
  const [rows, setRows] = useState<AdDetailRow[]>([]);
  const [errors, setErrors] = useState<{ row: number; field: string; message: string }[]>([]);
  const [bulkRows, setBulkRows] = useState<BulkSheetRow[]>([]);
  const [stats, setStats] = useState<BulkSheetStats | null>(null);
  const [mode, setMode] = useState<"keyword" | "asin">("keyword");
  const [defaultBid, setDefaultBid] = useState("0.50");
  const [step, setStep] = useState<"upload" | "config" | "result">("upload");

  const handleFile = useCallback((file: File) => {
    file.arrayBuffer().then((buffer) => {
      const wb = XLSX.read(buffer, { type: "array" });
      const sheetName = wb.SheetNames.find((n) =>
        ["AD Detail Keywords", "AD Detail ASIN", "AD Detail", "Sheet1", "数据"].includes(n),
      ) || wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][];
      const parsed = parseAdDetailRows(raw);
      setFileName(file.name); setRows(parsed);
      setErrors(validateAdDetailRows(parsed, { mode }));
      setStep("config");
    });
  }, [mode]);

  const handleGenerate = useCallback(() => {
    const opts: BulkSheetOptions = { defaultAdGroupBid: defaultBid, targetingType: "manual", mode };
    const result = generateBulkSheet(rows, opts);
    setBulkRows(result); setStats(getBulkSheetStats(result)); setStep("result");
  }, [rows, defaultBid, mode]);

  const handleDownload = useCallback(() => {
    const prefix = mode === "asin" ? "BULK ASIN" : "BULK MANUAL";
    downloadXlsx(
      [{ name: prefix, data: [[...BULK_SHEET_HEADERS], ...bulkRows.map(bulkRowToArray)] }],
      `${prefix} ${new Date().toISOString().slice(0, 10)}.xlsx`,
    );
  }, [bulkRows, mode]);

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-semibold">Bulk Sheet 生成器</h2>
        <p className="text-sm text-gray-500">上传 AD Detail → 一键生成 Amazon SP Bulk Sheet</p>
      </div>

      {step === "upload" && (
        <Card className="p-6">
          <FileUpload onFile={handleFile} label="拖拽 AD Detail 文件到此处" hint="支持 .xlsx / .xls / .xlsm" />
        </Card>
      )}

      {step === "config" && (
        <>
          <Card className="p-4">
            <div className="flex items-center justify-between">
              <div>
                <p className="font-medium">{fileName}</p>
                <p className="text-sm text-gray-500">
                  {rows.length} 行数据
                  {errors.length > 0 && <span className="text-amber-600 ml-2">· {errors.length} 个警告</span>}
                </p>
              </div>
              <button className="text-sm text-gray-500 hover:text-gray-700"
                onClick={() => { setStep("upload"); setRows([]); }}>重新选择</button>
            </div>
          </Card>
          <Card className="p-6 space-y-4">
            <h3 className="font-semibold">参数配置</h3>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium mb-1">投放模式</label>
                <select className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm"
                  value={mode} onChange={(e) => setMode(e.target.value as "keyword" | "asin")}>
                  <option value="keyword">关键词投放</option>
                  <option value="asin">商品投放 (ASIN)</option>
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium mb-1">默认竞价 ($)</label>
                <input type="number" step="0.01" min="0.02"
                  className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm"
                  value={defaultBid} onChange={(e) => setDefaultBid(e.target.value)} />
              </div>
            </div>
            <div className="flex justify-end">
              <button className="px-4 py-2 bg-blue-600 text-white rounded-md text-sm font-medium hover:bg-blue-700"
                onClick={handleGenerate}>⚡ 生成 Bulk Sheet</button>
            </div>
          </Card>
        </>
      )}

      {step === "result" && stats && (
        <>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <StatCard label="广告活动" value={stats.campaignCount} />
            <StatCard label="广告组" value={stats.adGroupCount} />
            <StatCard label="商品广告" value={stats.productAdCount} />
            <StatCard label="总行数" value={stats.totalRows} color="text-green-600" />
            {stats.keywordCount > 0 && <StatCard label="关键词" value={stats.keywordCount} />}
            {stats.productTargetingCount > 0 && <StatCard label="商品投放" value={stats.productTargetingCount} />}
          </div>
          <Card className="p-4">
            <div className="flex items-center justify-between mb-3">
              <span className="font-medium text-green-700">✅ 生成完成</span>
              <div className="flex gap-2">
                <button className="px-3 py-1.5 text-sm border border-gray-300 rounded-md hover:bg-gray-50"
                  onClick={() => { setStep("upload"); setRows([]); setBulkRows([]); setStats(null); }}>重新开始</button>
                <button className="px-3 py-1.5 text-sm bg-blue-600 text-white rounded-md hover:bg-blue-700"
                  onClick={handleDownload}>📥 下载 XLSX</button>
              </div>
            </div>
            <div className="overflow-x-auto max-h-80">
              <table className="w-full text-xs">
                <thead className="sticky top-0 bg-white">
                  <tr className="border-b">
                    <th className="p-1.5 text-left">#</th>
                    <th className="p-1.5 text-left">Entity</th>
                    <th className="p-1.5 text-left">Campaign</th>
                    <th className="p-1.5 text-left">Ad Group</th>
                    <th className="p-1.5 text-left">SKU</th>
                    <th className="p-1.5 text-left">Keyword/ASIN</th>
                  </tr>
                </thead>
                <tbody>
                  {bulkRows.slice(0, 50).map((row, idx) => (
                    <tr key={idx} className="border-b border-gray-100 hover:bg-gray-50">
                      <td className="p-1.5 text-gray-400">{idx + 1}</td>
                      <td className="p-1.5"><EntityBadge entity={row.entity} /></td>
                      <td className="p-1.5 max-w-[150px] truncate">{row.campaignNameInfo}</td>
                      <td className="p-1.5 max-w-[120px] truncate">{row.adGroupNameInfo}</td>
                      <td className="p-1.5 font-mono">{row.sku}</td>
                      <td className="p-1.5 max-w-[120px] truncate">{row.keywordText || row.productTargetingExpression}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

// ===========================================================================
// Tab 2: ASIN Prep
// ===========================================================================

function AsinPrepTab() {
  const [products, setProducts] = useState<AsinProduct[]>([]);
  const [combinations, setCombinations] = useState<AsinCombinationRow[]>([]);
  const [fileName, setFileName] = useState("");
  const [budget, setBudget] = useState("5");
  const [bid, setBid] = useState("0.20");
  const [generated, setGenerated] = useState(false);

  const handleFile = useCallback((file: File) => {
    file.arrayBuffer().then((buffer) => {
      const wb = XLSX.read(buffer, { type: "array" });
      const sheetName = wb.SheetNames.find((n) => n.includes("ASIN Data")) || wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][];
      setProducts(parseAsinDataSheet(raw));
      setFileName(file.name);
    });
  }, []);

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

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-semibold">ASIN 数据准备</h2>
        <p className="text-sm text-gray-500">上传 ASIN 列表 → 生成 N×(N-1) 防御组合</p>
      </div>
      {products.length === 0 ? (
        <Card className="p-6">
          <FileUpload onFile={handleFile} label="上传 ASIN Data 文件" hint="需包含 SKU 和 ASIN 列" />
        </Card>
      ) : (
        <>
          <Card className="p-4">
            <div className="flex items-center justify-between">
              <div>
                <p className="font-medium">{fileName} — {products.length} 个商品</p>
                {products.length > 1 && <p className="text-sm text-gray-500">预计 {expectedCombos} 个组合</p>}
              </div>
              <button className="text-sm text-gray-500" onClick={() => { setProducts([]); setGenerated(false); }}>重新选择</button>
            </div>
          </Card>
          <Card className="p-6 space-y-4">
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium mb-1">预算 ($)</label>
                <input type="number" step="0.01" className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm" value={budget} onChange={(e) => setBudget(e.target.value)} />
              </div>
              <div>
                <label className="block text-sm font-medium mb-1">竞价 ($)</label>
                <input type="number" step="0.01" className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm" value={bid} onChange={(e) => setBid(e.target.value)} />
              </div>
            </div>
            <div className="flex justify-end gap-2">
              {!generated ? (
                <button className="px-4 py-2 bg-blue-600 text-white rounded-md text-sm font-medium hover:bg-blue-700" onClick={handleGenerate}>🔗 生成组合</button>
              ) : (
                <>
                  <span className="text-sm text-green-600">✅ {combinations.length} 个组合</span>
                  <button className="px-4 py-2 bg-blue-600 text-white rounded-md text-sm font-medium hover:bg-blue-700" onClick={handleDownload}>📥 下载</button>
                </>
              )}
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
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-semibold">Auto Campaign 生成器</h2>
        <p className="text-sm text-gray-500">一键创建 SP 自动广告（4 种匹配组）</p>
      </div>

      <Card className="p-6">
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-semibold">商品列表</h3>
          <button className="px-2 py-1 text-sm border border-gray-300 rounded-md hover:bg-gray-50"
            onClick={() => setProducts((p) => [...p, { sku: "", asin: "" }])}>+ 添加</button>
        </div>
        <div className="space-y-2">
          {products.map((p, idx) => (
            <div key={idx} className="flex gap-2 items-center">
              <span className="text-xs text-gray-400 w-5 text-right">{idx + 1}</span>
              <input placeholder="SKU" className="flex-1 border border-gray-300 rounded-md px-2 py-1.5 text-sm font-mono"
                value={p.sku} onChange={(e) => setProducts((prev) => prev.map((pr, i) => i === idx ? { ...pr, sku: e.target.value } : pr))} />
              <input placeholder="ASIN" className="flex-1 border border-gray-300 rounded-md px-2 py-1.5 text-sm font-mono"
                value={p.asin} onChange={(e) => setProducts((prev) => prev.map((pr, i) => i === idx ? { ...pr, asin: e.target.value } : pr))} />
              {products.length > 1 && (
                <button className="text-red-400 hover:text-red-600 text-sm"
                  onClick={() => setProducts((prev) => prev.filter((_, i) => i !== idx))}>✕</button>
              )}
            </div>
          ))}
        </div>
      </Card>

      <Card className="p-6 space-y-4">
        <h3 className="font-semibold">广告配置</h3>
        <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
          <div>
            <label className="block text-sm font-medium mb-1">品牌</label>
            <input className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm" value={brand} onChange={(e) => setBrand(e.target.value)} placeholder="e.g. AGU" />
          </div>
          <div>
            <label className="block text-sm font-medium mb-1">品类</label>
            <input className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm" value={category} onChange={(e) => setCategory(e.target.value)} placeholder="e.g. Jersey" />
          </div>
          <div>
            <label className="block text-sm font-medium mb-1">日预算 ($)</label>
            <input type="number" step="0.01" className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm" value={budget} onChange={(e) => setBudget(e.target.value)} />
          </div>
          <div>
            <label className="block text-sm font-medium mb-1">默认竞价 ($)</label>
            <input type="number" step="0.01" className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm" value={defaultBid} onChange={(e) => setDefaultBid(e.target.value)} />
          </div>
          <div>
            <label className="block text-sm font-medium mb-1">结构</label>
            <select className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm" value={structure} onChange={(e) => setStructure(e.target.value as "per-product" | "single")}>
              <option value="per-product">每商品一个广告活动</option>
              <option value="single">合并为一个广告活动</option>
            </select>
          </div>
        </div>
        <div className="flex justify-end">
          <button className="px-4 py-2 bg-blue-600 text-white rounded-md text-sm font-medium hover:bg-blue-700 disabled:opacity-50"
            disabled={validProducts.length === 0 || !brand || !category} onClick={handleGenerate}>🤖 生成 Auto Campaign</button>
        </div>
      </Card>

      {generated && stats && (
        <>
          <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
            <StatCard label="广告活动" value={stats.campaignCount} />
            <StatCard label="广告组" value={stats.adGroupCount} />
            <StatCard label="商品广告" value={stats.productAdCount} />
            <StatCard label="竞价调整" value={stats.biddingAdjustmentCount} />
            <StatCard label="总行数" value={stats.totalRows} color="text-green-600" />
          </div>
          <Card className="p-4">
            <div className="flex items-center justify-between mb-3">
              <span className="text-green-700 font-medium">✅ 生成完成</span>
              <button className="px-3 py-1.5 text-sm bg-blue-600 text-white rounded-md hover:bg-blue-700" onClick={handleDownload}>📥 下载 XLSX</button>
            </div>
            <div className="overflow-x-auto max-h-72">
              <table className="w-full text-xs">
                <thead className="sticky top-0 bg-white"><tr className="border-b">
                  <th className="p-1.5 text-left">#</th><th className="p-1.5 text-left">Entity</th>
                  <th className="p-1.5 text-left">Campaign</th><th className="p-1.5 text-left">Ad Group</th><th className="p-1.5 text-left">SKU</th>
                </tr></thead>
                <tbody>
                  {bulkRows.slice(0, 40).map((row, idx) => (
                    <tr key={idx} className="border-b border-gray-100 hover:bg-gray-50">
                      <td className="p-1.5 text-gray-400">{idx + 1}</td>
                      <td className="p-1.5"><EntityBadge entity={row.entity} /></td>
                      <td className="p-1.5 max-w-[180px] truncate">{row.campaignNameInfo}</td>
                      <td className="p-1.5 max-w-[180px] truncate">{row.adGroupNameInfo}</td>
                      <td className="p-1.5 font-mono">{row.sku}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
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
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-semibold">搜索词收割分析</h2>
        <p className="text-sm text-gray-500">分析高效词（→ Manual Exact）和低效词（→ 否定）</p>
      </div>

      {rows.length === 0 && (
        <Card className="p-6">
          <FileUpload onFile={handleFile} accept=".xlsx,.xls,.csv,.xlsm" label="上传 Search Term Report" hint="Amazon 后台导出的搜索词报告" />
        </Card>
      )}

      {rows.length > 0 && !result && (
        <>
          <Card className="p-4"><p className="font-medium">{fileName} — {rows.length} 条记录</p></Card>
          <Card className="p-6 space-y-4">
            <h3 className="font-semibold">分析阈值</h3>
            <div className="grid grid-cols-3 gap-4">
              <div><label className="block text-xs text-gray-500 mb-1">最低点击数</label>
                <input type="number" className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm" value={minClicks} onChange={(e) => setMinClicks(e.target.value)} /></div>
              <div><label className="block text-xs text-gray-500 mb-1">最大 ACOS (%)</label>
                <input type="number" className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm" value={maxAcos} onChange={(e) => setMaxAcos(e.target.value)} /></div>
              <div><label className="block text-xs text-gray-500 mb-1">否定最低点击</label>
                <input type="number" className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm" value={negMinClicks} onChange={(e) => setNegMinClicks(e.target.value)} /></div>
            </div>
            <div className="flex justify-end">
              <button className="px-4 py-2 bg-blue-600 text-white rounded-md text-sm font-medium hover:bg-blue-700" onClick={handleAnalyze}>🔍 开始分析</button>
            </div>
          </Card>
        </>
      )}

      {result && (
        <>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <StatCard label="分析词数" value={result.stats.totalTermsAnalyzed} />
            <StatCard label="收割候选" value={result.stats.harvestCount} color="text-green-600" />
            <StatCard label="否定候选" value={result.stats.negativeCount} color="text-red-600" />
            <StatCard label="预计节省" value={`$${result.stats.estimatedSavings.toFixed(2)}`} color="text-amber-600" />
          </div>

          <div className="flex gap-1 p-1 bg-gray-100 rounded-lg w-fit">
            <button className={`px-3 py-1.5 rounded-md text-sm font-medium ${subTab === "harvest" ? "bg-white text-green-700 shadow-sm" : "text-gray-500"}`} onClick={() => setSubTab("harvest")}>
              📈 收割 ({result.harvestCandidates.length})
            </button>
            <button className={`px-3 py-1.5 rounded-md text-sm font-medium ${subTab === "negative" ? "bg-white text-red-700 shadow-sm" : "text-gray-500"}`} onClick={() => setSubTab("negative")}>
              🚫 否定 ({result.negativeCandidates.length})
            </button>
          </div>

          {subTab === "harvest" && (
            <Card className="p-4">
              <div className="flex justify-end mb-2">
                <button className="px-3 py-1.5 text-sm bg-blue-600 text-white rounded-md hover:bg-blue-700" onClick={handleDownload}>📥 导出</button>
              </div>
              <div className="overflow-x-auto max-h-80">
                <table className="w-full text-xs">
                  <thead className="sticky top-0 bg-white"><tr className="border-b">
                    <th className="p-1.5 text-left">#</th><th className="p-1.5 text-left">搜索词</th>
                    <th className="p-1.5 text-right">点击</th><th className="p-1.5 text-right">订单</th>
                    <th className="p-1.5 text-right">ACOS</th><th className="p-1.5 text-right">建议竞价</th>
                  </tr></thead>
                  <tbody>
                    {result.harvestCandidates.slice(0, 100).map((c, idx) => (
                      <tr key={idx} className="border-b border-gray-100 hover:bg-green-50">
                        <td className="p-1.5 text-gray-400">{idx + 1}</td>
                        <td className="p-1.5 font-medium max-w-[200px] truncate">{c.term}</td>
                        <td className="p-1.5 text-right">{c.clicks}</td>
                        <td className="p-1.5 text-right">{c.orders}</td>
                        <td className="p-1.5 text-right text-green-700">{(c.acos * 100).toFixed(1)}%</td>
                        <td className="p-1.5 text-right font-mono">${c.recommendedBid.toFixed(2)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {result.harvestCandidates.length === 0 && <p className="text-center py-6 text-gray-400">未发现收割候选词</p>}
              </div>
            </Card>
          )}

          {subTab === "negative" && (
            <Card className="p-4">
              <div className="overflow-x-auto max-h-80">
                <table className="w-full text-xs">
                  <thead className="sticky top-0 bg-white"><tr className="border-b">
                    <th className="p-1.5 text-left">#</th><th className="p-1.5 text-left">搜索词</th>
                    <th className="p-1.5 text-left">原因</th><th className="p-1.5 text-right">点击</th>
                    <th className="p-1.5 text-right">花费</th><th className="p-1.5 text-right">ACOS</th>
                  </tr></thead>
                  <tbody>
                    {result.negativeCandidates.slice(0, 100).map((c, idx) => (
                      <tr key={idx} className="border-b border-gray-100 hover:bg-red-50">
                        <td className="p-1.5 text-gray-400">{idx + 1}</td>
                        <td className="p-1.5 font-medium max-w-[200px] truncate">{c.term}</td>
                        <td className="p-1.5">
                          <span className={`px-1.5 py-0.5 rounded text-[10px] font-medium ${c.reason === "no_conversion" ? "bg-red-100 text-red-800" : "bg-amber-100 text-amber-800"}`}>
                            {c.reason === "no_conversion" ? "零转化" : "高ACOS"}
                          </span>
                        </td>
                        <td className="p-1.5 text-right">{c.clicks}</td>
                        <td className="p-1.5 text-right text-red-600">${c.spend.toFixed(2)}</td>
                        <td className="p-1.5 text-right">{c.acos === Infinity ? "∞" : (c.acos * 100).toFixed(1) + "%"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {result.negativeCandidates.length === 0 && <p className="text-center py-6 text-gray-400">未发现否定候选词</p>}
              </div>
            </Card>
          )}

          <div className="flex justify-center">
            <button className="px-4 py-2 border border-gray-300 rounded-md text-sm hover:bg-gray-50"
              onClick={() => { setRows([]); setFileName(""); setResult(null); }}>重新分析</button>
          </div>
        </>
      )}
    </div>
  );
}
