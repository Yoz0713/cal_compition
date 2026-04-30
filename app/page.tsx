"use client";

import { ChangeEvent, useMemo, useState } from "react";
import {
  DistanceResult,
  getExcelDataRowCount,
  groupByPharmacy,
  isExcelFile,
  normalizeName,
  parseExcel,
  PharmacySummary,
  ReferralRow,
  VALID_REGIONS,
  WeightedPharmacySummary,
} from "../lib/referrals";
import { paginate } from "../lib/paginate";

type UploadStats = {
  totalRows: number;
  validRows: number;
  invalidRows: number;
  pharmacyCount: number;
  centerCount: number;
};

type DistanceApiResponse = {
  centerName: string;
  pharmacyName: string;
  centerResolvedAddress?: string;
  pharmacyResolvedAddress?: string;
  distanceKm?: number;
  distanceScore: number;
  isSameStore: boolean;
  status: "success" | "same_store" | "place_not_found" | "distance_failed";
  errorMessage?: string;
};

// Concurrency limiter: run async tasks with max N in parallel
async function runWithConcurrency<T>(
  tasks: (() => Promise<T>)[],
  maxConcurrency: number,
): Promise<T[]> {
  const results: T[] = new Array(tasks.length);
  let nextIndex = 0;

  async function worker() {
    while (nextIndex < tasks.length) {
      const index = nextIndex++;
      results[index] = await tasks[index]();
    }
  }

  const workers = Array.from(
    { length: Math.min(maxConcurrency, tasks.length) },
    () => worker(),
  );
  await Promise.all(workers);
  return results;
}

const ACTIVE_REGIONS = [...VALID_REGIONS];
const PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

const DETAIL_FILTER_DEFAULTS = { keyword: "" };
const RANK_FILTER_DEFAULTS = { region: "全部" as string, keyword: "" };

export default function Home() {
  const [rows, setRows] = useState<ReferralRow[]>([]);
  const [invalidRows, setInvalidRows] = useState<ReferralRow[]>([]);
  const [summaries, setSummaries] = useState<PharmacySummary[]>([]);
  const [stats, setStats] = useState<UploadStats | null>(null);
  const [error, setError] = useState("");
  const [fileName, setFileName] = useState("");
  const [isLoading, setIsLoading] = useState(false);

  const [isCalculating, setIsCalculating] = useState(false);
  const [weightedResults, setWeightedResults] = useState<WeightedPharmacySummary[] | null>(null);
  const [distanceError, setDistanceError] = useState("");

  const [detailFilters, setDetailFilters] = useState(DETAIL_FILTER_DEFAULTS);
  const [detailPage, setDetailPage] = useState(1);
  const [detailPageSize, setDetailPageSize] = useState(25);

  const [rankFilters, setRankFilters] = useState(RANK_FILTER_DEFAULTS);
  const [rankPage, setRankPage] = useState(1);
  const [rankPageSize, setRankPageSize] = useState(25);

  const centerCount = useMemo(() => new Set(rows.map((r) => r.center)).size, [rows]);

  // ─── Region checks ───────────────────────────────────────────────

  const conflictPharmacies = useMemo(
    () => summaries.filter((s) => s.hasRegionConflict),
    [summaries],
  );

  const regionSummary = useMemo(() => {
    const result: Record<string, { pharmacyCount: number; referralCount: number }> = {};
    for (const r of ACTIVE_REGIONS) {
      result[r] = { pharmacyCount: 0, referralCount: 0 };
    }
    for (const s of summaries) {
      if (!s.hasRegionConflict && result[s.region]) {
        result[s.region].pharmacyCount += 1;
        result[s.region].referralCount += s.referralCount;
      }
    }
    return result;
  }, [summaries]);

  const canCalculate = useMemo(
    () =>
      summaries.length > 0 &&
      invalidRows.length === 0 &&
      conflictPharmacies.length === 0,
    [summaries, invalidRows, conflictPharmacies],
  );

  const disableReason = useMemo(() => {
    if (summaries.length === 0) return "";
    const reasons: string[] = [];
    if (invalidRows.length > 0) reasons.push(`${invalidRows.length} 筆異常資料`);
    if (conflictPharmacies.length > 0) reasons.push(`${conflictPharmacies.length} 間藥局區域衝突`);
    return reasons.join("、");
  }, [summaries, invalidRows, conflictPharmacies]);

  // ─── Flat distance detail list ───────────────────────────────────

  const allDistanceDetails = useMemo(() => {
    if (!weightedResults) return [];
    const details: DistanceResult[] = [];
    for (const w of weightedResults) {
      for (const d of w.distanceDetails) details.push(d);
    }
    return details;
  }, [weightedResults]);

  const filteredDetails = useMemo(() => {
    let items = allDistanceDetails;
    const { keyword } = detailFilters;
    if (keyword) {
      const kw = keyword.toLowerCase();
      items = items.filter((d) => d.center.toLowerCase().includes(kw) || d.pharmacy.toLowerCase().includes(kw));
    }
    return items;
  }, [allDistanceDetails, detailFilters]);

  const paginatedDetails = useMemo(
    () => paginate(filteredDetails, detailPage, detailPageSize),
    [filteredDetails, detailPage, detailPageSize],
  );

  // ─── Weighted rank list ──────────────────────────────────────────

  const filteredRanks = useMemo(() => {
    if (!weightedResults) return [];
    let items = [...weightedResults];
    const { region, keyword } = rankFilters;
    if (region !== "全部") items = items.filter((w) => w.region === region);
    if (keyword) {
      const kw = keyword.toLowerCase();
      items = items.filter((w) => w.name.toLowerCase().includes(kw));
    }
    return items.sort((a, b) => b.weightedTotalScore - a.weightedTotalScore);
  }, [weightedResults, rankFilters]);

  const paginatedRanks = useMemo(
    () => paginate(filteredRanks, rankPage, rankPageSize),
    [filteredRanks, rankPage, rankPageSize],
  );

  // ─── Handlers ────────────────────────────────────────────────────

  async function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    setError("");
    setIsLoading(true);
    setRows([]);
    setInvalidRows([]);
    setSummaries([]);
    setStats(null);
    setFileName(file.name);
    setWeightedResults(null);
    setDistanceError("");

    try {
      if (!isExcelFile(file)) throw new Error("檔案不是 Excel，請上傳 .xlsx、.xls 或 .xlsm 檔。");
      const parsedResult = await parseExcel(file);
      const parsedRows = parsedResult.rows;
      if (parsedRows.length === 0 && parsedResult.invalidRows.length === 0) {
        throw new Error("空資料：沒有可解析的轉介資料，請確認聽力中心與藥局名稱都有填寫。");
      }
      const pharmacySummaries = groupByPharmacy(parsedRows);
      const totalRows = await getExcelDataRowCount(file);
      setRows(parsedRows);
      setInvalidRows(parsedResult.invalidRows);
      setSummaries(pharmacySummaries);
      setStats({
        totalRows,
        validRows: parsedRows.length,
        invalidRows: parsedResult.invalidRows.length,
        pharmacyCount: pharmacySummaries.length,
        centerCount: new Set(parsedRows.map((r) => r.center)).size,
      });
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Excel 格式錯誤，請確認檔案內容。");
      setFileName("");
    } finally {
      setIsLoading(false);
    }
  }

  async function handleCalculateDistance() {
    if (!canCalculate) return;
    setIsCalculating(true);
    setDistanceError("");
    setWeightedResults(null);

    try {
      // 1. Collect unique center-pharmacy pairs
      const pairSet = new Map<string, { center: string; pharmacy: string }>();
      for (const row of rows) {
        const key = `${normalizeName(row.center)}__${normalizeName(row.pharmacy)}`;
        if (!pairSet.has(key)) pairSet.set(key, { center: row.center, pharmacy: row.pharmacy });
      }
      const uniquePairs = Array.from(pairSet.values());

      // 2. Call API for each pair (with cache + concurrency limit of 5)
      const distanceCache = new Map<string, DistanceApiResponse>();

      const tasks = uniquePairs.map((pair) => async () => {
        const response = await fetch("/api/distance", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ centerName: pair.center, pharmacyName: pair.pharmacy }),
        });
        const result: DistanceApiResponse = await response.json();
        const key = `${normalizeName(pair.center)}__${normalizeName(pair.pharmacy)}`;
        distanceCache.set(key, result);
      });

      await runWithConcurrency(tasks, 5);

      // 3. Check for failures
      const failedPairs = Array.from(distanceCache.values()).filter(
        (r) => r.status === "place_not_found" || r.status === "distance_failed",
      );
      if (failedPairs.length > 0) {
        setDistanceError(`部分距離計算失敗（${failedPairs.length} 組），請檢查門市名稱。失敗的距離分以 0 計算。`);
      }

      // 4. Build weighted summaries
      const weighted: WeightedPharmacySummary[] = summaries.map((summary) => {
        const details: DistanceResult[] = [];
        for (const row of summary.rows) {
          const cacheKey = `${normalizeName(row.center)}__${normalizeName(row.pharmacy)}`;
          const cached = distanceCache.get(cacheKey);
          const isSameStore = cached?.isSameStore ?? (normalizeName(row.center) === normalizeName(row.pharmacy));
          const distanceKm = cached?.distanceKm ?? 0;
          const distanceScore = cached?.distanceScore ?? 0;
          details.push({ center: row.center, pharmacy: row.pharmacy, distanceKm, distanceScore, isSameStore });
        }
        const distanceScoreTotal = details.reduce((sum, d) => sum + d.distanceScore, 0);
        return { ...summary, distanceScoreTotal, weightedTotalScore: summary.hearingScoreTotal + distanceScoreTotal, distanceDetails: details };
      });
      setWeightedResults(weighted);
      setDetailFilters(DETAIL_FILTER_DEFAULTS);
      setDetailPage(1);
      setRankFilters(RANK_FILTER_DEFAULTS);
      setRankPage(1);
    } catch (caughtError) {
      setDistanceError(caughtError instanceof Error ? caughtError.message : "距離計算失敗，請稍後再試。");
    } finally {
      setIsCalculating(false);
    }
  }

  function updateDetailFilter(patch: Partial<typeof DETAIL_FILTER_DEFAULTS>) {
    setDetailFilters((prev) => ({ ...prev, ...patch }));
    setDetailPage(1);
  }

  function updateRankFilter(patch: Partial<typeof RANK_FILTER_DEFAULTS>) {
    setRankFilters((prev) => ({ ...prev, ...patch }));
    setRankPage(1);
  }

  // ─── Render ──────────────────────────────────────────────────────

  return (
    <main className="page">
      <section className="header">
        <p className="eyebrow">第二階段 V2</p>
        <h1>轉介競賽計分小工具</h1>
        <p className="subtitle">上傳 Excel 後統計參與轉介藥局，依 Excel 區域欄位進行距離加權計分。</p>
      </section>

      <section className="panel guide-panel">
        <div className="step-title"><span>使用規範</span><h2>操作教學</h2></div>
        <div className="guide-grid">
          <div>
            <h3>Excel 欄位</h3>
            <ol>
              <li>第 1 列可放標題，系統從第 2 列開始讀取。</li>
              <li>A 欄為聽力中心，B 欄為轉介藥局。</li>
              <li>C 欄為左耳聽損，D 欄為右耳聽損。</li>
              <li>E 欄為參賽區域（桃區 / 竹苗區 / 宜花區 / 中彰投區）。</li>
            </ol>
          </div>
          <div>
            <h3>資料規則</h3>
            <ol>
              <li>聽力中心或藥局空白會直接略過。</li>
              <li>左右耳必須為數字，否則為異常資料。</li>
              <li>左右耳超出 -10 到 120 會列為異常資料。</li>
              <li>參賽區域必須為指定四區之一，否則為異常資料。</li>
            </ol>
          </div>
          <div>
            <h3>分數規則</h3>
            <ol>
              <li>任一耳大於等於 55，該筆為 10 分。</li>
              <li>否則任一耳大於等於 30，該筆為 3 分。</li>
              <li>其餘為 0 分，優先順序為 55 分級高於 30 分級。</li>
            </ol>
          </div>
          <div>
            <h3>區域規則</h3>
            <ol>
              <li>參賽區域由 Excel E 欄提供，系統不可手動修改。</li>
              <li>同一藥局只能有一個參賽區域。</li>
              <li>若同一藥局出現多區域，視為區域衝突，需回 Excel 修正。</li>
            </ol>
          </div>
        </div>
      </section>

      <section className="panel">
        <div className="step-title"><span>Step 1</span><h2>上傳 Excel</h2></div>
        <label className="upload-box">
          <input type="file" accept=".xlsx,.xls,.xlsm" onChange={handleFileChange} disabled={isLoading} />
          <span>{isLoading ? "解析中..." : "選擇 Excel 檔案"}</span>
          {fileName ? <strong>{fileName}</strong> : null}
        </label>
        {error ? <p className="error">{error}</p> : null}
      </section>

      {stats ? (
        <section className="panel">
          <div className="step-title"><span>Step 2</span><h2>摘要資訊</h2></div>
          <div className="stats-grid">
            <StatCard label="總資料筆數" value={stats.totalRows} />
            <StatCard label="有效資料筆數" value={stats.validRows} />
            <StatCard label="異常資料筆數" value={stats.invalidRows} />
            <StatCard label="藥局數量" value={stats.pharmacyCount} />
            <StatCard label="聽力中心數量" value={centerCount || stats.centerCount} />
          </div>
          {invalidRows.length > 0 ? (
            <p className="error">有 {invalidRows.length} 筆異常資料（左右耳數值無效、超出範圍、或參賽區域不在允許清單內），已排除於計算。請回 Excel 修正後重新上傳。</p>
          ) : null}
          {conflictPharmacies.length > 0 ? (
            <p className="error">
              {conflictPharmacies.length} 間藥局出現多個參賽區域（區域衝突），請回 Excel 修正後重新上傳。
              衝突藥局：{conflictPharmacies.map((p) => p.name).join("、")}
            </p>
          ) : null}
        </section>
      ) : null}

      {/* ─── Step 3: Region summary + pharmacy table ────────────── */}
      {summaries.length > 0 ? (
        <section className="panel">
          <div className="step-title"><span>Step 3</span><h2>藥局統計與區域檢查</h2></div>

          <div className="region-summary">
            {ACTIVE_REGIONS.map((r) => (
              <div className="region-card" key={r}>
                <strong>{r}</strong>
                <span>{regionSummary[r].pharmacyCount} 間藥局</span>
                <span>{regionSummary[r].referralCount} 筆轉介</span>
              </div>
            ))}
            <div className={`region-card ${invalidRows.length > 0 ? "region-card-warn" : ""}`}>
              <strong>異常資料</strong>
              <span>{invalidRows.length} 筆</span>
            </div>
            <div className={`region-card ${conflictPharmacies.length > 0 ? "region-card-warn" : ""}`}>
              <strong>區域衝突</strong>
              <span>{conflictPharmacies.length} 間</span>
            </div>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>藥局名稱</th><th>參賽區域</th><th>轉介筆數</th><th>0分</th><th>3分</th><th>10分</th>
                  <th>聽損總分</th><th>平均分</th>
                </tr>
              </thead>
              <tbody>
                {summaries.map((s) => (
                  <tr key={s.name} className={s.hasRegionConflict ? "row-conflict" : ""}>
                    <td>{s.name}</td>
                    <td className={s.hasRegionConflict ? "conflict-cell" : ""}>{s.region}</td>
                    <td>{s.referralCount}</td>
                    <td>{s.hearingBonus0Count}</td><td>{s.hearingBonus3Count}</td>
                    <td>{s.hearingBonus10Count}</td><td>{s.hearingScoreTotal}</td>
                    <td>{s.avgScore.toFixed(1)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="distance-action">
            {disableReason ? <p className="distance-hint">⚠ 無法計算：{disableReason}，請回 Excel 修正後重新上傳</p> : null}
            <button className="btn-distance" disabled={!canCalculate || isCalculating} onClick={handleCalculateDistance}>
              {isCalculating ? "計算中..." : "加權距離"}
            </button>
          </div>
          {distanceError ? <p className="notice">{distanceError}</p> : null}
        </section>
      ) : null}

      {/* ─── Step 4: Distance detail list ─────────────────────────── */}
      {weightedResults ? (
        <section className="panel">
          <div className="step-title"><span>Step 4</span><h2>Google API 查詢結果核對列表</h2></div>
          <div className="filter-bar">
            <label className="filter-label">店家搜尋
              <input className="filter-input" type="text" placeholder="輸入聽力中心或藥局名稱" value={detailFilters.keyword} onChange={(e) => updateDetailFilter({ keyword: e.target.value })} />
            </label>
            <button className="btn-clear" onClick={() => { setDetailFilters(DETAIL_FILTER_DEFAULTS); setDetailPage(1); }}>清除篩選</button>
          </div>
          {paginatedDetails.totalItems === 0 ? (
            <p className="empty-region">沒有符合條件的資料</p>
          ) : (
            <>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr><th>#</th><th>聽力中心</th><th>轉介藥局</th><th>距離(km)</th><th>距離分</th><th>同店</th></tr>
                  </thead>
                  <tbody>
                    {paginatedDetails.items.map((d, i) => (
                      <tr key={`${d.center}-${d.pharmacy}-${i}`}>
                        <td>{(paginatedDetails.currentPage - 1) * paginatedDetails.pageSize + i + 1}</td>
                        <td>{d.center}</td><td>{d.pharmacy}</td>
                        <td>{d.distanceKm.toFixed(1)}</td><td>{d.distanceScore}</td>
                        <td>{d.isSameStore ? "是" : "否"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <Pagination current={paginatedDetails.currentPage} total={paginatedDetails.totalPages} totalItems={paginatedDetails.totalItems} pageSize={detailPageSize} onPageChange={setDetailPage} onPageSizeChange={(s) => { setDetailPageSize(s); setDetailPage(1); }} />
            </>
          )}
        </section>
      ) : null}

      {/* ─── Step 5: Weighted ranking ─────────────────────────────── */}
      {weightedResults ? (
        <section className="panel">
          <div className="step-title"><span>Step 5</span><h2>加權後排名結果</h2></div>
          <div className="filter-bar">
            <label className="filter-label">參賽區域
              <select className="filter-select" value={rankFilters.region} onChange={(e) => updateRankFilter({ region: e.target.value })}>
                <option>全部</option>
                {ACTIVE_REGIONS.map((r) => <option key={r}>{r}</option>)}
              </select>
            </label>
            <label className="filter-label">藥局搜尋
              <input className="filter-input" type="text" placeholder="輸入藥局名稱" value={rankFilters.keyword} onChange={(e) => updateRankFilter({ keyword: e.target.value })} />
            </label>
            <button className="btn-clear" onClick={() => { setRankFilters(RANK_FILTER_DEFAULTS); setRankPage(1); }}>清除篩選</button>
          </div>
          {paginatedRanks.totalItems === 0 ? (
            <p className="empty-region">沒有符合條件的資料</p>
          ) : (
            <>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr><th>排名</th><th>藥局名稱</th><th>參賽區域</th><th>轉介筆數</th><th>聽損總分</th><th>距離加權分</th><th>加權後總分</th></tr>
                  </thead>
                  <tbody>
                    {paginatedRanks.items.map((w, i) => (
                      <tr key={w.name}>
                        <td className="rank-cell">{(paginatedRanks.currentPage - 1) * paginatedRanks.pageSize + i + 1}</td>
                        <td>{w.name}</td><td>{w.region}</td><td>{w.referralCount}</td>
                        <td>{w.hearingScoreTotal}</td><td>{w.distanceScoreTotal}</td>
                        <td className="total-cell">{w.weightedTotalScore}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <Pagination current={paginatedRanks.currentPage} total={paginatedRanks.totalPages} totalItems={paginatedRanks.totalItems} pageSize={rankPageSize} onPageChange={setRankPage} onPageSizeChange={(s) => { setRankPageSize(s); setRankPage(1); }} />
            </>
          )}
        </section>
      ) : null}
    </main>
  );
}

function StatCard({ label, value }: { label: string; value: number }) {
  return (
    <div className="stat-card">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function Pagination({
  current, total, totalItems, pageSize, onPageChange, onPageSizeChange,
}: {
  current: number; total: number; totalItems: number; pageSize: number;
  onPageChange: (p: number) => void; onPageSizeChange: (s: number) => void;
}) {
  return (
    <div className="pagination">
      <div className="page-size-wrap">
        <span>每頁</span>
        <select value={pageSize} onChange={(e) => onPageSizeChange(Number(e.target.value))}>
          {PAGE_SIZE_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <span>筆，共 {totalItems} 筆</span>
      </div>
      <div className="page-nav">
        <button disabled={current <= 1} onClick={() => onPageChange(1)}>«</button>
        <button disabled={current <= 1} onClick={() => onPageChange(current - 1)}>‹</button>
        <span className="page-info">{current} / {total}</span>
        <button disabled={current >= total} onClick={() => onPageChange(current + 1)}>›</button>
        <button disabled={current >= total} onClick={() => onPageChange(total)}>»</button>
      </div>
    </div>
  );
}
