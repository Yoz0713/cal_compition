"use client";

import { ChangeEvent, useMemo, useState } from "react";
import {
  DistanceResult,
  getDistanceScore,
  getExcelDataRowCount,
  groupByPharmacy,
  isExcelFile,
  normalizeName,
  parseExcel,
  PharmacySummary,
  ReferralRow,
  WeightedPharmacySummary,
} from "../lib/referrals";

type UploadStats = {
  totalRows: number;
  validRows: number;
  invalidRows: number;
  pharmacyCount: number;
  centerCount: number;
};

type DistanceApiResult = {
  center: string;
  pharmacy: string;
  distanceKm: number | null;
  error?: string;
};

const REGIONS = ["未分類", "桃區", "竹苗區", "宜花區", "中彰投區"];
const ACTIVE_REGIONS = REGIONS.filter((r) => r !== "未分類");

export default function Home() {
  const [rows, setRows] = useState<ReferralRow[]>([]);
  const [invalidRows, setInvalidRows] = useState<ReferralRow[]>([]);
  const [summaries, setSummaries] = useState<PharmacySummary[]>([]);
  const [stats, setStats] = useState<UploadStats | null>(null);
  const [error, setError] = useState("");
  const [fileName, setFileName] = useState("");
  const [isLoading, setIsLoading] = useState(false);

  // V2 state
  const [isCalculating, setIsCalculating] = useState(false);
  const [weightedResults, setWeightedResults] = useState<
    WeightedPharmacySummary[] | null
  >(null);
  const [distanceError, setDistanceError] = useState("");
  const [activeTab, setActiveTab] = useState(ACTIVE_REGIONS[0]);

  const centerCount = useMemo(
    () => new Set(rows.map((row) => row.center)).size,
    [rows],
  );

  const allClassified = useMemo(
    () =>
      summaries.length > 0 &&
      summaries.every((s) => s.region !== "未分類"),
    [summaries],
  );

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
    // Reset V2 state on new upload
    setWeightedResults(null);
    setDistanceError("");

    try {
      if (!isExcelFile(file)) {
        throw new Error("檔案不是 Excel，請上傳 .xlsx、.xls 或 .xlsm 檔。");
      }

      const parsedResult = await parseExcel(file);
      const parsedRows = parsedResult.rows;

      if (parsedRows.length === 0) {
        throw new Error("空資料：沒有有效轉介資料，請確認聽力中心與藥局名稱都有填寫。");
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
        centerCount: new Set(parsedRows.map((row) => row.center)).size,
      });
    } catch (caughtError) {
      setError(
        caughtError instanceof Error
          ? caughtError.message
          : "Excel 格式錯誤，請確認檔案內容。",
      );
      setFileName("");
    } finally {
      setIsLoading(false);
    }
  }

  function handleRegionChange(name: string, region: string) {
    setSummaries((current) =>
      current.map((summary) =>
        summary.name === name ? { ...summary, region } : summary,
      ),
    );
    // Clear previous weighted results when region changes
    setWeightedResults(null);
    setDistanceError("");
  }

  // ─── V2: Calculate weighted distance scores ────────────────────────

  async function handleCalculateDistance() {
    if (!allClassified) return;

    setIsCalculating(true);
    setDistanceError("");
    setWeightedResults(null);

    try {
      // 1. Collect unique center-pharmacy pairs
      const pairSet = new Map<string, { center: string; pharmacy: string }>();

      for (const row of rows) {
        const key = `${normalizeName(row.center)}||${normalizeName(row.pharmacy)}`;
        if (!pairSet.has(key)) {
          pairSet.set(key, { center: row.center, pharmacy: row.pharmacy });
        }
      }

      const uniquePairs = Array.from(pairSet.values());

      // 2. Identify same-store pairs (skip API call for these)
      const sameStorePairs: typeof uniquePairs = [];
      const apiPairs: typeof uniquePairs = [];

      for (const pair of uniquePairs) {
        if (normalizeName(pair.center) === normalizeName(pair.pharmacy)) {
          sameStorePairs.push(pair);
        } else {
          apiPairs.push(pair);
        }
      }

      // 3. Call API for non-same-store pairs
      const distanceCache = new Map<string, DistanceApiResult>();

      // Pre-fill same-store results
      for (const pair of sameStorePairs) {
        const key = `${normalizeName(pair.center)}||${normalizeName(pair.pharmacy)}`;
        distanceCache.set(key, {
          center: pair.center,
          pharmacy: pair.pharmacy,
          distanceKm: 0,
        });
      }

      if (apiPairs.length > 0) {
        const response = await fetch("/api/distance", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ pairs: apiPairs }),
        });

        if (!response.ok) {
          const errBody = await response.json().catch(() => ({}));
          throw new Error(
            (errBody as { error?: string }).error ??
              `API 錯誤 ${response.status}`,
          );
        }

        const data: { results: DistanceApiResult[] } = await response.json();

        for (const result of data.results) {
          const key = `${normalizeName(result.center)}||${normalizeName(result.pharmacy)}`;
          distanceCache.set(key, result);
        }
      }

      // 4. Check for partial failures
      const failedPairs = Array.from(distanceCache.values()).filter(
        (r) => r.distanceKm === null,
      );

      if (failedPairs.length > 0) {
        setDistanceError(
          `部分距離計算失敗（${failedPairs.length} 組），請檢查門市名稱。失敗的距離以 0 公里計算。`,
        );
      }

      // 5. Build weighted summaries
      const weighted: WeightedPharmacySummary[] = summaries.map((summary) => {
        const details: DistanceResult[] = [];

        for (const row of summary.rows) {
          const isSameStore =
            normalizeName(row.center) === normalizeName(row.pharmacy);
          const cacheKey = `${normalizeName(row.center)}||${normalizeName(row.pharmacy)}`;
          const cached = distanceCache.get(cacheKey);
          const distanceKm = cached?.distanceKm ?? 0;
          const distanceScore = getDistanceScore(distanceKm, isSameStore);

          details.push({
            center: row.center,
            pharmacy: row.pharmacy,
            distanceKm,
            distanceScore,
            isSameStore,
          });
        }

        const distanceScoreTotal = details.reduce(
          (sum, d) => sum + d.distanceScore,
          0,
        );

        return {
          ...summary,
          distanceScoreTotal,
          weightedTotalScore: summary.hearingScoreTotal + distanceScoreTotal,
          distanceDetails: details,
        };
      });

      setWeightedResults(weighted);
    } catch (caughtError) {
      setDistanceError(
        caughtError instanceof Error
          ? caughtError.message
          : "距離計算失敗，請稍後再試。",
      );
    } finally {
      setIsCalculating(false);
    }
  }

  // Group weighted results by region
  const resultsByRegion = useMemo(() => {
    if (!weightedResults) return null;

    const grouped: Record<string, WeightedPharmacySummary[]> = {};

    for (const region of ACTIVE_REGIONS) {
      grouped[region] = weightedResults
        .filter((w) => w.region === region)
        .sort((a, b) => b.weightedTotalScore - a.weightedTotalScore);
    }

    return grouped;
  }, [weightedResults]);

  return (
    <main className="page">
      <section className="header">
        <p className="eyebrow">第二階段 V2</p>
        <h1>轉介競賽計分小工具</h1>
        <p className="subtitle">上傳 Excel 後統計參與轉介藥局，分類區域後進行距離加權計分。</p>
      </section>

      <section className="panel guide-panel">
        <div className="step-title">
          <span>使用規範</span>
          <h2>操作教學</h2>
        </div>
        <div className="guide-grid">
          <div>
            <h3>Excel 欄位</h3>
            <ol>
              <li>第 1 列可放標題，系統從第 2 列開始讀取。</li>
              <li>A 欄為聽力中心，B 欄為轉介藥局。</li>
              <li>C 欄為左耳聽損，D 欄為右耳聽損。</li>
            </ol>
          </div>
          <div>
            <h3>資料規則</h3>
            <ol>
              <li>聽力中心或藥局空白會直接略過。</li>
              <li>左右耳不是數字會以 0 計算。</li>
              <li>左右耳超出 -10 到 120 會列為異常資料，不納入藥局統計。</li>
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
            <h3>區域分類</h3>
            <ol>
              <li>每間藥局預設為未分類。</li>
              <li>每間藥局只能選一個區域。</li>
              <li>名稱只差空白或大小寫時會合併為同一間藥局。</li>
            </ol>
          </div>
        </div>
      </section>

      <section className="panel">
        <div className="step-title">
          <span>Step 1</span>
          <h2>上傳 Excel</h2>
        </div>
        <label className="upload-box">
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={handleFileChange}
            disabled={isLoading}
          />
          <span>{isLoading ? "解析中..." : "選擇 Excel 檔案"}</span>
          {fileName ? <strong>{fileName}</strong> : null}
        </label>
        {error ? <p className="error">{error}</p> : null}
      </section>

      {stats ? (
        <section className="panel">
          <div className="step-title">
            <span>Step 2</span>
            <h2>摘要資訊</h2>
          </div>
          <div className="stats-grid">
            <StatCard label="總資料筆數" value={stats.totalRows} />
            <StatCard label="有效資料筆數" value={stats.validRows} />
            <StatCard label="異常資料筆數" value={stats.invalidRows} />
            <StatCard label="藥局數量" value={stats.pharmacyCount} />
            <StatCard label="聽力中心數量" value={centerCount || stats.centerCount} />
          </div>
          {invalidRows.length > 0 ? (
            <p className="notice">
              有 {invalidRows.length} 筆左右耳數值超出 -10 到 120，已排除於藥局統計。
            </p>
          ) : null}
        </section>
      ) : null}

      {summaries.length > 0 ? (
        <section className="panel">
          <div className="step-title">
            <span>Step 3</span>
            <h2>尚未加權距離，請分類各藥局參賽區域</h2>
          </div>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>藥局名稱</th>
                  <th>轉介筆數</th>
                  <th>0分</th>
                  <th>3分</th>
                  <th>10分</th>
                  <th>聽損總分</th>
                  <th>平均分</th>
                  <th>區域</th>
                </tr>
              </thead>
              <tbody>
                {summaries.map((summary) => (
                  <tr key={summary.name}>
                    <td>{summary.name}</td>
                    <td>{summary.referralCount}</td>
                    <td>{summary.hearingBonus0Count}</td>
                    <td>{summary.hearingBonus3Count}</td>
                    <td>{summary.hearingBonus10Count}</td>
                    <td>{summary.hearingScoreTotal}</td>
                    <td>{summary.avgScore.toFixed(1)}</td>
                    <td>
                      <select
                        value={summary.region}
                        onChange={(event) =>
                          handleRegionChange(summary.name, event.target.value)
                        }
                      >
                        {REGIONS.map((region) => (
                          <option key={region} value={region}>
                            {region}
                          </option>
                        ))}
                      </select>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* V2: Weighted distance button */}
          <div className="distance-action">
            {!allClassified ? (
              <p className="distance-hint">
                ⚠ 請先完成所有藥局的區域分類
              </p>
            ) : null}
            <button
              className="btn-distance"
              disabled={!allClassified || isCalculating}
              onClick={handleCalculateDistance}
            >
              {isCalculating ? "計算中..." : "加權距離"}
            </button>
          </div>

          {distanceError ? (
            <p className="notice">{distanceError}</p>
          ) : null}
        </section>
      ) : null}

      {/* V2: Weighted results */}
      {resultsByRegion ? (
        <section className="panel">
          <div className="step-title">
            <span>Step 4</span>
            <h2>加權後排名結果</h2>
          </div>

          <div className="region-tabs">
            {ACTIVE_REGIONS.map((region) => (
              <button
                key={region}
                className={`tab-btn ${activeTab === region ? "tab-active" : ""}`}
                onClick={() => setActiveTab(region)}
              >
                {region}
                <span className="tab-count">
                  {resultsByRegion[region]?.length ?? 0}
                </span>
              </button>
            ))}
          </div>

          {resultsByRegion[activeTab] &&
          resultsByRegion[activeTab].length > 0 ? (
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>排名</th>
                    <th>藥局名稱</th>
                    <th>參賽區域</th>
                    <th>轉介筆數</th>
                    <th>聽損總分</th>
                    <th>距離加權分</th>
                    <th>加權後總分</th>
                  </tr>
                </thead>
                <tbody>
                  {resultsByRegion[activeTab].map((w, index) => (
                    <tr key={w.name}>
                      <td className="rank-cell">{index + 1}</td>
                      <td>{w.name}</td>
                      <td>{w.region}</td>
                      <td>{w.referralCount}</td>
                      <td>{w.hearingScoreTotal}</td>
                      <td>{w.distanceScoreTotal}</td>
                      <td className="total-cell">{w.weightedTotalScore}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="empty-region">此區域沒有藥局參賽。</p>
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
