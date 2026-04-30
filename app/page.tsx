"use client";

import { ChangeEvent, useMemo, useState } from "react";
import {
  getExcelDataRowCount,
  groupByPharmacy,
  isExcelFile,
  parseExcel,
  PharmacySummary,
  ReferralRow,
} from "../lib/referrals";

type UploadStats = {
  totalRows: number;
  validRows: number;
  pharmacyCount: number;
  centerCount: number;
};

const REGIONS = ["未分類", "桃區", "竹苗區", "宜花區", "中彰投區"];

export default function Home() {
  const [rows, setRows] = useState<ReferralRow[]>([]);
  const [summaries, setSummaries] = useState<PharmacySummary[]>([]);
  const [stats, setStats] = useState<UploadStats | null>(null);
  const [error, setError] = useState("");
  const [fileName, setFileName] = useState("");
  const [isLoading, setIsLoading] = useState(false);

  const centerCount = useMemo(
    () => new Set(rows.map((row) => row.center)).size,
    [rows],
  );

  async function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";

    if (!file) return;

    setError("");
    setIsLoading(true);
    setRows([]);
    setSummaries([]);
    setStats(null);
    setFileName(file.name);

    try {
      if (!isExcelFile(file)) {
        throw new Error("檔案不是 Excel，請上傳 .xlsx、.xls 或 .xlsm 檔。");
      }

      const parsedRows = await parseExcel(file);

      if (parsedRows.length === 0) {
        throw new Error("空資料：沒有有效轉介資料，請確認聽力中心與藥局名稱都有填寫。");
      }

      const pharmacySummaries = groupByPharmacy(parsedRows);
      const totalRows = await getExcelDataRowCount(file);

      setRows(parsedRows);
      setSummaries(pharmacySummaries);
      setStats({
        totalRows,
        validRows: parsedRows.length,
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
  }

  return (
    <main className="page">
      <section className="header">
        <p className="eyebrow">第一階段 MVP</p>
        <h1>轉介競賽計分小工具</h1>
        <p className="subtitle">上傳 Excel 後統計參與轉介藥局，並手動分類區域。</p>
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
            <StatCard label="藥局數量" value={stats.pharmacyCount} />
            <StatCard label="聽力中心數量" value={centerCount || stats.centerCount} />
          </div>
        </section>
      ) : null}

      {summaries.length > 0 ? (
        <section className="panel">
          <div className="step-title">
            <span>Step 3</span>
            <h2>藥局列表</h2>
          </div>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>藥局名稱</th>
                  <th>轉介筆數</th>
                  <th>0分筆數</th>
                  <th>3分筆數</th>
                  <th>10分筆數</th>
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
