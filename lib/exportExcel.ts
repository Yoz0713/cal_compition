import * as XLSX from "xlsx";
import type { WeightedPharmacySummary, DistanceResult } from "./referrals";
import { VALID_REGIONS } from "./referrals";

/**
 * Build and download an Excel workbook with:
 *   - Dynamic region ranking sheets (only regions with data)
 *   - Google API 核對結果 sheet
 *   - 原始明細資料 sheet
 */
export function exportWeightedExcel(
  weightedResults: WeightedPharmacySummary[],
  allDetails: DistanceResult[],
  rows: { center: string; pharmacy: string; leftEar: number; rightEar: number; hearingScore: number; region: string }[],
) {
  const wb = XLSX.utils.book_new();

  // ─── 1. Dynamic region sheets ────────────────────────────────────

  // Determine which regions actually have data, in fixed order
  const activeRegions = VALID_REGIONS.filter((region) =>
    weightedResults.some((w) => w.region === region),
  );

  for (const region of activeRegions) {
    const regionData = weightedResults
      .filter((w) => w.region === region)
      .sort((a, b) => {
        // 1. 總分高者優先
        if (b.weightedTotalScore !== a.weightedTotalScore)
          return b.weightedTotalScore - a.weightedTotalScore;
        // 2. 聽損總分高者優先
        if (b.hearingScoreTotal !== a.hearingScoreTotal)
          return b.hearingScoreTotal - a.hearingScoreTotal;
        // 3. 總轉介人數高者優先
        if (b.referralCount !== a.referralCount)
          return b.referralCount - a.referralCount;
        // 4. 藥局名稱字典序
        return a.name.localeCompare(b.name, "zh-Hant");
      });

    const sheetData = [
      [
        "排名",
        "藥局名稱",
        "總轉介人數",
        "30分貝以下人次",
        "30分貝以上人次",
        "50分貝以上人次",
        "聽損總分",
        "距離加權分數",
        "總分",
      ],
      ...regionData.map((w, i) => [
        i + 1,
        w.name,
        w.referralCount,
        w.hearingBonus0Count,
        w.hearingBonus3Count,
        w.hearingBonus10Count,
        w.hearingScoreTotal,
        w.distanceScoreTotal,
        w.weightedTotalScore,
      ]),
    ];

    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    // Set column widths
    ws["!cols"] = [
      { wch: 6 },  // 排名
      { wch: 20 }, // 藥局名稱
      { wch: 12 }, // 總轉介人數
      { wch: 16 }, // 30分貝以下
      { wch: 16 }, // 30分貝以上
      { wch: 16 }, // 50分貝以上
      { wch: 10 }, // 聽損總分
      { wch: 14 }, // 距離加權分數
      { wch: 8 },  // 總分
    ];

    XLSX.utils.book_append_sheet(wb, ws, region);
  }

  // ─── 2. Google API 核對結果 ──────────────────────────────────────

  const apiSheetData = [
    ["#", "聽力中心", "轉介藥局", "距離(km)", "距離分", "同店"],
    ...allDetails.map((d, i) => [
      i + 1,
      d.center,
      d.pharmacy,
      Number(d.distanceKm.toFixed(1)),
      d.distanceScore,
      d.isSameStore ? "是" : "否",
    ]),
  ];

  const apiWs = XLSX.utils.aoa_to_sheet(apiSheetData);
  apiWs["!cols"] = [
    { wch: 6 },
    { wch: 24 },
    { wch: 24 },
    { wch: 12 },
    { wch: 8 },
    { wch: 6 },
  ];
  XLSX.utils.book_append_sheet(wb, apiWs, "Google API 核對結果");

  // ─── 3. 原始明細資料 ────────────────────────────────────────────

  const rawSheetData = [
    ["聽力中心", "轉介藥局", "左耳聽損", "右耳聽損", "聽損分數", "參賽區域"],
    ...rows.map((r) => [
      r.center,
      r.pharmacy,
      r.leftEar,
      r.rightEar,
      r.hearingScore,
      r.region,
    ]),
  ];

  const rawWs = XLSX.utils.aoa_to_sheet(rawSheetData);
  rawWs["!cols"] = [
    { wch: 24 },
    { wch: 24 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 12 },
  ];
  XLSX.utils.book_append_sheet(wb, rawWs, "原始明細資料");

  // ─── Download ────────────────────────────────────────────────────

  const now = new Date();
  const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
  XLSX.writeFile(wb, `轉介競賽報表_${dateStr}.xlsx`);
}
