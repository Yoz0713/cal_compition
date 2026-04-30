import * as XLSX from "xlsx";
import type { StoreAddressMap, WeightedPharmacySummary } from "./referrals";
import { getStoreAddress, VALID_REGIONS } from "./referrals";

export type ExportVerificationItem = {
  center: string;
  centerAddress: string;
  pharmacy: string;
  distanceKm: number;
  distanceScore: number;
  isSameStore: boolean;
  status: "success" | "same_store" | "address_missing" | "distance_failed";
  errorMessage: string;
  pharmacyAddress: string;
};

/**
 * Build and download an Excel workbook with:
 *   - Dynamic region ranking sheets (only regions with data)
 *   - 距離計算核對結果 sheet
 *   - 原始明細資料 sheet
 */
export function exportWeightedExcel(
  weightedResults: WeightedPharmacySummary[],
  allDetails: ExportVerificationItem[],
  rows: { center: string; pharmacy: string; leftEar: number; rightEar: number; hearingScore: number; region: string }[],
  addressMap: StoreAddressMap,
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

  // ─── 2. 距離計算核對結果 ────────────────────────────────────────

  const apiSheetData = [
    [
      "聽力中心門市",
      "聽力中心地址",
      "藥局門市",
      "藥局地址",
      "距離 km",
      "距離分",
      "是否同店",
      "狀態",
      "錯誤原因",
    ],
    ...allDetails.map((d) => [
      d.center,
      d.centerAddress || "—",
      d.pharmacy,
      d.pharmacyAddress || "—",
      d.status === "address_missing" || d.status === "distance_failed" ? "—" : Number(d.distanceKm.toFixed(1)),
      d.distanceScore,
      d.isSameStore ? "是" : "否",
      formatDistanceStatus(d.status),
      d.errorMessage || "—",
    ]),
  ];

  const apiWs = XLSX.utils.aoa_to_sheet(apiSheetData);
  apiWs["!cols"] = [
    { wch: 22 }, // 聽力中心門市
    { wch: 40 }, // 聽力中心地址
    { wch: 22 }, // 藥局門市
    { wch: 40 }, // 藥局地址
    { wch: 12 }, // 距離 km
    { wch: 8 },  // 距離分
    { wch: 10 }, // 是否同店
    { wch: 14 }, // 狀態
    { wch: 30 }, // 錯誤原因
  ];
  XLSX.utils.book_append_sheet(wb, apiWs, "距離計算核對結果");

  // ─── 3. 原始明細資料 ────────────────────────────────────────────

  const rawSheetData = [
    ["聽力中心", "聽力中心地址", "轉介藥局", "藥局地址", "左耳聽損", "右耳聽損", "聽損分數", "參賽區域"],
    ...rows.map((r) => [
      r.center,
      getStoreAddress(addressMap, r.center) || "—",
      r.pharmacy,
      getStoreAddress(addressMap, r.pharmacy) || "—",
      r.leftEar,
      r.rightEar,
      r.hearingScore,
      r.region,
    ]),
  ];

  const rawWs = XLSX.utils.aoa_to_sheet(rawSheetData);
  rawWs["!cols"] = [
    { wch: 24 },
    { wch: 40 },
    { wch: 24 },
    { wch: 40 },
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

function formatDistanceStatus(status: ExportVerificationItem["status"]) {
  if (status === "success") return "成功";
  if (status === "same_store") return "同店";
  if (status === "address_missing") return "地址缺失";
  return "距離計算失敗";
}
