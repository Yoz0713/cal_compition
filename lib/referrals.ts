import * as XLSX from "xlsx";

export type ReferralRow = {
  center: string;
  pharmacy: string;
  leftEar: number;
  rightEar: number;
  hearingScore: number;
};

export type PharmacySummary = {
  name: string;
  referralCount: number;
  hearingBonus0Count: number;
  hearingBonus3Count: number;
  hearingBonus10Count: number;
  hearingScoreTotal: number;
  avgScore: number;
  region: string;
  rows: ReferralRow[];
};

export type ParsedExcelResult = {
  rows: ReferralRow[];
  invalidRows: ReferralRow[];
};

const EXCEL_TYPES = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "application/vnd.ms-excel.sheet.macroEnabled.12",
];

export function getHearingScore(left: number, right: number): 0 | 3 | 10 {
  if (left >= 55 || right >= 55) return 10;
  if (left >= 30 || right >= 30) return 3;
  return 0;
}

export function normalizeName(name: string): string {
  return name.replace(/\s+/g, "").toLowerCase();
}

export function isExcelFile(file: File): boolean {
  const lowerName = file.name.toLowerCase();
  return (
    EXCEL_TYPES.includes(file.type) ||
    lowerName.endsWith(".xlsx") ||
    lowerName.endsWith(".xls") ||
    lowerName.endsWith(".xlsm")
  );
}

export function toNumber(value: unknown): number {
  const numberValue = Number(value);
  return Number.isFinite(numberValue) ? numberValue : 0;
}

export async function parseExcel(file: File): Promise<ParsedExcelResult> {
  if (!isExcelFile(file)) {
    throw new Error("檔案不是 Excel，請上傳 .xlsx、.xls 或 .xlsm 檔。");
  }

  const rawRows = await readFirstSheetRows(file);

  if (rawRows.length <= 1) {
    throw new Error("Excel 沒有可解析的資料列。");
  }

  return rawRows.slice(1).reduce<ParsedExcelResult>((result, row) => {
    const center = String(row[0] ?? "").trim();
    const pharmacy = String(row[1] ?? "").trim();

    if (!center || !pharmacy) return result;

    const leftEar = toNumber(row[2]);
    const rightEar = toNumber(row[3]);
    const isOutOfRange =
      leftEar < -10 || leftEar > 120 || rightEar < -10 || rightEar > 120;

    if (isOutOfRange) {
      result.invalidRows.push({
        center,
        pharmacy,
        leftEar,
        rightEar,
        hearingScore: 0,
      });

      return result;
    }

    result.rows.push({
      center,
      pharmacy,
      leftEar,
      rightEar,
      hearingScore: getHearingScore(leftEar, rightEar),
    });

    return result;
  }, { rows: [], invalidRows: [] });
}

export async function getExcelDataRowCount(file: File): Promise<number> {
  const rawRows = await readFirstSheetRows(file);
  return Math.max(rawRows.length - 1, 0);
}

export function groupByPharmacy(rows: ReferralRow[]): PharmacySummary[] {
  const summaries = new Map<string, PharmacySummary>();

  rows.forEach((row) => {
    const key = normalizeName(row.pharmacy);
    const current =
      summaries.get(key) ??
      ({
        name: row.pharmacy,
        referralCount: 0,
        hearingBonus0Count: 0,
        hearingBonus3Count: 0,
        hearingBonus10Count: 0,
        hearingScoreTotal: 0,
        avgScore: 0,
        region: "未分類",
        rows: [],
      } satisfies PharmacySummary);

    current.referralCount += 1;
    current.rows.push(row);

    if (row.hearingScore === 10) current.hearingBonus10Count += 1;
    else if (row.hearingScore === 3) current.hearingBonus3Count += 1;
    else current.hearingBonus0Count += 1;

    current.hearingScoreTotal =
      current.hearingBonus3Count * 3 + current.hearingBonus10Count * 10;
    current.avgScore =
      current.referralCount === 0
        ? 0
        : current.hearingScoreTotal / current.referralCount;

    summaries.set(key, current);
  });

  return Array.from(summaries.values()).sort((a, b) =>
    a.name.localeCompare(b.name, "zh-Hant"),
  );
}

async function readFirstSheetRows(file: File): Promise<unknown[][]> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    throw new Error("Excel 格式錯誤：找不到工作表。");
  }

  return XLSX.utils.sheet_to_json<unknown[]>(
    workbook.Sheets[firstSheetName],
    {
      header: 1,
      blankrows: false,
    },
  );
}

// ─── V2: Distance scoring ───────────────────────────────────────────

export type DistanceResult = {
  center: string;
  pharmacy: string;
  distanceKm: number;
  distanceScore: number;
  isSameStore: boolean;
};

export type WeightedPharmacySummary = PharmacySummary & {
  distanceScoreTotal: number;
  weightedTotalScore: number;
  distanceDetails: DistanceResult[];
};

export function getDistanceScore(
  distanceKm: number,
  isSameStore: boolean,
): number {
  if (isSameStore) return 0.5;
  if (distanceKm <= 4) return 1;
  if (distanceKm <= 9) return 2;
  if (distanceKm <= 45) return 3;
  if (distanceKm <= 65) return 4;
  return 5;
}
