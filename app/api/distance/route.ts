import { NextRequest, NextResponse } from "next/server";

type DistancePair = {
  center: string;
  pharmacy: string;
};

type DistanceResponse = {
  center: string;
  pharmacy: string;
  distanceKm: number | null;
  error?: string;
};

export async function POST(request: NextRequest) {
  const apiKey = process.env.GOOGLE_MAPS_API_KEY;

  if (!apiKey) {
    return NextResponse.json(
      { error: "Google Maps API Key 未設定" },
      { status: 500 },
    );
  }

  let pairs: DistancePair[];

  try {
    const body = await request.json();
    pairs = body.pairs;

    if (!Array.isArray(pairs) || pairs.length === 0) {
      return NextResponse.json(
        { error: "請提供至少一組 center/pharmacy" },
        { status: 400 },
      );
    }
  } catch {
    return NextResponse.json(
      { error: "請求格式錯誤" },
      { status: 400 },
    );
  }

  const results: DistanceResponse[] = [];

  // Process each pair sequentially to avoid hitting rate limits
  for (const pair of pairs) {
    try {
      const response = await fetch(
        "https://routes.googleapis.com/directions/v2:computeRoutes",
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "X-Goog-Api-Key": apiKey,
            "X-Goog-FieldMask": "routes.distanceMeters",
          },
          body: JSON.stringify({
            origin: { address: `大樹藥局${pair.center}` },
            destination: { address: `大樹藥局${pair.pharmacy}` },
            travelMode: "DRIVE",
            languageCode: "zh-TW",
          }),
        },
      );

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        const errMsg =
          (errBody as { error?: { message?: string } }).error?.message ??
          `HTTP ${response.status}`;
        results.push({
          center: pair.center,
          pharmacy: pair.pharmacy,
          distanceKm: null,
          error: `API 錯誤: ${errMsg}`,
        });
        continue;
      }

      const data = await response.json();
      const distanceMeters = data.routes?.[0]?.distanceMeters;

      if (distanceMeters == null) {
        results.push({
          center: pair.center,
          pharmacy: pair.pharmacy,
          distanceKm: null,
          error: "無法計算距離：找不到路線",
        });
        continue;
      }

      results.push({
        center: pair.center,
        pharmacy: pair.pharmacy,
        distanceKm: distanceMeters / 1000,
      });
    } catch (fetchError) {
      results.push({
        center: pair.center,
        pharmacy: pair.pharmacy,
        distanceKm: null,
        error:
          fetchError instanceof Error
            ? fetchError.message
            : "Google API 查詢失敗",
      });
    }
  }

  return NextResponse.json({ results });
}
