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
      const url = new URL(
        "https://maps.googleapis.com/maps/api/distancematrix/json",
      );
      url.searchParams.set("origins", pair.center);
      url.searchParams.set("destinations", pair.pharmacy);
      url.searchParams.set("mode", "driving");
      url.searchParams.set("language", "zh-TW");
      url.searchParams.set("key", apiKey);

      const response = await fetch(url.toString());
      const data = await response.json();

      if (data.status !== "OK") {
        results.push({
          center: pair.center,
          pharmacy: pair.pharmacy,
          distanceKm: null,
          error: `API 狀態: ${data.status}`,
        });
        continue;
      }

      const element = data.rows?.[0]?.elements?.[0];

      if (!element || element.status !== "OK") {
        results.push({
          center: pair.center,
          pharmacy: pair.pharmacy,
          distanceKm: null,
          error: `無法計算距離: ${element?.status ?? "UNKNOWN"}`,
        });
        continue;
      }

      // distance.value is in meters
      const distanceKm = element.distance.value / 1000;

      results.push({
        center: pair.center,
        pharmacy: pair.pharmacy,
        distanceKm,
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
