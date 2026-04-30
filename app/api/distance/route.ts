import { NextRequest, NextResponse } from "next/server";

type DistanceRequest = {
  centerName: string;
  pharmacyName: string;
};

type DistanceResponse = {
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

function normalizeName(name: string): string {
  return name.replace(/\s+/g, "").toLowerCase();
}

function getDistanceScore(distanceKm: number, isSameStore: boolean): number {
  if (isSameStore) return 0.5;
  if (distanceKm <= 4) return 1;
  if (distanceKm <= 9) return 2;
  if (distanceKm <= 45) return 3;
  if (distanceKm <= 65) return 4;
  return 5;
}

async function geocode(
  address: string,
  apiKey: string,
): Promise<{ lat: number; lng: number; resolvedAddress: string } | null> {
  const url = new URL("https://maps.googleapis.com/maps/api/geocode/json");
  url.searchParams.set("address", `大樹藥局${address}`);
  url.searchParams.set("language", "zh-TW");
  url.searchParams.set("region", "tw");
  url.searchParams.set("key", apiKey);

  const response = await fetch(url.toString());
  const data = await response.json();

  if (data.status !== "OK" || !data.results?.[0]) {
    return null;
  }

  const result = data.results[0];
  return {
    lat: result.geometry.location.lat,
    lng: result.geometry.location.lng,
    resolvedAddress: result.formatted_address,
  };
}

export async function POST(request: NextRequest) {
  const apiKey = process.env.GOOGLE_MAPS_API_KEY;

  if (!apiKey) {
    return NextResponse.json(
      { error: "Google Maps API Key 未設定" },
      { status: 500 },
    );
  }

  let body: DistanceRequest;

  try {
    body = await request.json();
    if (!body.centerName || !body.pharmacyName) {
      return NextResponse.json(
        { error: "請提供 centerName 與 pharmacyName" },
        { status: 400 },
      );
    }
  } catch {
    return NextResponse.json({ error: "請求格式錯誤" }, { status: 400 });
  }

  const { centerName, pharmacyName } = body;

  // 1. Same-store check
  if (normalizeName(centerName) === normalizeName(pharmacyName)) {
    const result: DistanceResponse = {
      centerName,
      pharmacyName,
      distanceKm: 0,
      distanceScore: 0.5,
      isSameStore: true,
      status: "same_store",
    };
    return NextResponse.json(result);
  }

  try {
    // 2. Geocoding
    const [centerGeo, pharmacyGeo] = await Promise.all([
      geocode(centerName, apiKey),
      geocode(pharmacyName, apiKey),
    ]);

    if (!centerGeo || !pharmacyGeo) {
      const missing: string[] = [];
      if (!centerGeo) missing.push(`聽力中心「${centerName}」`);
      if (!pharmacyGeo) missing.push(`藥局「${pharmacyName}」`);

      const result: DistanceResponse = {
        centerName,
        pharmacyName,
        centerResolvedAddress: centerGeo?.resolvedAddress,
        pharmacyResolvedAddress: pharmacyGeo?.resolvedAddress,
        distanceScore: 0,
        isSameStore: false,
        status: "place_not_found",
        errorMessage: `查無地點：${missing.join("、")}`,
      };
      return NextResponse.json(result);
    }

    // 3. Routes API
    const routeResponse = await fetch(
      "https://routes.googleapis.com/directions/v2:computeRoutes",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Goog-Api-Key": apiKey,
          "X-Goog-FieldMask": "routes.distanceMeters",
        },
        body: JSON.stringify({
          origin: {
            location: {
              latLng: { latitude: centerGeo.lat, longitude: centerGeo.lng },
            },
          },
          destination: {
            location: {
              latLng: { latitude: pharmacyGeo.lat, longitude: pharmacyGeo.lng },
            },
          },
          travelMode: "DRIVE",
          languageCode: "zh-TW",
        }),
      },
    );

    if (!routeResponse.ok) {
      const errBody = await routeResponse.json().catch(() => ({}));
      const errMsg =
        (errBody as { error?: { message?: string } }).error?.message ??
        `HTTP ${routeResponse.status}`;

      const result: DistanceResponse = {
        centerName,
        pharmacyName,
        centerResolvedAddress: centerGeo.resolvedAddress,
        pharmacyResolvedAddress: pharmacyGeo.resolvedAddress,
        distanceScore: 0,
        isSameStore: false,
        status: "distance_failed",
        errorMessage: `Routes API 錯誤：${errMsg}`,
      };
      return NextResponse.json(result);
    }

    const routeData = await routeResponse.json();
    const distanceMeters = routeData.routes?.[0]?.distanceMeters;

    if (distanceMeters == null) {
      const result: DistanceResponse = {
        centerName,
        pharmacyName,
        centerResolvedAddress: centerGeo.resolvedAddress,
        pharmacyResolvedAddress: pharmacyGeo.resolvedAddress,
        distanceScore: 0,
        isSameStore: false,
        status: "distance_failed",
        errorMessage: "無法計算行車距離：找不到路線",
      };
      return NextResponse.json(result);
    }

    const distanceKm = distanceMeters / 1000;
    const distanceScore = getDistanceScore(distanceKm, false);

    const result: DistanceResponse = {
      centerName,
      pharmacyName,
      centerResolvedAddress: centerGeo.resolvedAddress,
      pharmacyResolvedAddress: pharmacyGeo.resolvedAddress,
      distanceKm,
      distanceScore,
      isSameStore: false,
      status: "success",
    };
    return NextResponse.json(result);
  } catch (fetchError) {
    const result: DistanceResponse = {
      centerName,
      pharmacyName,
      distanceScore: 0,
      isSameStore: false,
      status: "distance_failed",
      errorMessage:
        fetchError instanceof Error
          ? fetchError.message
          : "Google API 查詢失敗",
    };
    return NextResponse.json(result);
  }
}
