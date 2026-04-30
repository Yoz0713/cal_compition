import { NextRequest, NextResponse } from "next/server";

type DistanceRequest = {
  centerName: string;
  centerAddress: string;
  pharmacyName: string;
  pharmacyAddress: string;
};

type DistanceResponse = {
  centerName: string;
  centerAddress: string;
  pharmacyName: string;
  pharmacyAddress: string;
  distanceKm?: number;
  distanceScore: number;
  isSameStore: boolean;
  status: "success" | "same_store" | "address_missing" | "distance_failed";
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

function createAddressMissingResponse(
  body: DistanceRequest,
): DistanceResponse {
  return {
    centerName: body.centerName,
    centerAddress: body.centerAddress,
    pharmacyName: body.pharmacyName,
    pharmacyAddress: body.pharmacyAddress,
    distanceKm: 0,
    distanceScore: 0,
    isSameStore: false,
    status: "address_missing",
    errorMessage: "地址對應表找不到此門市地址，請回 Excel 補上",
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

  const { centerName, centerAddress, pharmacyName, pharmacyAddress } = body;

  if (normalizeName(centerName) === normalizeName(pharmacyName)) {
    const result: DistanceResponse = {
      centerName,
      centerAddress,
      pharmacyName,
      pharmacyAddress,
      distanceKm: 0,
      distanceScore: 0.5,
      isSameStore: true,
      status: "same_store",
    };
    return NextResponse.json(result);
  }

  if (!centerAddress || !pharmacyAddress) {
    return NextResponse.json(createAddressMissingResponse(body));
  }

  try {
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
            address: centerAddress,
          },
          destination: {
            address: pharmacyAddress,
          },
          travelMode: "DRIVE",
          languageCode: "zh-TW",
          regionCode: "TW",
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
        centerAddress,
        pharmacyName,
        pharmacyAddress,
        distanceKm: 0,
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
        centerAddress,
        pharmacyName,
        pharmacyAddress,
        distanceKm: 0,
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
      centerAddress,
      pharmacyName,
      pharmacyAddress,
      distanceKm,
      distanceScore,
      isSameStore: false,
      status: "success",
    };
    return NextResponse.json(result);
  } catch (fetchError) {
    const result: DistanceResponse = {
      centerName,
      centerAddress,
      pharmacyName,
      pharmacyAddress,
      distanceKm: 0,
      distanceScore: 0,
      isSameStore: false,
      status: "distance_failed",
      errorMessage:
        fetchError instanceof Error
          ? fetchError.message
          : "Google Routes API 查詢失敗",
    };
    return NextResponse.json(result);
  }
}
