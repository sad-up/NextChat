import { getServerSideConfig } from "@/app/config/server";
import { ModelProvider, OPENAI_BASE_URL } from "@/app/constant";
import { NextRequest, NextResponse } from "next/server";
import { auth } from "../auth";

export const runtime = "nodejs";

export async function OPTIONS() {
  return NextResponse.json({ body: "OK" }, { status: 200 });
}

export async function POST(req: NextRequest) {
  const formData = await req.formData();
  const file = formData.get("file");

  if (!(file instanceof File)) {
    return NextResponse.json(
      { error: { message: "No file provided" } },
      { status: 400 },
    );
  }

  const authResult = auth(req, ModelProvider.GPT);
  if (authResult.error) {
    return NextResponse.json(authResult, {
      status: 401,
    });
  }

  const serverConfig = getServerSideConfig();
  const authHeader = req.headers.get("Authorization");
  const apiKey = authHeader?.trim()
    ? authHeader
    : serverConfig.apiKey
    ? `Bearer ${serverConfig.apiKey}`
    : "";

  if (!apiKey) {
    return NextResponse.json(
      { error: { message: "Missing OpenAI API key" } },
      { status: 401 },
    );
  }

  let baseUrl = serverConfig.baseUrl || OPENAI_BASE_URL;
  if (!baseUrl.startsWith("http")) {
    baseUrl = `https://${baseUrl}`;
  }
  if (baseUrl.endsWith("/")) {
    baseUrl = baseUrl.slice(0, -1);
  }
  if (baseUrl.endsWith("/v1")) {
    baseUrl = baseUrl.slice(0, -3);
  }

  const body = new FormData();
  body.append("purpose", "user_data");
  body.append("file", file);

  const uploadUrl = `${baseUrl}/v1/files`;
  const response = await fetch(uploadUrl, {
    method: "POST",
    headers: {
      Authorization: apiKey,
      ...(serverConfig.openaiOrgId && {
        "OpenAI-Organization": serverConfig.openaiOrgId,
      }),
    },
    body,
  });

  const text = await response.text();

  if (!response.ok) {
    let errorBody: unknown = text;
    try {
      errorBody = JSON.parse(text);
    } catch {}

    return NextResponse.json(
      {
        error: {
          message:
            typeof errorBody === "object" &&
            errorBody &&
            "error" in errorBody &&
            typeof (errorBody as any).error?.message === "string"
              ? (errorBody as any).error.message
              : response.statusText || "OpenAI file upload failed",
          status: response.status,
          uploadUrl,
          body: errorBody,
        },
      },
      { status: response.status },
    );
  }

  return new NextResponse(text, {
    status: response.status,
    headers: {
      "Content-Type":
        response.headers.get("content-type") || "application/json",
    },
  });
}
