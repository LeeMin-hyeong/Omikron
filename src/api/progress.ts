import { baseAPI } from "pyloid-js";

export type ProgressLevel = "info" | "success" | "warning" | "error";
export type ProgressStatus = "running" | "done" | "error" | "cancelled" | "unknown";

export type ProgressPayload = {
  step: number;
  total: number;
  phase_step?: number | null;
  phase_total?: number | null;
  level: ProgressLevel;
  status: ProgressStatus;
  message: string;
  error?: string;
  code?: string;
  detail?: string;
  warnings: string[];
  ts: number;
  revision?: number;
};

export const initialProgress: ProgressPayload = {
  step: 0,
  total: 0,
  phase_step: null,
  phase_total: null,
  level: "info",
  status: "unknown",
  message: "",
  error: "",
  detail: "",
  warnings: [],
  ts: 0,
  revision: 0,
};

let rpcEndpoint: string | null = null;
let rpcWindowId: string | null = null;

export async function rpcCallWithTimeout<T>(
  method: string,
  params: Record<string, unknown>,
  timeoutMs: number,
): Promise<T> {
  if (!rpcEndpoint) rpcEndpoint = await baseAPI.getServerUrl();
  if (!rpcWindowId) rpcWindowId = await baseAPI.getWindowId();

  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(rpcEndpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ jsonrpc: "2.0", method, params, id: rpcWindowId }),
      signal: controller.signal,
    });
    const data = await response.json();
    if (!response.ok) throw new Error(`RPC HTTP ${response.status}`);
    if (data?.error) throw new Error(data.error.message || "RPC 요청에 실패했습니다.");
    return data?.result as T;
  } finally {
    window.clearTimeout(timeoutId);
  }
}

export function normalizeProgress(value: Record<string, unknown>): ProgressPayload {
  return {
    step: Number(value.step ?? 0),
    total: Number(value.total ?? 0),
    phase_step: value.phase_step == null ? null : Number(value.phase_step),
    phase_total: value.phase_total == null ? null : Number(value.phase_total),
    level: (value.level ?? "info") as ProgressLevel,
    status: (value.status ?? "unknown") as ProgressStatus,
    message: String(value.message ?? ""),
    error: value.error == null ? "" : String(value.error),
    code: value.code == null ? "" : String(value.code),
    detail: value.detail == null ? "" : String(value.detail),
    warnings: Array.isArray(value.warnings) ? value.warnings.map(String) : [],
    ts: Number(value.ts ?? Date.now()),
    revision: Number(value.revision ?? 0),
  };
}
