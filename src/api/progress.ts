// src/lib/progress.ts
import { useEffect, useRef, useState } from "react";
import { baseAPI, rpc } from "pyloid-js";

export type ProgressLevel = "info" | "success" | "warning" | "error";
export type ProgressStatus = "running" | "done" | "error" | "unknown";

export type ProgressPayload = {
  step: number;
  total: number;
  phase_step?: number | null;
  phase_total?: number | null;
  level: ProgressLevel;
  status: ProgressStatus;
  message: string;
  error?: string;
  detail?: string;
  warnings: string[];
  ts: number;
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
};

let rpcEndpoint: string | null = null;
let rpcWindowId: string | null = null;


async function rpcCallWithTimeout<T>(
  method: string,
  params: Record<string, any>,
  timeoutMs: number,
): Promise<T> {
  if (!rpcEndpoint) rpcEndpoint = await baseAPI.getServerUrl();
  if (!rpcWindowId) rpcWindowId = await baseAPI.getWindowId();

  const controller = new AbortController();
  let timeoutId: number | null = null;
  const timeoutPromise = new Promise<never>((_, reject) => {
    timeoutId = window.setTimeout(() => {
      try {
        controller.abort();
      } finally {
        reject(new Error("progress poll timeout"));
      }
    }, timeoutMs);
  });
  try {
    const fetchPromise = fetch(rpcEndpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        jsonrpc: "2.0",
        method,
        params,
        id: rpcWindowId,
      }),
      signal: controller.signal,
    }).then(async (response) => {
      const text = await response.text();
      let data: any = null;
      try {
        data = text ? JSON.parse(text) : null;
      } catch (e) {
        throw new Error(`RPC non-json response (status ${response.status})`);
      }
      if (!response.ok) {
        throw new Error(`RPC HTTP ${response.status}`);
      }
      if (data?.error) {
        throw new Error(`RPC Error: ${data.error.message} (Code: ${data.error.code})`);
      }
      return data?.result as T;
    });
    return await Promise.race([fetchPromise, timeoutPromise]);
  } catch (e) {
    throw new Error(`RPC unknown error: ${String(e)}`);
  } finally {
    if (timeoutId != null) window.clearTimeout(timeoutId);
  }
}

function isSameProgress(a: ProgressPayload, b: ProgressPayload) {
  if (a === b) return true;
  const aLastWarn = a.warnings[a.warnings.length - 1];
  const bLastWarn = b.warnings[b.warnings.length - 1];
  return (
    a.step === b.step &&
    a.total === b.total &&
    a.phase_step === b.phase_step &&
    a.phase_total === b.phase_total &&
    a.level === b.level &&
    a.status === b.status &&
    a.message === b.message &&
    (a.error ?? "") === (b.error ?? "") &&
    (a.detail ?? "") === (b.detail ?? "") &&
    a.ts === b.ts &&
    a.warnings.length === b.warnings.length &&
    aLastWarn === bLastWarn
  );
}

export function useProgressPoller(jobId?: string, interval = 500) {
  const [prog, setProg] = useState<ProgressPayload>(initialProgress);
  const timer = useRef<number | null>(null);
  const failCount = useRef(0);
  const lastSuccess = useRef(0);
  const delayRef = useRef(interval);
  const inFlight = useRef(false);
  const inFlightStartedAt = useRef(0);
  const requestSeq = useRef(0);
  const activeRequestId = useRef(0);
  const lastPayloadRef = useRef<ProgressPayload>(initialProgress);

  useEffect(() => {
    if (!jobId){
      setProg(initialProgress);
      failCount.current = 0;
      lastSuccess.current = 0;
      delayRef.current = interval;
      lastPayloadRef.current = initialProgress;
      inFlight.current = false;
      inFlightStartedAt.current = 0;
      activeRequestId.current = 0;
      if (timer.current) {
        clearTimeout(timer.current);
        timer.current = null;
      }
      return;
    }
    if (timer.current) clearTimeout(timer.current);
    delayRef.current = interval;
    let cancelled = false;

    const schedule = (delayMs: number) => {
      if (cancelled) return;
      timer.current = window.setTimeout(tick, delayMs);
    };

    const tick = async () => {
      if (cancelled) return;
      if (inFlight.current) {
        const age = Date.now() - inFlightStartedAt.current;
        if (age < Math.max(4_000, interval * 8)) {
          schedule(delayRef.current);
          return;
        }
        inFlight.current = false;
      }
      inFlight.current = true;
      inFlightStartedAt.current = Date.now();
      const requestId = ++requestSeq.current;
      activeRequestId.current = requestId;
      try {
        const p = await rpcCallWithTimeout<Record<string, any>>(
          "get_progress",
          { job_id: jobId },
          Math.max(3000, interval * 4),
        );
        if (activeRequestId.current !== requestId) {
          return;
        }
        failCount.current = 0;
        lastSuccess.current = Date.now();
        delayRef.current = interval;

        const status = (p?.status ?? "unknown") as ProgressStatus;
        const warnings = Array.isArray(p?.warnings) ? p.warnings.map((w: any) => String(w)) : [];
        const nextProg = {
          step: Number(p?.step ?? 0),
          total: Number(p?.total ?? 0),
          phase_step: p?.phase_step == null ? null : Number(p?.phase_step),
          phase_total: p?.phase_total == null ? null : Number(p?.phase_total),
          level: (p?.level ?? "info") as ProgressLevel,
          status,
          message: String(p?.message ?? ""),
          error: p?.error == null ? "" : String(p.error),
          detail: p?.detail == null ? "" : String(p.detail),
          warnings,
          ts: Number(p?.ts ?? Date.now()),
        } as ProgressPayload;
        if (!isSameProgress(lastPayloadRef.current, nextProg)) {
          lastPayloadRef.current = nextProg;
          setProg(nextProg);
        }

        if (status === "done" || status === "error") {
          return;
        }
        schedule(delayRef.current);
      } catch (e) {
        if (activeRequestId.current !== requestId) {
          return;
        }
        failCount.current += 1;
        const now = Date.now();
        if (!lastSuccess.current) lastSuccess.current = now;
        delayRef.current = Math.min(delayRef.current * 2, 5_000);
        setProg((prev) => ({
          ...prev,
          message: prev.message || "진행 상태 확인 중...",
          error: prev.error || "",
          detail: prev.detail || "",
          ts: now,
        }));
        schedule(delayRef.current);
      } finally {
        if (activeRequestId.current === requestId) {
          inFlight.current = false;
          inFlightStartedAt.current = 0;
        }
      }
    };

    tick();

    return () => {
      cancelled = true;
      if (timer.current) {
        clearTimeout(timer.current);
        timer.current = null;
      }
    };
  }, [jobId, interval]);

  return prog;
}

export async function startJob(method: string, params?: Record<string, any>) {
  const res = await rpc.call(method, params ?? {});
  const jobId = res?.job_id ?? res?.jobId;
  if (!jobId) throw new Error("job_id가 없습니다.");
  return jobId as string;
}
