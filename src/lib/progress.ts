// src/lib/progress.ts
import { useEffect, useRef, useState } from "react";
import { rpc } from "pyloid-js";

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
  warnings: [],
  ts: 0,
};

export function useProgressPoller(jobId?: string, interval = 500) {
  const [prog, setProg] = useState<ProgressPayload>(initialProgress);
  const timer = useRef<number | null>(null);
  const failCount = useRef(0);
  const lastSuccess = useRef(0);
  const delayRef = useRef(interval);
  const inFlight = useRef(false);

  useEffect(() => {
    if (!jobId){
      setProg(initialProgress);
      failCount.current = 0;
      lastSuccess.current = 0;
      delayRef.current = interval;
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

    const withTimeout = async <T,>(promise: Promise<T>, ms: number) => new Promise<T>((resolve, reject) => {
      const timeoutId = window.setTimeout(() => reject(new Error("progress poll timeout")), ms);
      promise
        .then((value) => {
          clearTimeout(timeoutId);
          resolve(value);
        })
        .catch((err) => {
          clearTimeout(timeoutId);
          reject(err);
        });
    });

    const tick = async () => {
      if (cancelled) return;
      if (inFlight.current) {
        schedule(delayRef.current);
        return;
      }
      inFlight.current = true;
      try {
        const p = await withTimeout(
          rpc.call("get_progress", { job_id: jobId }),
          Math.max(3000, interval * 4),
        );
        failCount.current = 0;
        lastSuccess.current = Date.now();
        delayRef.current = interval;

        const status = (p?.status ?? "unknown") as ProgressStatus;
        const warnings = Array.isArray(p?.warnings) ? p.warnings.map((w: any) => String(w)) : [];
        setProg({
          step: Number(p?.step ?? 0),
          total: Number(p?.total ?? 0),
          phase_step: p?.phase_step == null ? null : Number(p?.phase_step),
          phase_total: p?.phase_total == null ? null : Number(p?.phase_total),
          level: (p?.level ?? "info") as ProgressLevel,
          status,
          message: String(p?.message ?? ""),
          warnings,
          ts: Number(p?.ts ?? Date.now()),
        });

        if (status === "done" || status === "error") {
          return;
        }
        schedule(delayRef.current);
      } catch (e) {
        failCount.current += 1;
        const now = Date.now();
        if (!lastSuccess.current) lastSuccess.current = now;
        delayRef.current = Math.min(delayRef.current * 2, 5_000);
        setProg((prev) => ({
          ...prev,
          message: prev.message || "진행 상태 확인 중...",
          ts: now,
        }));
        schedule(delayRef.current);
      } finally {
        inFlight.current = false;
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
