// src/lib/progress.ts
import { useEffect, useRef, useState } from "react";
import { rpc } from "pyloid-js";

export type ProgressLevel = "info" | "success" | "warning" | "error";
export type ProgressStatus = "running" | "done" | "error" | "unknown";

export type ProgressPayload = {
  step: number;
  total: number;
  level: ProgressLevel;
  status: ProgressStatus;
  message: string;
  ts: number;
};

export const initialProgress: ProgressPayload = {
  step: 0, total: 0, level: "info", status: "unknown", message: "", ts: 0,
};

export function useProgressPoller(jobId?: string, interval = 500) {
  const [prog, setProg] = useState<ProgressPayload>(initialProgress);
  const timer = useRef<number | null>(null);

  useEffect(() => {
    if (!jobId){
      setProg(initialProgress);
      return;
    }
    if (timer.current) clearInterval(timer.current);

    const tick = async () => {
      try {
        const p = await rpc.call("get_progress", { job_id: jobId });
        setProg({
          step: Number(p?.step ?? 0),
          total: Number(p?.total ?? 0),
          level: (p?.level ?? "info") as ProgressLevel,
          status: (p?.status ?? "unknown") as ProgressStatus,
          message: String(p?.message ?? ""),
          ts: Number(p?.ts ?? Date.now()),
        });
      } catch (e) {
        setProg((prev) => ({ ...prev, status: "error", message: String(e) }));
      }
    };

    tick();
    timer.current = window.setInterval(tick, interval);

    return () => {
      if (timer.current) {
        clearInterval(timer.current);
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
