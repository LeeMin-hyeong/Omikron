import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  initialProgress,
  normalizeProgress,
} from "@/api/progress";
import { JobContext, type JobEntry } from "@/app/jobContext";
import { jobRpc, type JobType } from "@/api/rpc";

export function JobProvider({ children }: { children: React.ReactNode }) {
  const [jobs, setJobs] = useState<Record<string, JobEntry>>({});
  const [pendingNotifications, setPendingNotifications] = useState<JobEntry[]>([]);
  const [activeOperations, setActiveOperations] = useState<Set<string>>(() => new Set());
  const jobsRef = useRef(jobs);
  const unchangedCount = useRef(0);
  const pollInFlightRef = useRef(false);
  const [pollGeneration, setPollGeneration] = useState(0);
  jobsRef.current = jobs;

  const startJob = useCallback(async (method: string, params: Record<string, unknown> = {}) => {
    const existing = jobsRef.current[method];
    if (existing?.progress.status === "running") return existing.jobId;
    const response = await jobRpc.start(method as JobType, params as Record<string, never>);
    if (!response.ok) {
      throw new Error(String(response.error ?? "작업을 시작하지 못했습니다."));
    }
    const jobId = response.ok ? response.data.jobId : "";
    if (!jobId) throw new Error("작업 ID를 받지 못했습니다.");
    const jobType = response.ok ? response.data.jobType : method;
    setJobs((previous) => ({
      ...previous,
      [jobType]: {
        jobId,
        jobType,
        progress: { ...initialProgress, status: "running", message: "작업 대기 중..." },
      },
    }));
    setPollGeneration((generation) => generation + 1);
    return jobId;
  }, []);

  const clearJob = useCallback((jobType: string) => {
    setJobs((previous) => {
      if (!previous[jobType]) return previous;
      const next = { ...previous };
      delete next[jobType];
      return next;
    });
    setPollGeneration((generation) => generation + 1);
  }, []);

  const beginOperation = useCallback((operation: string) => {
    setActiveOperations((current) => new Set(current).add(operation));
  }, []);

  const endOperation = useCallback((operation: string) => {
    setActiveOperations((current) => {
      if (!current.has(operation)) return current;
      const next = new Set(current);
      next.delete(operation);
      return next;
    });
  }, []);

  const dismissNotification = useCallback((jobId: string) => {
    setPendingNotifications((current) => current.filter((item) => item.jobId !== jobId));
  }, []);

  useEffect(() => {
    let cancelled = false;
    let timer: number | undefined;

    const schedule = (delay: number) => {
      if (!cancelled) timer = window.setTimeout(poll, delay);
    };

    const poll = async () => {
      if (pollInFlightRef.current) {
        schedule(100);
        return;
      }
      const active = Object.values(jobsRef.current).filter(
        (job) => job.progress.status === "running" || job.progress.status === "unknown",
      );
      if (active.length === 0) return;
      pollInFlightRef.current = true;
      try {
        const result = await jobRpc.getBatch(
          active.map((job) => ({
              jobId: job.jobId,
              revision: job.progress.revision ?? 0,
          })),
        );
        if (cancelled) return;
        if (!result.ok) throw new Error(result.error);
        const changedJobs = result.data.jobs.filter((item) => item.changed && item.state);
        unchangedCount.current = changedJobs.length > 0 ? 0 : unchangedCount.current + 1;
        if (changedJobs.length > 0) {
          const completed = changedJobs.flatMap((item) => {
            const state = item.state!;
            const progress = normalizeProgress(state);
            if (
              progress.status !== "done" &&
              progress.status !== "error" &&
              progress.status !== "cancelled"
            ) {
              return [];
            }
            return [{
              jobId: item.jobId,
              jobType: String(state.jobType ?? ""),
              progress,
            }];
          });
          if (completed.length > 0) {
            setPendingNotifications((current) => {
              const known = new Set(current.map((item) => item.jobId));
              const additions = completed.filter((item) => !known.has(item.jobId));
              return [...current, ...additions];
            });
          }
          setJobs((previous) => {
            const next = { ...previous };
            for (const item of changedJobs) {
              const state = item.state!;
              const jobType = String(state.jobType ?? "");
              const current = next[jobType];
              const incoming = normalizeProgress(state);
              if (
                current?.jobId === item.jobId &&
                (incoming.revision ?? 0) > (current.progress.revision ?? 0)
              ) {
                next[jobType] = { ...current, progress: incoming };
              }
            }
            return next;
          });
        }
        const delay = document.hidden
          ? 4_000
          : unchangedCount.current >= 10
            ? 2_000
            : unchangedCount.current >= 3
              ? 1_000
              : 500;
        schedule(delay);
      } catch {
        if (cancelled) return;
        unchangedCount.current += 1;
        schedule(document.hidden ? 5_000 : Math.min(5_000, 1_000 * unchangedCount.current));
      } finally {
        pollInFlightRef.current = false;
      }
    };

    poll();
    return () => {
      cancelled = true;
      if (timer !== undefined) window.clearTimeout(timer);
    };
  }, [pollGeneration]);

  const value = useMemo(
    () => ({
      jobs,
      pendingNotifications,
      isBusy:
        activeOperations.size > 0 ||
        Object.values(jobs).some(
          (job) => job.progress.status === "running" || job.progress.status === "unknown",
        ),
      beginOperation,
      endOperation,
      startJob,
      clearJob,
      dismissNotification,
    }),
    [jobs, pendingNotifications, activeOperations, beginOperation, endOperation, startJob, clearJob, dismissNotification],
  );
  return <JobContext.Provider value={value}>{children}</JobContext.Provider>;
}
