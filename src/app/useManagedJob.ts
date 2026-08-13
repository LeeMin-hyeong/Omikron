import { useCallback, useContext } from "react";
import { initialProgress } from "@/api/progress";
import { JobContext } from "@/app/jobContext";

export function useManagedJob(jobType: string) {
  const context = useContext(JobContext);
  if (!context) throw new Error("useManagedJob must be used inside JobProvider");
  const entry = context.jobs[jobType];
  const clearManagedJob = context.clearJob;
  const clearJob = useCallback(() => clearManagedJob(jobType), [clearManagedJob, jobType]);
  return {
    jobId: entry?.jobId,
    prog: entry?.progress ?? initialProgress,
    beginOperation: context.beginOperation,
    endOperation: context.endOperation,
    startJob: context.startJob,
    clearJob,
  };
}
