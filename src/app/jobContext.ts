import { createContext } from "react";
import type { ProgressPayload } from "@/api/progress";

export type JobEntry = {
  jobId: string;
  jobType: string;
  progress: ProgressPayload;
};

export type JobContextValue = {
  jobs: Record<string, JobEntry>;
  pendingNotifications: JobEntry[];
  isBusy: boolean;
  beginOperation: (operation: string) => void;
  endOperation: (operation: string) => void;
  startJob: (method: string, params?: Record<string, unknown>) => Promise<string>;
  clearJob: (jobType: string) => void;
  dismissNotification: (jobId: string) => void;
};

export const JobContext = createContext<JobContextValue | null>(null);
