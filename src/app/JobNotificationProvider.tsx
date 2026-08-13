import { useEffect, useRef } from "react";
import { useAppDialog } from "@/shared/components/dialogs/app/useAppDialog";
import { getJobLabel, jobSuccessMessages } from "@/app/jobMeta";
import { useContext } from "react";
import { JobContext } from "@/app/jobContext";
import { jobRpc } from "@/api/rpc";

export function JobNotificationProvider({ children }: { children: React.ReactNode }) {
  const dialog = useAppDialog();
  const context = useContext(JobContext);
  if (!context) throw new Error("JobNotificationProvider must be used inside JobProvider");
  const { pendingNotifications, dismissNotification } = context;
  const showingJobIdRef = useRef<string | null>(null);
  const pending = pendingNotifications[0];

  useEffect(() => {
    if (!pending || showingJobIdRef.current) return;
    showingJobIdRef.current = pending.jobId;

    const notify = async () => {
      const label = getJobLabel(pending.jobType);
      if (pending.progress.status === "error") {
        await dialog.error({
          title: `${label} 실패`,
          message: pending.progress.error || pending.progress.message || `${label} 중 오류가 발생했습니다.`,
          detail: pending.progress.detail,
        });
      } else if (pending.progress.status === "cancelled") {
        await dialog.confirm({
          title: `${label} 취소`,
          message: pending.progress.message || `${label} 작업이 취소되었습니다.`,
        });
      } else {
        await dialog.confirm({
          title: `${label} 완료`,
          message: jobSuccessMessages[pending.jobType] || pending.progress.message || `${label}이 완료되었습니다.`,
        });
      }
      await jobRpc.acknowledge(pending.jobId);
      dismissNotification(pending.jobId);
      showingJobIdRef.current = null;
    };

    void notify();
  }, [pending, dialog, dismissNotification]);

  return children;
}
