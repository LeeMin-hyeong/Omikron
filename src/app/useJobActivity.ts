import { useContext } from "react";
import { JobContext } from "@/app/jobContext";

export function useJobActivity() {
  const context = useContext(JobContext);
  if (!context) throw new Error("useJobActivity must be used inside JobProvider");
  return { isBusy: context.isBusy };
}
