import { createContext, useContext } from "react";

export type PrereqContextValue = {
  enforcePrereq: () => Promise<boolean>;
  openPrereq: () => Promise<void>;
};

export const PrereqContext = createContext<PrereqContextValue | null>(null);

export function usePrereq() {
  const context = useContext(PrereqContext);
  if (!context) throw new Error("PrereqProvider로 감싸야 합니다.");
  return context;
}
