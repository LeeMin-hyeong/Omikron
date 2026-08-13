import { createContext, useContext } from "react";

export type AppDialogOptions = {
  title?: string;
  message?: string;
  detail?: string;
  confirmText?: string;
  cancelText?: string;
  allowOutsideClose?: boolean;
  blockReplacement?: boolean;
};

export type AppDialogContextValue = {
  warning: (opts?: AppDialogOptions) => Promise<boolean>;
  error: (opts?: AppDialogOptions) => Promise<void>;
  confirm: (opts?: AppDialogOptions) => Promise<void>;
  update: (opts: Partial<AppDialogOptions>) => void;
};

export const AppDialogContext = createContext<AppDialogContextValue | null>(null);

export function useAppDialog() {
  const context = useContext(AppDialogContext);
  if (!context) throw new Error("AppDialogProvider로 감싸야 합니다.");
  return context;
}
