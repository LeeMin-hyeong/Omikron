// src/components/app-dialog/AppDialogProvider.tsx
import React, { createContext, useContext, useMemo, useState } from "react";
import { Dialog, DialogContent, DialogFooter, DialogHeader, DialogTitle, DialogDescription } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { AlertTriangle, XCircle, CircleHelp } from "lucide-react";

type Kind = "warning" | "error" | "confirm";

type BaseOpts = {
  title?: string;
  message?: string;
  confirmText?: string;
  cancelText?: string;
  /** 바깥 클릭/ESC로 닫기 허용 (기본 false) */
  allowOutsideClose?: boolean;
};

type InternalState = {
  kind: Kind;
  opts: Required<Omit<BaseOpts, "allowOutsideClose">> & { allowOutsideClose: boolean };
  resolve: (v: any) => void;
};

type Ctx = {
  warning: (opts?: BaseOpts) => Promise<boolean>; // 확인:true / 취소:false
  error:   (opts?: BaseOpts) => Promise<void>;    // 확인만
  confirm: (opts?: BaseOpts) => Promise<void>;    // 확인만(알림)
};

const AppDialogCtx = createContext<Ctx | null>(null);
export const useAppDialog = () => {
  const ctx = useContext(AppDialogCtx);
  if (!ctx) throw new Error("AppDialogProvider로 감싸야 합니다.");
  return ctx;
};

export function AppDialogProvider({ children }: { children: React.ReactNode }) {
  const [open, setOpen] = useState(false);
  const [state, setState] = useState<InternalState | null>(null);

  const defaults: Required<BaseOpts> = {
    title: "",
    message: "",
    confirmText: "확인",
    cancelText: "취소",
    allowOutsideClose: false,
  };

  const show = (kind: Kind, opts?: BaseOpts) =>
    new Promise<any>((resolve) => {
      setState({
        kind,
        resolve,
        opts: { ...defaults, ...(opts || {}) },
      });
      setOpen(true);
    });

  const onClose = () => {
    // warning은 닫힘을 "취소"로 처리, 나머지는 resolve()
    if (state?.kind === "warning") state.resolve(false);
    else state?.resolve(undefined);
    setOpen(false);
    // cleanup
    setTimeout(() => setState(null), 0);
  };

  const confirm = () => {
    if (state?.kind === "warning") state.resolve(true);
    else state?.resolve(undefined);
    setOpen(false);
    setTimeout(() => setState(null), 0);
  };

  const value = useMemo<Ctx>(
    () => ({
      warning: (opts) => show("warning", opts),
      error:   (opts) => show("error",   opts).then(() => {}),
      confirm: (opts) => show("confirm", opts).then(() => {}),
    }),
    []
  );

  const tone =
    state?.kind === "warning" ? { icon: <AlertTriangle className="h-5 w-5 text-amber-600" />, headerCls: "text-amber-700" } :
    state?.kind === "error"   ? { icon: <XCircle className="h-5 w-5 text-rose-600" />,   headerCls: "text-rose-700" } :
                                { icon: <CircleHelp className="h-5 w-5 text-sky-600" />,  headerCls: "text-sky-700" };

  return (
    <AppDialogCtx.Provider value={value}>
      {children}
      <Dialog
        open={open}
        onOpenChange={(o) => {
          if (o === false) {
            // 바깥 클릭/ESC로 닫힘 허용 여부
            if (state?.opts.allowOutsideClose) onClose();
          }
        }}
      >
        {state && (
          <DialogContent
            className="sm:max-w-md max-h-[80vh] flex flex-col"
            onInteractOutside={(e) => !state.opts.allowOutsideClose && e.preventDefault()}
            onEscapeKeyDown={(e) => !state.opts.allowOutsideClose && e.preventDefault()}
          >
            <DialogHeader className="flex flex-row items-center gap-2 flex-none">
              {tone.icon}
              <DialogTitle className={tone.headerCls}>
                {state.opts.title || (state.kind === "warning" ? "경고" : state.kind === "error" ? "오류" : "확인")}
              </DialogTitle>
            </DialogHeader>

            {/* ✅ 본문만 스크롤 (flex-1 + min-h-0 중요) */}
            <div className="mt-1 flex-1 min-h-0 overflow-auto pr-1">
              {state.opts.message && (
                <DialogDescription className="whitespace-pre-wrap">
                  {state.opts.message}
                </DialogDescription>
              )}
            </div>

            {/* ✅ 푸터는 고정 */}
            <DialogFooter className="mt-2 flex-none">
              {state.kind === "warning" ? (
                <>
                  <Button variant="outline" className="rounded-lg" onClick={onClose}>
                    {state.opts.cancelText}
                  </Button>
                  <Button className="rounded-lg bg-black text-white" onClick={confirm}>
                    {state.opts.confirmText}
                  </Button>
                </>
              ) : (
                <Button className="rounded-lg bg-black text-white" onClick={confirm}>
                  {state.opts.confirmText}
                </Button>
              )}
            </DialogFooter>
          </DialogContent>
        )}
      </Dialog>
    </AppDialogCtx.Provider>
  );
}
