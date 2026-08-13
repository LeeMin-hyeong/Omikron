// src/components/app-dialog/AppDialogProvider.tsx
import React, { useCallback, useMemo, useRef, useState } from "react";
import { Dialog, DialogContent, DialogFooter, DialogHeader, DialogTitle, DialogDescription } from "@/shared/components/ui/dialog";
import { Button } from "@/shared/components/ui/button";
import { AlertTriangle, XCircle, CircleHelp } from "lucide-react";
import {
  AppDialogContext,
  type AppDialogContextValue,
  type AppDialogOptions,
} from "@/shared/components/dialogs/app/useAppDialog";

type Kind = "warning" | "error" | "confirm";

type InternalState = {
  kind: Kind;
  opts: Required<AppDialogOptions>;
  resolve: (value: boolean | undefined) => void;
};

const DEFAULT_OPTIONS: Required<AppDialogOptions> = {
  title: "",
  message: "",
  detail: "",
  confirmText: "확인",
  cancelText: "취소",
  allowOutsideClose: false,
  blockReplacement: false,
};

export function AppDialogProvider({ children }: { children: React.ReactNode }) {
  const [open, setOpen] = useState(false);
  const [state, setState] = useState<InternalState | null>(null);
  const stateRef = useRef<InternalState | null>(null);
  const pendingRef = useRef<InternalState[]>([]);

  const show = useCallback((kind: Kind, opts?: AppDialogOptions) =>
    new Promise<boolean | undefined>((resolve) => {
      const next = {
        kind,
        resolve,
        opts: { ...DEFAULT_OPTIONS, ...(opts || {}) },
      };
      if (stateRef.current?.opts.blockReplacement) {
        pendingRef.current.push(next);
        return;
      }
      stateRef.current = next;
      setState(next);
      setOpen(true);
    }), []);

  const finish = (value: boolean | undefined) => {
    const current = stateRef.current;
    current?.resolve(value);

    if (current?.opts.blockReplacement && value === true) {
      for (const pending of pendingRef.current.splice(0)) {
        pending.resolve(pending.kind === "warning" ? false : undefined);
      }
    }

    const next = pendingRef.current.shift() ?? null;
    stateRef.current = next;
    setState(next);
    setOpen(next !== null);
  };

  const onClose = () => {
    // warning은 닫힘을 "취소"로 처리, 나머지는 resolve()
    finish(state?.kind === "warning" ? false : undefined);
  };

  const confirm = () => {
    finish(state?.kind === "warning" ? true : undefined);
  };

  const value = useMemo<AppDialogContextValue>(
    () => ({
      warning: async (opts) => Boolean(await show("warning", opts)),
      error:   (opts) => show("error",   opts).then(() => {}),
      confirm: (opts) => show("confirm", opts).then(() => {}),
      update: (opts) => {
        const current = stateRef.current;
        if (!current) return;
        const next = { ...current, opts: { ...current.opts, ...opts } };
        stateRef.current = next;
        setState(next);
      },
    }),
    [show]
  );

  const tone =
    state?.kind === "warning" ? { icon: <AlertTriangle className="h-5 w-5 text-amber-600" />, headerCls: "text-amber-700" } :
    state?.kind === "error"   ? { icon: <XCircle className="h-5 w-5 text-rose-600" />,   headerCls: "text-rose-700" } :
                                { icon: <CircleHelp className="h-5 w-5 text-sky-600" />,  headerCls: "text-sky-700" };

  return (
    <AppDialogContext.Provider value={value}>
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
            showCloseButton={false}
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
            <div className="mt-1 flex-1 min-h-0 overflow-auto pr-1 space-y-3">
              {state.opts.message && (
                <DialogDescription className="whitespace-pre-wrap">
                  {state.opts.message}
                </DialogDescription>
              )}
              {state.opts.detail && (
                <details className="rounded-md border bg-slate-50">
                  <summary className="cursor-pointer select-none px-3 py-2 text-sm text-slate-700">
                    여기를 클릭하여 자세한 에러 내용 확인
                  </summary>
                  <div className="border-t px-3 py-2">
                    <pre className="max-h-48 overflow-auto whitespace-pre-wrap text-xs font-mono text-slate-700">
                      {state.opts.detail}
                    </pre>
                  </div>
                </details>
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
    </AppDialogContext.Provider>
  );
}
