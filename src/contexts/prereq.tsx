// src/contexts/prereq.tsx
import React, { createContext, useCallback, useContext, useState } from "react";
import { rpc } from "pyloid-js";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import PrereqSetupView from "@/views/PrereqSetupView";

type State = {
  ok: boolean;
  has_class: boolean;
  has_data: boolean;
  has_student: boolean;
  data_file_name?: string;
  cwd: string;
  data_dir: string;
  missing: string[];
};

type Ctx = {
  /** 실행 전 필수 파일 점검. 모달이 필요한 경우 띄우고 false 반환 */
  enforcePrereq: () => Promise<boolean>;
  /** 수동으로 모달 오픈(디버그/메뉴용) */
  openPrereq: () => Promise<void>;
};

const PrereqCtx = createContext<Ctx | null>(null);
export const usePrereq = () => {
  const ctx = useContext(PrereqCtx);
  if (!ctx) throw new Error("PrereqProvider로 감싸야 합니다.");
  return ctx;
};

export function PrereqProvider({ children }: { children: React.ReactNode }) {
  const [open, setOpen] = useState(false);
  const [state, setState] = useState<State | null>(null);
  const [loading, setLoading] = useState(false);

  const fetchState = useCallback(async () => {
    setLoading(true);
    try {
      const res = await rpc.call("check_data_files", {});
      setState(res);
      return res as State;
    } catch {
      // 브라우저 단독 실행 등 RPC 불가 시엔 통과
      const res = {
        ok: true, has_class: true, has_data: true, has_student: true,
        missing: [], cwd: "", data_dir: "", data_file_name: ""
      } as State;
      setState(res);
      return res;
    } finally {
      setLoading(false);
    }
  }, []);

  const openPrereq = useCallback(async () => {
    await fetchState();
    setOpen(true);
  }, [fetchState]);

  const enforcePrereq = useCallback(async () => {
    const res = await fetchState();
    if (res?.ok) return true;
    setOpen(true);
    return false;
    // 모달에서 설치 완료 후 "다시 확인" 누르면 자동 갱신/닫힘
  }, [fetchState]);

  // 설치 뷰 내부에서 "다시 확인" → 상태 OK면 자동 닫힘
  const handleRefresh = useCallback(async () => {
    const res = await fetchState();
    if (res?.ok) setOpen(false);
  }, [fetchState]);

  return (
    <PrereqCtx.Provider value={{ enforcePrereq, openPrereq }}>
      {children}
      <Dialog open={open} onOpenChange={setOpen}>
        <DialogContent className="max-w-4xl">
          <DialogHeader>
            <DialogTitle>필수 파일 설치</DialogTitle>
          </DialogHeader>

          {loading || !state ? (
            <div className="py-10 text-center text-sm text-muted-foreground">확인 중…</div>
          ) : (
            <PrereqSetupView state={state} onRefresh={handleRefresh} />
          )}
        </DialogContent>
      </Dialog>
    </PrereqCtx.Provider>
  );
}
