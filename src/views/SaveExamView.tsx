// src/views/SaveExamView.tsx
import { rpc } from "pyloid-js";
import { useEffect, useRef, useState } from "react";
import type { ViewProps } from "@/types/omikron";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { CheckCircle2, Circle, FileUp, Loader2, Play, X } from "lucide-react";
import { fileToBase64 } from "@/utils/rpc";
import { usePrereq } from "@/contexts/prereq";

// 공통 진행 훅/시작 함수 (앞서 만든 표준 유틸)
import { startJob, useProgressPoller, type ProgressPayload } from "@/lib/progress";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";

export default function SaveExamView({ meta, onAction }: ViewProps) {
  const dialog = useAppDialog();

  // ===== 파일 업로드(단일) =====
  const [file, setFile] = useState<File | null>(null);
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const acceptExt = /\.(xlsx|xlsm|xls|csv)$/i;

  function handleFiles(list: FileList | null) {
    if (!list || list.length === 0) return;
    const arr = Array.from(list);
    const first = arr.find((f) => acceptExt.test(f.name));
    if (first) setFile(first);
  }

  // ===== 진행/폴링 =====
  const [jobId, setJobId] = useState<string>();
  const prog: ProgressPayload = useProgressPoller(jobId); // {step,total,level,status,message,...}
  const [doneCount, setDoneCount] = useState(0);
  const running = prog.status === "running";

  // step -> 완료 단계 개수로 간주하여 UI 반영
  useEffect(() => {
    const max = meta.steps.length;
    const clamped = Math.max(0, Math.min(max, Number(prog.step ?? 0)));
    setDoneCount(clamped);
  }, [prog.step, meta.steps.length]);

  // 완료/에러 마무리
  useEffect(() => {
    if (prog.status === "done") {
      alert("작업이 완료되었습니다.");
    } else if (prog.status === "error") {
      if (prog.message) alert(prog.message);
    }
  }, [prog.status]);

  // ===== 시작 =====
  const start = async () => {
    if (running || !file) return;

    try {
      onAction?.("save-exam");

      // 파일 → base64
      const b64 = await fileToBase64(file);

      // 작업 시작 → job_id
      const id = await startJob("start_save_exam", { filename: file.name, b64 });
      setJobId(id);
      setDoneCount(0);
    } catch (e: any) {
      alert(String(e?.message || e));
    }
  };

  return (
    <div className="h-full min-h-0 min-w-0">
      <Card className="relative h-full min-h-0 rounded-2xl border-border/80 shadow-sm py-0">
        <CardContent className="h-full min-h-0 p-1">
          {/* 상/하 1:1 + 하단 버튼 */}
          <div className="grid h-full min-h-0 grid-rows-[1fr_auto] gap-1">
            {/* 상/하 콘텐츠 */}
            <div className="grid min-h-0 grid-rows-2 gap-1 p-1">
              {/* 상단: 설명 + 드롭존 */}
              <div className="flex min-h-0 flex-col rounded-xl border border-border/60 bg-card/50 p-3 text-card-foreground">
                <h3 className="mb-1 text-base font-semibold">작업 설명</h3>
                <Separator className="mb-1" />
                <p className="text-sm text-muted-foreground">{meta.guide}</p>

                {/* Dropzone / Selected File */}
                <div
                  role="button"
                  tabIndex={0}
                  onClick={() => !running && inputRef.current?.click()}
                  onKeyDown={(e) => {
                    if (running) return;
                    if (e.key === "Enter" || e.key === " ") inputRef.current?.click();
                  }}
                  onDragOver={(e) => {
                    if (running) return;
                    e.preventDefault();
                    setDragging(true);
                  }}
                  onDragLeave={() => !running && setDragging(false)}
                  onDrop={(e) => {
                    if (running) return;
                    e.preventDefault();
                    setDragging(false);
                    handleFiles(e.dataTransfer.files);
                  }}
                  className={`mt-2 flex h-full min-h-0 flex-col justify-center rounded-xl border transition ${
                    dragging ? "border-point bg-point/5" : "border-dashed border-border/70 hover:bg-accent/40"
                  } ${running ? "pointer-events-none opacity-90" : "cursor-pointer"}`}
                >
                  {!file ? (
                    <div className="flex flex-col items-center justify-center px-3 py-2 text-center">
                      <FileUp className="mb-1 h-5 w-5" />
                      <div className="text-sm">
                        파일을 여기로 끌어다 놓거나 <span className="underline">클릭하여 선택</span>
                      </div>
                      <div className="text-xs text-muted-foreground">허용: .xlsx .xlsm .xls .csv</div>
                    </div>
                  ) : (
                    <div className="flex items-center justify-between gap-3 px-3 py-2">
                      <div className="flex min-w-0 items-center gap-2">
                        <FileUp className="h-4 w-4" />
                        <span className="truncate" title={file.name}>
                          {file.name}
                        </span>
                      </div>
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          className="inline-flex h-7 w-7 items-center justify-center rounded-md hover:bg-accent/50 disabled:opacity-50"
                          onClick={(e) => {
                            e.stopPropagation();
                            if (!running) setFile(null);
                          }}
                          disabled={running}
                          aria-label="remove"
                        >
                          <X className="h-4 w-4" />
                        </button>
                      </div>
                    </div>
                  )}

                  {/* 숨겨진 파일 입력 */}
                  <input
                    ref={inputRef}
                    type="file"
                    accept=".xlsx,.xlsm,.xls,.csv"
                    className="hidden"
                    onChange={(e) => handleFiles(e.target.files)}
                  />

                  {file && !running && (
                    <div className="px-3 pb-2 text-center text-xs text-muted-foreground">
                      새 파일을 드래그하면 <span className="font-medium">교체</span>됩니다.
                    </div>
                  )}
                </div>
              </div>

              {/* 하단: 단계 진행 */}
              <div className="min-h-0 rounded-xl border border-border/60 bg-card/50 p-3 text-card-foreground">
                <h3 className="mb-1 text-base font-semibold">진행 단계</h3>
                <Separator className="mb-1" />
                {/* 최신 메시지(있을 때) */}
                {prog.message && (
                  <div className="mb-2 text-xs text-muted-foreground">
                    {prog.message}
                    {prog.total > 0 && prog.step > 0 ? ` (${prog.step}/${prog.total})` : prog.step > 0 ? ` (${prog.step})` : ""}
                  </div>
                )}
                <ol className="space-y-1">
                  {meta.steps.map((stepLabel: string, idx: number) => {
                    const complete = idx < doneCount;
                    const current = idx === doneCount && running;
                    return (
                      <li key={idx} className="flex items-center gap-2 text-sm">
                        {complete ? (
                          <CheckCircle2 className="h-4 w-4 text-green-600" />
                        ) : current ? (
                          <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />
                        ) : (
                          <Circle className="h-4 w-4 text-muted-foreground" />
                        )}
                        <span
                          className={
                            complete
                              ? "font-medium text-green-700"
                              : current
                              ? "text-foreground"
                              : "text-muted-foreground"
                          }
                        >
                          {idx + 1}. {stepLabel}
                        </span>
                      </li>
                    );
                  })}
                </ol>
              </div>
            </div>

            {/* Footer: 우측 하단 실행 버튼 */}
            <div className="flex shrink-0 items-center justify-end gap-2 p-1">
              <Button
                className="rounded-xl bg-black text-white"
                onClick={start}
                disabled={!file || running}
                title={!file ? "파일을 먼저 선택하세요" : running ? "진행 중…" : "실행"}
              >
                {running ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Play className="mr-2 h-4 w-4" />}
                {running ? "실행 중…" : "실행"}
              </Button>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
