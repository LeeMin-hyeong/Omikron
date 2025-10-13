import { rpc } from "pyloid-js"
import { useEffect, useRef, useState } from "react"
import type { ViewProps } from "@/types/omikron"
import { Button } from "@/components/ui/button"
import { Card, CardContent } from "@/components/ui/card"
import { Separator } from "@/components/ui/separator"
import { CheckCircle2, Circle, FileUp, Play, Square, X } from "lucide-react"
import { fileToBase64 } from "@/utils/rpc"
import { usePrereq } from "@/contexts/prereq"
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider"

// 헤더는 루트 레이아웃에 있으므로 이 뷰에는 포함하지 않음
export default function SendExamMessageView({ meta, onAction }: ViewProps) {
  // ===== Single-file Drag & Drop =====
  const dialog = useAppDialog();
  const { enforcePrereq } = usePrereq();
  const [file, setFile] = useState<File | null>(null)
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement | null>(null)
  const [running, setRunning] = useState(false)
  const [doneCount, setDoneCount] = useState(0)

  const acceptExt = /\.(xlsx|xlsm|xls|csv)$/i

  function handleFiles(list: FileList | null) {
    if (!list || list.length === 0) return
    const arr = Array.from(list)
    const first = arr.find((f) => acceptExt.test(f.name))
    if (first) setFile(first)
  }

  const pollRef = useRef<number | null>(null)
  
  const start = async () => {
    if (running || !file) return;

    // 실행 전 프리체크
    const ok = await enforcePrereq();
    if (!ok) return;

    setRunning(true);
    setDoneCount(0);
    onAction?.("save-exam");

    try {
      // 1) 파일 → base64
      const b64 = await fileToBase64(file);

      // 2) 작업 시작 (job_id 받기)
      const res = await rpc.call("start_save_exam", { filename: file.name, b64 });
      const jobId = res?.job_id;
      if (!jobId) throw new Error("job 시작 실패");

      // 3) 폴링으로 진행상태 업데이트 (0.5초)
      pollRef.current = window.setInterval(async () => {
        try {
          const p: any = await rpc.call("get_progress", { job_id: jobId });
          // p: { step, total, level, status, message, ts }
          const step = Number(p?.step ?? 0);
          const status = String(p?.status ?? "");
          const message = typeof p?.message === "string" ? p.message : "";

          // 프런트 단계 표시(뷰 정의와 길이 맞춤)
          setDoneCount(Math.max(0, Math.min(meta.steps.length, step)));

          if (status === "done") {
            clearInterval(pollRef.current!);
            pollRef.current = null;
            setRunning(false);
            // (선택) 완료 안내
            // await dialog.confirm({ title: "완료", message: "작업이 성공적으로 끝났습니다." });
          } else if (status === "error") {
            clearInterval(pollRef.current!);
            pollRef.current = null;
            setRunning(false);
            await dialog.error({
              title: "오류",
              message: message || "처리 중 오류가 발생했습니다.",
            });
          }
          // status === "running" 인 동안은 진행만 업데이트
          // p.level / p.message 를 실시간 로그로 쓰고 싶으면 여기서 상태에 저장하세요.
        } catch (err: any) {
          clearInterval(pollRef.current!);
          pollRef.current = null;
          setRunning(false);
          await dialog.error({
            title: "통신 오류",
            message: String(err?.message || err || "진행 상태를 가져오지 못했습니다."),
          });
        }
      }, 500);
    } catch (err: any) {
      setRunning(false);
      await dialog.error({
        title: "시작 실패",
        message: String(err?.message || err),
      });
    }
  };
  
  useEffect(() => {
    return () => {
      if (pollRef.current) clearInterval(pollRef.current)
    }
  }, [])
  
    return (
      <div className="h-full min-h-0 min-w-0">
        <Card className="relative h-full min-h-0 rounded-2xl border-border/80 shadow-sm py-0">
          <CardContent className="h-full min-h-0 p-1">
            <div className="grid h-full min-h-0 grid-rows-[1fr_auto] gap-1">
              <div className="grid min-h-0 grid-rows-2 gap-1 p-1">
                <div className="flex min-h-0 flex-col rounded-xl border border-border/60 bg-card/50 p-3 text-card-foreground">
                  <h3 className="mb-1 text-base font-semibold">작업 설명</h3>
                  <Separator className="mb-1" />
                  <p className="text-sm text-muted-foreground">{meta.guide}</p>
  
                  {/* Dropzone / Selected File */}
                  <div
                    role="button"
                    tabIndex={0}
                    onClick={() => inputRef.current?.click()}
                    onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && inputRef.current?.click()}
                    onDragOver={(e) => { e.preventDefault(); setDragging(true) }}
                    onDragLeave={() => setDragging(false)}
                    onDrop={(e) => { e.preventDefault(); setDragging(false); handleFiles(e.dataTransfer.files) }}
                    className={`mt-2 flex h-full min-h-0 flex-col justify-center rounded-xl border transition ${
                      dragging ? "border-point bg-point/5" : "border-dashed border-border/70 hover:bg-accent/40"
                    }`}
                    style={{ cursor: "pointer" }}
                  >
                    {!file ? (
                      <div className="flex flex-col items-center justify-center px-3 py-2 text-center">
                        <FileUp className="mb-1 h-5 w-5" />
                        <div className="text-sm">파일을 여기로 끌어다 놓거나 <span className="underline">클릭하여 선택</span></div>
                        <div className="text-xs text-muted-foreground">허용: .xlsx .xlsm .xls .csv</div>
                      </div>
                    ) : (
                      <div className="flex items-center justify-between gap-3 px-3 py-2">
                        <div className="flex min-w-0 items-center gap-2">
                          <FileUp className="h-4 w-4" />
                          <span className="truncate" title={file.name}>{file.name}</span>
                        </div>
                        <div className="flex items-center gap-2">
                          {/* 클릭하거나 파일 드롭으로 교체 가능 */}
                          <button
                            type="button"
                            className="inline-flex h-7 w-7 items-center justify-center rounded-md hover:bg-accent/50"
                            onClick={(e) => { e.stopPropagation(); setFile(null); }}
                            aria-label="remove"
                          >
                            <X className="h-4 w-4" />
                          </button>
                        </div>
                      </div>
                    )}
  
                    {/* 숨겨진 파일 입력: 클릭/교체용 */}
                    <input
                      ref={inputRef}
                      type="file"
                      accept=".xlsx,.xlsm,.xls,.csv"
                      className="hidden"
                      onChange={(e) => handleFiles(e.target.files)}
                    />
  
                    {file && (
                      <div className="px-3 pb-2 text-center text-xs text-muted-foreground">
                        새 파일을 여기로 드래그하면 <span className="font-medium">교체</span>됩니다.
                      </div>
                    )}
                  </div>
                </div>
  
                {/* 하단: 단계 진행 */}
                <div className="min-h-0 rounded-xl border border-border/60 bg-card/50 p-3 text-card-foreground">
                  <h3 className="mb-1 text-base font-semibold">진행 단계</h3>
                  <Separator className="mb-1" />
                  <ol className="space-y-1">
                    {meta.steps.map((step, idx) => {
                      const done = idx < doneCount
                      return (
                        <li key={idx} className="flex items-center gap-2 text-sm">
                          {done ? (
                            <CheckCircle2 className="h-4 w-4 text-green-600" />
                          ) : (
                            <Circle className="h-4 w-4 text-muted-foreground" />
                          )}
                          <span className={done ? "font-medium text-green-700" : "text-muted-foreground"}>
                            {idx + 1}. {step}
                          </span>
                        </li>
                      )
                    })}
                  </ol>
                </div>
              </div>
  
              {/* Footer: 우측 하단 버튼 (항상 보임) */}
              <div className="flex shrink-0 items-center justify-end gap-2 p-1">
                {!running ? (
                  <Button
                    className="rounded-xl bg-black text-white"
                    onClick={start}
                    disabled={!file}
                    title={!file ? "파일을 먼저 선택하세요" : "실행"}
                  >
                    <Play className="h-4 w-4" /> 실행
                  </Button>
                ) : (
                  <Button variant="destructive" className="rounded-xl" onClick={() => stop()}>
                    <Square className="h-4 w-4" /> 중지
                  </Button>
                )}
              </div>
            </div>
          </CardContent>
        </Card>
      </div>
    )
}
