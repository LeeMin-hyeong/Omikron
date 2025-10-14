import { useEffect, useRef, useState } from "react"
import type { ViewProps } from "@/types/omikron"
import { Button } from "@/components/ui/button"
import { Card, CardContent } from "@/components/ui/card"
import { Separator } from "@/components/ui/separator"
import { AlertTriangle, FileSpreadsheet, FileUp, Loader2, Play, Square, SquareCheck, X } from "lucide-react"
import { fileToBase64 } from "@/utils/rpc"
import { usePrereq } from "@/contexts/prereq"
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider"
import { startJob, useProgressPoller, type ProgressPayload, type ProgressStatus } from "@/lib/progress"
import { ScrollArea } from "@/components/ui/scroll-area"
import React from "react"

export default function SendExamMessageView({ meta, onAction }: ViewProps) {
  const dialog = useAppDialog()
  const { enforcePrereq } = usePrereq()

  const [file, setFile] = useState<File | null>(null)
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement | null>(null)

  const [jobId, setJobId] = useState<string>()
  const prog: ProgressPayload = useProgressPoller(jobId)
  const running = prog.status === "running"

  const [doneCount, setDoneCount] = useState(0)
  const lastStatusRef = useRef<ProgressStatus>("unknown")

  // 경고 메시지 누적 리스트
  const [warnings, setWarnings] = useState<string[]>([])

  const acceptExt = /\.(xlsx|xlsm|xls|csv)$/i

  const handleFiles = (list: FileList | null) => {
    if (!list || list.length === 0 || running) return
    const arr = Array.from(list)
    const first = arr.find((f) => acceptExt.test(f.name))
    if (first) setFile(first)
  }

  const start = async () => {
    if (running || !file) return

    const ok = await enforcePrereq()
    if (!ok) return

    try {
      onAction?.("send-exam-message")
      const b64 = await fileToBase64(file)
      const id = await startJob("start_send_exam_message", { filename: file.name, b64 })
      setJobId(id)
      lastStatusRef.current = "running"
      setDoneCount(0)
      setWarnings([]) // 새 작업 시작 시 경고 초기화
    } catch (err: any) {
      setJobId(undefined)
      await dialog.error({
        title: "작업 실패",
        message: String(err?.message || err),
      })
    }
  }

  // 단계 카운트 갱신
  useEffect(() => {
    const max = meta.steps.length
    const clamped = Math.max(0, Math.min(max, Number(prog.step ?? 0)))
    setDoneCount(clamped)
  }, [prog.step, meta.steps.length])

  // 경고 수집
  useEffect(() => {
    if (!running) return
    if (prog.level !== "warning") return
    if (!prog.message) return

    setWarnings((prev) => {
      // 같은 메시지가 연속으로 들어올 때 중복 방지
      if (prev.length > 0 && prev[prev.length - 1] === prog.message) return prev
      return [...prev, prog.message ]
    })
  }, [running, prog.level, prog.message, prog.ts])

  // 상태 전이 처리
  useEffect(() => {
    if (!jobId) {
      lastStatusRef.current = "unknown"
      return
    }

    if (prog.status === "running") {
      lastStatusRef.current = "running"
      return
    }

    if (prog.status === "error" && lastStatusRef.current !== "error") {
      lastStatusRef.current = "error"
      void dialog
        .error({
          title: "오류",
          message: prog.message || "메시지 작성 중 오류가 발생했습니다.",
        })
        .then(() => {
          setJobId(undefined)
          if (inputRef.current) inputRef.current.value = ""
          setFile(null)
        })
    } else if (prog.status === "done" && lastStatusRef.current !== "done") {
      lastStatusRef.current = "done"
      void dialog
        .confirm({
          title: "메시지 작성 성공",
          message: prog.message || "메시지 작성을 완료했습니다.",
        })
        .then(() => {
          setJobId(undefined)
          if (inputRef.current) inputRef.current.value = ""
          setFile(null)
        })
    }
  }, [jobId, prog.status, prog.message, dialog])

  useEffect(() => {
    if (running) setDragging(false)
  }, [running])

  return (
    <div className="h-full min-h-0 min-w-0">
      <Card className="relative h-full min-h-0 rounded-2xl border-border/80 shadow-sm py-0">
        <CardContent className="h-full min-h-0 p-1">
          <div className="grid h-full min-h-0 grid-rows-[minmax(0,1fr)_minmax(0,1fr)_auto] gap-1">
            <div className="min-h-0 overflow-hidden rounded-xl border border-border/60 bg-card/50">
              <div className="flex h-full min-h-0 flex-col p-3">
                <h3 className="mb-1 text-base font-semibold">작업 안내</h3>
                <Separator className="mb-1" />
                <p className="text-sm text-muted-foreground">{meta.guide}</p>
                <div
                  role="button"
                  tabIndex={0}
                  onClick={() => !running && inputRef.current?.click()}
                  onKeyDown={(e) => {
                    if (running) return
                    if (e.key === "Enter" || e.key === " ") inputRef.current?.click()
                  }}
                  onDragOver={(e) => {
                    if (running) return
                    e.preventDefault()
                    setDragging(true)
                  }}
                  onDragLeave={() => !running && setDragging(false)}
                  onDrop={(e) => {
                    if (running) return
                    e.preventDefault()
                    setDragging(false)
                    handleFiles(e.dataTransfer.files)
                  }}
                  className={`relative mt-2 flex h-full min-h-0 flex-col items-center justify-center
                              rounded-xl border transition
                              ${dragging ? "border-point bg-point/5" : "border-dashed border-border/70 hover:bg-accent/40"}
                              ${running ? "pointer-events-none opacity-90" : "cursor-pointer"}`}
                >
                  {file && (
                    <button
                      type="button"
                      className="absolute right-2 top-2 inline-flex h-7 w-7 items-center justify-center
                                rounded-md hover:bg-accent/60"
                      onClick={(e) => {
                        e.stopPropagation()
                        if (!running) setFile(null)
                        if (inputRef.current) inputRef.current.value = ""
                      }}
                      aria-label="파일 제거"
                    >
                      <X className="h-4 w-4" />
                    </button>
                  )}
                  {!file ? (
                    <div className="flex flex-col items-center justify-center px-3 py-2 text-center">
                      <FileUp className="mb-1 h-6 w-6" />
                      <div className="text-sm">
                        파일을 끌어오거나 <span className="underline">클릭하여 선택</span>
                      </div>
                      <div className="text-xs text-muted-foreground">지원 형식: .xlsx .xlsm .xls .csv</div>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center justify-center px-3 py-2 text-center">
                      <FileSpreadsheet className="mb-1 h-6 w-6 text-green-600" />
                      <div className="mt-1 truncate text-sm" title={file.name}>
                        {file.name}
                      </div>
                      <div className="mt-1 text-xs text-muted-foreground">
                        다른 파일을 끌어오거나 클릭하면 <span className="font-medium">교체</span>됩니다.
                      </div>
                    </div>
                  )}
                  <input
                    ref={inputRef}
                    type="file"
                    accept=".xlsx,.xlsm,.xls,.csv"
                    className="hidden"
                    onChange={(e) => handleFiles(e.target.files)}
                  />
                </div>
              </div>
            </div>

            <div className="grid min-h-0 grid-cols-[minmax(0,1fr)_minmax(0,1fr)] gap-2 overflow-hidden">
              <div className="min-h-0 overflow-hidden rounded-xl border border-border/60 bg-card/50">
                <div className="flex h-full min-h-0 flex-col p-3">
                  <h3 className="mb-1 text-base font-semibold">진행 단계</h3>
                  <Separator className="mb-1" />
                  <ol className="space-y-1">
                    {meta.steps.map((stepLabel: string, idx: number) => {
                      const complete = idx < doneCount
                      const current = idx === doneCount && running
                      return (
                        <li key={idx} className="flex items-start gap-2 rounded-md border
                                                border-black-500/40 bg-white-50
                                                px-2 py-1 text-sm text-amber-900">
                          <div className="flex items-center gap-2 text-sm">
                            {complete ? (
                              <SquareCheck className="h-4 w-4 text-green-600" />
                            ) : current ? (
                              <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />
                            ) : (
                              <Square className="h-4 w-4 text-muted-foreground" />
                            )}
                            <span className={complete ? "font-medium text-green-700" : current ? "text-foreground" : "text-muted-foreground"}>
                              {stepLabel}
                            </span>
                          </div>
                        </li>
                      )
                    })}
                  </ol>
                  <div className="flex-1" />
                </div>
              </div>

              <div className="min-h-0 overflow-hidden rounded-xl border border-border/60 bg-card/50">
                <div className="flex h-full min-h-0 flex-col p-3 overflow-auto">
                  <h3 className="mb-1 text-base font-semibold">경고 메시지</h3>
                  <Separator className="mb-1" />
                  {warnings.length === 0 ? (
                    <div className="flex h-full min-h-0 items-center justify-center text-sm text-muted-foreground">
                      표시할 경고가 없습니다.
                    </div>
                  ) : (
                    <div>
                      <ScrollArea className="flex-1 h-58 pr-3">
                        <ul className="space-y-1">
                          {warnings.map((m, i) => (
                            <React.Fragment key={i}>
                              <div className="flex items-start gap-2 rounded-md border border-amber-500/40 bg-amber-50 px-2 py-1 text-xs text-amber-900">
                                <AlertTriangle className="mt-[2px] h-4 w-4 text-amber-600" />
                                <p className="w-0 flex-1 break-all">
                                  {m}
                                </p>
                              </div>
                            </React.Fragment>
                          ))}
                        </ul>
                      </ScrollArea>
                    </div>
                  )}
                </div>
              </div>
            </div>

            <div className="flex shrink-0 items-center justify-end gap-2 p-1">
              <Button
                className="rounded-xl bg-black text-white"
                onClick={start}
                disabled={!file || running}
                title={!file ? "파일을 먼저 선택하세요" : running ? "진행 중입니다" : "실행"}
              >
                {running ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Play className="mr-2 h-4 w-4" />}
                {running ? "진행 중" : "실행"}
              </Button>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  )
}
