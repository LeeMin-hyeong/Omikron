import { useState } from "react"
import type { ViewProps } from "@/types/omikron"
import { Button } from "@/components/ui/button"
import { Card, CardContent } from "@/components/ui/card"
import { Separator } from "@/components/ui/separator"
import { FileSpreadsheet, Play, Check } from "lucide-react";
import { Spinner } from "@/components/ui/spinner"
import { rpc } from "pyloid-js"
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider"

// 헤더는 루트 레이아웃에 있으므로 이 뷰에는 포함하지 않음
export default function UpdateStudentView({ meta }: ViewProps) {
  const dialog = useAppDialog();

  const [running, setRunning] = useState(false)
  const [done, setDone] = useState(false)

  const start = async () => {
    if (running) return
    setDone(false);

    try {
      setRunning(true);
      const res = await rpc.call("update_student_info", {});
      if(res?.ok){
        setDone(true);
        await dialog.confirm({ title: "성공", message: "학생 정보 파일 업데이트 완료\n학생 정보를 수정해주세요" });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    } finally {
      setRunning(false);
    }
  }

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="mt-1 text-sm text-muted-foreground">
            {meta.guide}
          </p>
        </div>
        <Separator className="mb-4" />
        <div className="w-full flex flex-col justify-center items-center flex-1">
          <Card className="rounded-2xl h-[250px] w-[300px] border-border/80 shadow-sm m-5">
            <CardContent className="flex h-full flex-col justify-center">
              {/* 중앙 영역 */}
              <div className="flex flex-col items-center gap-2 text-center">
                <FileSpreadsheet className="h-8 w-8 text-green-600 mb-2" />
                <div className="text-sm font-medium mb-2">학생 정보 파일 업데이트</div>
              </div>
            </CardContent>
          </Card>
          <Button
            className={`rounded-xl text-white w-[300px] mb-5 transition-colors ${
              done ? "bg-green-600 hover:bg-green-600/90" : "bg-black hover:bg-black/90"
            }`}
            onClick={start}
            disabled={running}
          >
            {running ? (
              <Spinner className="h-4 w-4" />
            ) : done ? (
              <Check className="h-4 w-4" />
            ) : (
              <Play className="h-4 w-4" />
            )}
            <span className="ml-2">{done ? "업데이트 완료" : "업데이트"}</span>
          </Button>
        </div>
      </CardContent>
    </Card>
  )
}
