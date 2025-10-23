import { useState } from "react"
import type { ViewProps } from "@/types/omikron"
import { Button } from "@/components/ui/button"
import { Card, CardContent } from "@/components/ui/card"
import { Separator } from "@/components/ui/separator"
import { Check, FileSpreadsheet, Play, } from "lucide-react"
import { Spinner } from "@/components/ui/spinner"
import { rpc } from "pyloid-js"
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider"

export default function GenerateDailyFormView({ meta }: ViewProps) {
  const dialog = useAppDialog();

  const [running, setRunning] = useState(false)
  const [generated, setGenerated] = useState(false)

  const start = async () => {
    if (running) return
    setGenerated(false);

    try {
      setRunning(true);
      const res = await rpc.call("make_data_form", {});
      if(res?.ok){
        await dialog.confirm({ title: "성공", message: "데일리테스트 기록 양식을 생성하였습니다." });
        setGenerated(true);
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
                <div className="text-sm font-medium mb-2">데일리 테스트 기록 양식</div>
              </div>
            </CardContent>
          </Card>
          <Button
            className={`rounded-xl text-white w-[300px] mb-5 transition-colors ${
              generated ? "bg-green-600 hover:bg-green-600/90" : "bg-black hover:bg-black/90"
            }`}
            onClick={start}
            disabled={running}
          >
            {running ? (
              <Spinner className="h-4 w-4" />
            ) : generated ? (
              <Check className="h-4 w-4" />
            ) : (
              <Play className="h-4 w-4" />
            )}
            <span className="ml-2">{generated ? "업데이트 완료" : "생성"}</span>
          </Button>
        </div>
      </CardContent>
    </Card>
  )
}
