import { useEffect, useRef, useState } from "react"
import type { ViewProps } from "@/types/omikron"
import { Button } from "@/components/ui/button"
import { Card, CardContent } from "@/components/ui/card"
import { Separator } from "@/components/ui/separator"
import { FileSpreadsheet, Play, } from "lucide-react"
import { Spinner } from "@/components/ui/spinner"

// 헤더는 루트 레이아웃에 있으므로 이 뷰에는 포함하지 않음
export default function UpdateStudentView({ meta }: ViewProps) {
  // ===== Single-file Drag & Drop =====
  const [running, setRunning] = useState(false)
  const [generated, setGenerated] = useState(false)

  const pollRef = useRef<number | null>(null)

const start = async () => {
  if (running) return
  setRunning(true)

  setGenerated(true)
}

useEffect(() => {
  return () => {
    if (pollRef.current) clearInterval(pollRef.current)
  }
}, [])

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
            className="rounded-xl bg-black text-white w-[300px] mb-5"
            onClick={start}
            disabled={running || generated}
          >
            {!running ? <Play className="h-4 w-4" /> : <Spinner className="h-4 w-4" />} { generated ? "생성 완료" : "생성"}
          </Button>
        </div>
      </CardContent>
    </Card>
  )
}
