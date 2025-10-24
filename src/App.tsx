import React, { useEffect, useRef, useState } from "react"
import { Button } from "@/components/ui/button"
import { Separator } from "@/components/ui/separator"
import {
  HelpCircle,
  RefreshCw,
  Settings2,
  Users,
  Database,
  MessageSquare,
  ClipboardList,
  Save,
  FileEdit,
  FlaskConical,
  UserPlus,
  UserMinus,
  Shuffle,
  GraduationCap,
  CalendarDays,
  FolderOpen,
} from "lucide-react"

import { getActionView } from "@/views";
import { descriptions } from "@/meta/descriptions";
import type { OmikronActionKey } from "@/types/omikron";
import FullHeader from "@/components/FullHeader";
import { rpc } from "pyloid-js";
import InitView from "./views/PrereqSetupView";
import useHolidayDialog from "./components/holiday-dialog/useHolidayDialog";
import { useAppDialog } from "./components/app-dialog/AppDialogProvider";

interface Props {
  onAction?: (key: OmikronActionKey) => void
  width?: number
  height?: number
  sidebarPercent?: number
}

// Sidebar button
function NavButton({
  icon: Icon,
  label,
  active,
  onClick,
}: {
  icon: React.ComponentType<{ className?: string }>
  label: string
  active?: boolean
  onClick?: () => void
}) {
  return (
    <Button
      variant={active ? "secondary" : "ghost"}
      className={
        "w-full justify-start gap-3 rounded-xl text-[15px] " +
        (active ? "bg-accent/100 shadow-sm" : "hover:bg-accent/200")
      }
      onClick={onClick}
    >
      <Icon className="h-4 w-4" />
      <span className="truncate">{label}</span>
    </Button>
  )
}

const groups: {
  title: string
  icon: React.ComponentType<{ className?: string }>
  items: { key: OmikronActionKey; label: string; icon: React.ComponentType<{ className?: string }> }[]
}[] = [
  {
    title: "기수 변경 관련",
    icon: Settings2,
    items: [
      { key: "update-class", label: "반 업데이트", icon: RefreshCw },
      { key: "rename-data-file", label: "데이터 파일 이름 변경", icon: FileEdit },
      { key: "update-students", label: "학생 정보 업데이트", icon: Users },
      { key: "update-teacher", label: "담당 선생님 변경", icon: GraduationCap}
    ],
  },
  {
    title: "일일 데이터 저장 및 메시지 전송",
    icon: Database,
    items: [
      { key: "generate-daily-form", label: "데일리 테스트 기록 양식 생성", icon: ClipboardList },
      { key: "save-exam", label: "시험 결과 저장", icon: Save },
      { key: "send-exam-message", label: "시험 결과 메시지 전송", icon: MessageSquare },
      { key: "save-individual-exam", label: "개별 시험 결과 저장", icon: FlaskConical },
      { key: "save-retest", label: "재시험 결과 저장", icon: Save },
    ],
  },
  {
    title: "데이터 관리",
    icon: ClipboardList,
    items: [
      { key: "reapply-conditional-format", label: "데이터 파일 조건부 서식 재지정", icon: RefreshCw },
    ],
  },
  {
    title: "학생 관리",
    icon: Users,
    items: [
      { key: "add-student", label: "신규생 추가", icon: UserPlus },
      { key: "remove-student", label: "퇴원 처리", icon: UserMinus },
      { key: "move-student", label: "학생 반 이동", icon: Shuffle },
    ],
  },
]

// === Main ===
export default function OmikronPanel({ onAction, width = 1400, height = 830, sidebarPercent = 10 }: Props) {
  const dialog = useAppDialog();
  const [selected, setSelected] = useState<OmikronActionKey>("welcome")
  const [mountedKeys, setMountedKeys] = useState<OmikronActionKey[]>(["welcome"]);
  // const View = useMemo(() => getActionView(selected), [selected])
  const [missing, setMissing] = useState(false);
  const { openHolidayDialog } = useHolidayDialog()
  const HELP_URL = "https://omikron-db.notion.site/ad673cca64c146d28adb3deaf8c83a0d?pvs=4"

  // ✅ 프리체크 상태 + 지속 폴링
  const [state, setState] = useState<any>(null);
  const pollRef = useRef<number | null>(null);

  const handleOpenHelp = async () => {
    try {
      const res = await rpc.call("open_url", { url: HELP_URL });
      if (!res?.ok) {
        console.error(res?.error);
      }
    } catch (error) {
      console.error(error);
    }
  };

  const fetchState = async () => {
    try {
      const res = await rpc.call("check_data_files", {});
      setState(res);
      if(!res.ok) setMissing(true);
    } catch {
      // RPC 사용 불가(브라우저 단독 실행 등) 시엔 통과
      setState({ ok: true, has_class: true, has_data: true, has_student: true, missing: [] });
    }
  };

  const changeDataDir = async () => {
    try {
      const res = await rpc.call("change_data_dir", {});
      if (res?.ok) {
        await dialog.confirm({title: "성공", message: "데이터 저장 위치를 변경하였습니다."})
      }
    } catch (e: any) {
      await dialog.error({title: "에러", message: `${e}`})
    } finally {
      fetchState();
    }
  }

  useEffect(() => {
    if(!missing){
      fetchState();
      // 2초마다 지속 추적
      pollRef.current = window.setInterval(fetchState, 2000);
      return () => {
        if (pollRef.current) window.clearInterval(pollRef.current);
      };
    }
  }, [missing]);

  useEffect(() => {
  setMountedKeys(prev => (prev.includes(selected) ? prev : [...prev, selected]));
}, [selected]);

  return (
    <div className="h-screen overflow-hidden bg-gradient-to-b from-point/10 to-transparent">
      <div
        className="mx-auto flex flex-col overflow-hidden rounded-2xl border border-border/80 bg-background shadow-xl"
        style={{ width, height }}
      >
        {/* Header */}
        <div className="flex h-16 items-center justify-between border-b border-border/80 px-6">
          <div className="flex items-center gap-3">
            <div>
              <h1 className="text-lg font-semibold tracking-tight text-foreground py-5">Omikron 데이터 프로그램</h1>
            </div>
          </div>
          <div className="flex gap-2">
            <Button variant="outline" className="rounded-xl" onClick={changeDataDir}>
              <FolderOpen className="h-4 w-4" /> 데이터 저장 위치 변경
            </Button>
            <Button variant="outline" className="rounded-xl" onClick={() => openHolidayDialog()}>
              <CalendarDays className="mr-2 h-4 w-4" /> 학원 휴일 설정
            </Button>
            <Button variant="outline" className="rounded-xl" onClick={handleOpenHelp}>
              <HelpCircle className="mr-2 h-4 w-4" /> 사용법 및 도움말
            </Button>
          </div>
        </div>

        {/* Body: sidebar + page view */}
        <div
          className="grid flex-1"
          style={{ gridTemplateColumns: `minmax(310px, ${sidebarPercent}%) 1fr` }}
        >
          {/* Sidebar */}
          <aside className="border-r border-border/80 bg-card/30">
            <div className="flex h-full flex-col">
              <div className="px-5 pt-4 pb-2 text-sm font-semibold text-muted-foreground">작업 메뉴</div>
              <Separator />
              <div className="flex-1 px-4 py-4">
                <div className="space-y-6">
                  {groups.map((g, gi) => (
                    <div key={gi}>
                      <div className="mb-2 flex items-center gap-2 px-1 text-sm font-semibold text-muted-foreground">
                        <g.icon className="h-4 w-4 text-point" />
                        {g.title}
                      </div>
                      <div className="space-y-1">
                        {g.items.map(({ key, label, icon }) => (
                          <NavButton
                            key={key}
                            icon={icon}
                            label={label}
                            active={selected === key}
                            onClick={() => setSelected(key)}
                          />
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </aside>

          {/* Right pane */}
          <section className="flex flex-col p-3 min-h-0">
            { state && !state.ok ? (
              <InitView state={state} onRefresh={fetchState} />
            ) : (
              <>
                {selected === "welcome" ? null : <FullHeader title={descriptions[selected].title} />}
                <div className="h-full w-full overflow-hidden relative">
                  {mountedKeys.map((key) => {
                    const ViewComp = getActionView(key);
                    const visible = key === selected;
                    return (
                      <div
                        key={key}
                        className={visible ? "block h-full w-full" : "hidden h-0 w-0"}
                      >
                        <ViewComp meta={descriptions[key]} onAction={onAction} />
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </section>
        </div>
      </div>
    </div>
  )
}
