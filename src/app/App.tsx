import React, { useEffect, useRef, useState } from "react";
import { Button } from "@/shared/components/ui/button";
import { Separator } from "@/shared/components/ui/separator";
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
  GraduationCap,
  CalendarDays,
  FolderOpen,
  FolderSync,
  CalendarCheck,
  Settings,
} from "lucide-react";

import { getActionView } from "@/app/viewRegistry";
import { descriptions } from "@/shared/meta/descriptions";
import type { tdmActionKey } from "@/shared/types/tdm";
import FullHeader from "@/shared/components/FullHeader";
import { rpc } from "pyloid-js";
import PrereqSetupView from "@/features/configuration/PrereqSetupView";
import useHolidayDialog from "@/shared/components/dialogs/holiday/useHolidayDialog";
import { useAppDialog } from "@/shared/components/dialogs/app/AppDialogProvider";
import InitialConfigView from "@/features/configuration/InitialConfigView";
import TermsAgreementDialog from "@/shared/components/dialogs/terms/TermsAgreementDialog";
import NoticeDialog from "@/shared/components/dialogs/notice/NoticeDialog";

interface Props {
  onAction?: (key: tdmActionKey) => void;
  width?: number;
  height?: number;
  sidebarPercent?: number;
}

interface SetupConfig {
  url: string;
  dataDir: string;
  dataFileName: string;
  dailyTest: string;
  makeupTest: string;
  makeupTestDate: string;
}

type NoticePopup = {
  open: boolean;
  title: string;
  message: string;
  noticeId: string;
};

function NavButton({
  icon: Icon,
  label,
  active,
  onClick,
}: {
  icon: React.ComponentType<{ className?: string }>;
  label: string;
  active?: boolean;
  onClick?: () => void;
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
  );
}

const groups: {
  title: string;
  icon: React.ComponentType<{ className?: string }>;
  items: {
    key: tdmActionKey;
    label: string;
    icon: React.ComponentType<{ className?: string }>;
  }[];
}[] = [
  {
    title: "기수 변경 관리",
    icon: Settings2,
    items: [
      { key: "update-class", label: "반 업데이트", icon: RefreshCw },
      { key: "rename-data-file", label: "데이터파일 이름 변경", icon: FileEdit },
      { key: "update-students", label: "학생 정보 업데이트", icon: Users },
      { key: "update-teacher", label: "담당 선생님 변경", icon: GraduationCap },
    ],
  },
  {
    title: "일일 데이터 저장 및 메시지 전송",
    icon: Database,
    items: [
      { key: "generate-daily-form", label: "일일테스트 기록 양식 생성", icon: ClipboardList },
      { key: "save-exam", label: "시험 결과 저장", icon: Save },
      { key: "send-exam-message", label: "시험 결과 메시지 전송", icon: MessageSquare },
      { key: "save-individual-exam", label: "개별 시험 결과 저장", icon: Save },
      { key: "save-retest", label: "재시험 결과 저장", icon: Save },
    ],
  },
  {
    title: "데이터 관리",
    icon: ClipboardList,
    items: [
      { key: "reapply-conditional-format", label: "데이터파일 조건부 서식 재적용", icon: RefreshCw },
    ],
  },
  {
    title: "학생 관리",
    icon: Users,
    items: [{ key: "manage-student", label: "학생 관리", icon: Users }],
  },
  {
    title: "설정",
    icon: Settings,
    items: [{ key: "edit-message-config", label: "아이소식 설정", icon: MessageSquare },]
  }
];

export default function TdmPanel({
  onAction,
  width = 1400,
  height = 830,
  sidebarPercent = 10,
}: Props) {
  const [viewport, setViewport] = useState(() => ({
    width: typeof window !== "undefined" ? window.innerWidth : width,
    height: typeof window !== "undefined" ? window.innerHeight : height,
  }));
  const dialog = useAppDialog();
  const [selected, setSelected] = useState<tdmActionKey>("welcome");
  const [mountedKeys, setMountedKeys] = useState<tdmActionKey[]>(["welcome"]);
  const [missing, setMissing] = useState(false);
  const [holidayChecked, setHolidayChecked] = useState(false);
  const [state, setState] = useState<any>(null);
  const [bootChecked, setBootChecked] = useState(false);
  const [needsInitialConfig, setNeedsInitialConfig] = useState(false);
  const [needsTermsAgreement, setNeedsTermsAgreement] = useState(false);
  const [noticePopup, setNoticePopup] = useState<NoticePopup>({
    open: false,
    title: "",
    message: "",
    noticeId: "",
  });
  const [setupConfig, setSetupConfig] = useState<SetupConfig>({
    url: "",
    dataDir: "",
    dataFileName: "",
    dailyTest: "",
    makeupTest: "",
    makeupTestDate: "",
  });
  const pollRef = useRef<number | null>(null);
  const dataDirPromptingRef = useRef(false);
  const noticeCheckedRef = useRef(false);
  const { openHolidayDialog, lastHolidaySelection } = useHolidayDialog();
  const HELP_URL = "https://tdm-db.notion.site/instruction?source=copy_link";

  const showStartupNotice = async () => {
    if (noticeCheckedRef.current) return;
    noticeCheckedRef.current = true;

    try {
      const res = await rpc.call("get_startup_notice", {});
      if (!res?.ok || !res?.enabled) return;
      setNoticePopup({
        open: true,
        title: res.title ?? "공지사항",
        message: res.message ?? "",
        noticeId: res.noticeId ?? "",
      });
    } catch {
      // ignore in browser-only mode
    }
  };

  const closeNoticePopup = () => {
    setNoticePopup((prev) => ({ ...prev, open: false }));
  };

  const fetchState = async () => {
    try {
      const res = await rpc.call("check_data_files", {});
      setState(res);
      setMissing(!res.ok);
      if (res?.data_dir_valid === false && !dataDirPromptingRef.current) {
        dataDirPromptingRef.current = true;
        await dialog.error({
          title: "데이터 저장 위치가 유효하지 않습니다",
          message: "데이터 저장 위치를 변경해 주세요.",
          confirmText: "변경",
        });
        await changeDataDir();
        dataDirPromptingRef.current = false;
      }
    } catch {
      setState({ ok: true, has_class: true, has_data: true, has_student: true, missing: [] });
    }
  };

  const checkBootstrap = async () => {
    try {
      const res = await rpc.call("get_config_status", {});
      if (!res?.ok) {
        setNeedsInitialConfig(true);
        setNeedsTermsAgreement(false);
        return;
      }

      // If config.json exists, skip initial setup even when some values are invalid
      // (e.g. dataDir). Those cases are handled by the separate reconfiguration flow.
      setNeedsInitialConfig(!res.exists);
      setNeedsTermsAgreement(Boolean(res.exists && !res.termsAccepted));
      setSetupConfig({
        url: res?.config?.url ?? "",
        dataDir: res?.config?.dataDir ?? "",
        dataFileName: res?.config?.dataFileName ?? "",
        dailyTest: res?.config?.dailyTest ?? "",
        makeupTest: res?.config?.makeupTest ?? "",
        makeupTestDate: res?.config?.makeupTestDate ?? "",
      });

      if (res.exists) {
        await fetchState();
        if (res.termsAccepted) {
          await showStartupNotice();
        }
      }
    } catch {
      // browser-only mode
    } finally {
      setBootChecked(true);
    }
  };

  const handleInitialSetupComplete = async () => {
    await checkBootstrap();
  };

  const handleTermsAccepted = async () => {
    setNeedsTermsAgreement(false);
    await showStartupNotice();
  };

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

  const changeDataDir = async () => {
    try {
      const res = await rpc.call("change_data_dir", {});
      if (res?.ok) {
        await dialog.confirm({ title: "성공", message: "데이터 저장 위치를 변경했습니다." });
      } else if (res?.error) {
        await dialog.error({
          title: "데이터 저장 위치 변경 실패",
          message: res?.error,
          detail: res?.detail,
        });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: `${e}` });
    } finally {
      await fetchState();
    }
  };

  useEffect(() => {
    void checkBootstrap();
  }, []);

  useEffect(() => {
    const onResize = () => {
      setViewport({ width: window.innerWidth, height: window.innerHeight });
    };
    onResize();
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  useEffect(() => {
    if (!bootChecked || missing || needsInitialConfig || needsTermsAgreement) return;
    pollRef.current = window.setInterval(() => {
      void fetchState();
    }, 2000);
    return () => {
      if (pollRef.current) window.clearInterval(pollRef.current);
    };
  }, [bootChecked, missing, needsInitialConfig, needsTermsAgreement]);

  useEffect(() => {
    if (lastHolidaySelection) setHolidayChecked(true);
  }, [lastHolidaySelection]);

  useEffect(() => {
    setMountedKeys((prev) => (prev.includes(selected) ? prev : [...prev, selected]));
  }, [selected]);

  if (!bootChecked) {
    return <div className="h-screen w-screen bg-gradient-to-b from-point/10 to-transparent" />;
  }

  const panelScale = Math.min(viewport.width / width, viewport.height / height, 1);
  const scaledWidth = Math.max(1, Math.round(width * panelScale));
  const scaledHeight = Math.max(1, Math.round(height * panelScale));

  return (
    <div className="h-screen w-screen overflow-hidden bg-gradient-to-b from-point/10 to-transparent">
      <TermsAgreementDialog
        open={!needsInitialConfig && needsTermsAgreement}
        onAccepted={handleTermsAccepted}
      />
      <NoticeDialog
        open={noticePopup.open}
        title={noticePopup.title}
        message={noticePopup.message}
        noticeId={noticePopup.noticeId}
        onClose={closeNoticePopup}
      />

      <div className="flex h-full w-full items-center justify-center">
        <div className="relative" style={{ width: scaledWidth, height: scaledHeight }}>
          <div
            className="flex flex-col overflow-hidden rounded-2xl border border-border/80 bg-background shadow-xl"
            style={{
              width,
              height,
              transform: `scale(${panelScale})`,
              transformOrigin: "top left",
            }}
          >
        <div className="flex h-16 items-center justify-between border-b border-border/80 px-6">
          <div className="flex items-center gap-3">
            <h1 className="py-5 text-lg font-semibold tracking-tight text-foreground">
              테스트 데이터 관리 프로그램
            </h1>
          </div>
          {!needsInitialConfig ? (
            <div className="flex gap-2">
              {state && !state.ok ? null : (
                <div className="flex gap-2">
                  <Button
                    variant="outline"
                    className="rounded-xl"
                    onClick={() => rpc.call("open_path", { path: state.data_dir })}
                  >
                    <FolderOpen className="h-4 w-4" /> 데이터 폴더
                  </Button>
                  <Button variant="outline" className="rounded-xl" onClick={changeDataDir}>
                    <FolderSync className="h-4 w-4" /> 저장 위치 변경
                  </Button>
                </div>
              )}
              <Button
                variant="outline"
                className="justify-between rounded-xl"
                onClick={() => openHolidayDialog()}
              >
                {holidayChecked ? (
                  <>
                    <CalendarCheck className="h-4 w-4" />
                    휴일 설정됨
                  </>
                ) : (
                  <>
                    <CalendarDays className="mr-5 h-4 w-4" />
                    휴일 설정
                  </>
                )}
              </Button>
              <Button variant="outline" className="rounded-xl" onClick={handleOpenHelp}>
                <HelpCircle className="h-4 w-4" /> 사용법 및 도움말
              </Button>
            </div>
          ) : null}
        </div>

        {needsInitialConfig ? (
          <section className="flex min-h-0 flex-1 flex-col p-3">
            <InitialConfigView initial={setupConfig} onComplete={handleInitialSetupComplete} />
          </section>
        ) : (
          <div className="grid flex-1" style={{ gridTemplateColumns: `minmax(310px, ${sidebarPercent}%) 1fr` }}>
            <aside className="border-r border-border/80 bg-card/30">
              <div className="flex h-full flex-col">
                <div className="px-5 pb-2 pt-4 text-sm font-semibold text-muted-foreground">작업 메뉴</div>
                <Separator />
                <div className="flex-1 px-4 py-4">
                  <div className="space-y-4">
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

            <section className="flex min-h-0 flex-col p-3">
              {state && !state.ok ? (
                <PrereqSetupView state={state} onRefresh={fetchState} />
              ) : (
                <>
                  {selected === "welcome" ? null : <FullHeader title={descriptions[selected].title} />}
                  <div className="relative h-full w-full overflow-hidden">
                    {mountedKeys.map((key) => {
                      const ViewComp = getActionView(key);
                      const visible = key === selected;
                      return (
                        <div key={key} className={visible ? "block h-full w-full" : "hidden h-0 w-0"}>
                          <ViewComp meta={descriptions[key]} onAction={onAction} />
                        </div>
                      );
                    })}
                  </div>
                </>
              )}
            </section>
          </div>
        )}
          </div>
        </div>
      </div>
    </div>
  );
}
