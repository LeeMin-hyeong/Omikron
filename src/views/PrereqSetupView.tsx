// src/views/PrereqSetupView.tsx
import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { FileSpreadsheet, ChevronsRight, FolderOpen, } from "lucide-react";
import { rpc } from "pyloid-js";
import { Spinner } from "@/components/ui/spinner";

type State = {
  has_class: boolean;
  has_data: boolean;
  has_student: boolean;
  data_file_name?: string;
  cwd: string;
  data_dir: string;
  missing: string[];
};

export default function InitView({
  state,
  onRefresh,
}: {
  state: State;
  onRefresh: () => void;
}) {
  const canInstallDataAndStudent = state.has_class;

  // 각 작업 실행 중 상태
  const [runClass, setRunClass] = useState(false);
  const [runData, setRunData] = useState(false);
  const [runStudent, setRunStudent] = useState(false);

  // 서버 호출 핸들러
  const installClass = async () => {
    if (runClass) return;
    try {
      setRunClass(true);
      const res = await rpc.call("make_class_info", {});
      if (res?.ok) {
        onRefresh();
      } else {
        alert(res?.error || "반 정보 생성에 실패했습니다.");
      }
    } catch (e: any) {
      alert(String(e?.message || e));
    } finally {
      setRunClass(false);
    }
  };

  const installData = async () => {
    if (runData) return;
    try {
      setRunData(true);
      const res = await rpc.call("make_data_file", {});
      if (res?.ok) {
        onRefresh();
      } else {
        alert(res?.error || "데이터 파일 생성에 실패했습니다.");
      }
    } catch (e: any) {
      alert(String(e?.message || e));
    } finally {
      setRunData(false);
    }
  };

  const installStudent = async () => {
    if (runStudent) return;
    try {
      setRunStudent(true);
      const res = await rpc.call("make_student_info", {});
      if (res?.ok) {
        onRefresh();
      } else {
        alert(res?.error || "학생 정보 생성에 실패했습니다.");
      }
    } catch (e: any) {
      alert(String(e?.message || e));
    } finally {
      setRunStudent(false);
    }
  };

  // 공통 타일
  const Tile = ({
    title,
    present,
    onInstall,
    disabled,
    running,
  }: {
    title: string;
    present: boolean;
    onInstall: () => void;
    disabled?: boolean;
    running?: boolean;
  }) => {
    const label = present ? "생성 완료" : running ? "생성 중…" : "생성";
    return (
      <Card className="rounded-2xl border-border/80 shadow-sm">
        <CardContent className="flex h-44 flex-col justify-between pt-4">
          <div className="flex flex-col items-center gap-2 text-center">
            <FileSpreadsheet className={`h-8 w-8 ${present ? "text-green-600" : "text-muted-foreground"}`} />
            <div className="text-sm font-medium">{title}</div>
          </div>
          <div className="grid gap-1">
            <Button
              className="w-full rounded-xl"
              disabled={present || !!disabled || !!running}
              onClick={onInstall}
            >
              {running && <Spinner />}
              {label}
            </Button>
          </div>
        </CardContent>
      </Card>
    );
  };

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col p-4">
        <div className="mb-3">
          <h3 className="text-base font-semibold">필수 파일 생성</h3>
          <p className="mt-1 text-sm text-muted-foreground">
            반 정보 → 데이터 파일 → 학생 정보 순서로 생성합니다.
            {state.data_file_name ? (
              <span className="ml-2">
                (저장된 데이터 파일 이름: <span className="font-medium">{state.data_file_name}</span>)
              </span>
            ) : (
              <span className="ml-2 text-amber-600">
                config.json의 <b>dataFileName</b> 설정이 필요합니다.
              </span>
            )}
          </p>
        </div>
        <Separator className="mb-4" />

        <div className="h-full grid grid-cols-[1fr_auto_1fr_auto_1fr] items-center gap-3">
          <Tile
            title="반 정보.xlsx"
            present={state.has_class}
            onInstall={installClass}
            running={runClass}
          />
          <ChevronsRight className="mx-1 h-5 w-5 text-muted-foreground" />
          <Tile
            title="데이터 파일.xlsx"
            present={state.has_data}
            onInstall={installData}
            disabled={!canInstallDataAndStudent}
            running={runData}
          />
          <ChevronsRight className="mx-1 h-5 w-5 text-muted-foreground" />
          <Tile
            title="학생 정보.xlsx"
            present={state.has_student}
            onInstall={installStudent}
            disabled={!canInstallDataAndStudent}
            running={runStudent}
          />
        </div>

        <div className="mt-auto flex items-center justify-between">
          <div className="text-xs text-muted-foreground">
            부족: {state.missing.length === 0 ? "없음" : state.missing.join(", ")}
          </div>
          <div className="flex items-center gap-2">
            <Button variant="outline" className="rounded-xl" onClick={() => rpc.call("open_path", { path: state.cwd })}>
              <FolderOpen className="h-4 w-4" /> 프로그램 폴더
            </Button>
            <Button variant="outline" className="rounded-xl" onClick={() => rpc.call("open_path", { path: state.data_dir })}>
              <FolderOpen className="h-4 w-4" /> data 폴더
            </Button>
            <Button className="rounded-xl bg-black text-white" onClick={onRefresh}>
              다시 확인
            </Button>
          </div>
        </div>
      </CardContent>
    </Card>
  );
}
