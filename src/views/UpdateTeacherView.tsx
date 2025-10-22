import { useEffect, useState } from "react";
import type { ViewProps } from "@/types/omikron";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { Input } from "@/components/ui/input";
import { School, User, Check, Play } from "lucide-react";
import { rpc } from "pyloid-js";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";
import { Spinner } from "@/components/ui/spinner";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

export default function UpdateTeacherView({ meta }: ViewProps) {
  const dialog = useAppDialog();

  const [loading, setLoading] = useState(false);
  const [infoLoading, setInfoLoading] = useState(false);
  const [running, setRunning] = useState(false);
  const [done, setDone] = useState(false);

  const [classList, setClassList] = useState<string[]>([]);
  const [selectedClass, setSelectedClass] = useState<string>("");
  const [selectedTeacher, setSelectedTeacher] = useState<string>("");
  const [teacherName, setTeacherName] = useState<string>("");

  const canSubmit = !!selectedClass && teacherName.trim().length > 0;

  // 반 목록 로드
  const loadClasses = async () => {
    try {
      setLoading(true);
      const res = await rpc.call("get_class_list", {});
      setClassList(Array.isArray(res) ? (res as string[]) : []);
      if (selectedClass && !res?.includes(selectedClass)) {
        setSelectedClass("");
      }
    } catch (e: any) {
      await dialog.error({title: "에러", message: `${e?.message || e}\n반 목록을 가져오는데 실패했습니다.`});
      setClassList([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
  if (!selectedClass) {
    setSelectedTeacher("");
    return;
  }
  (async () => {
    try {
      setInfoLoading(true);
      const res = await rpc.call("get_class_info", { class_name: selectedClass });
      // 서버가 [exists, teacher, day, time] 형태라고 가정 → 두 번째가 담당 교사명
      setSelectedTeacher(res?.[1] ?? "");
    } catch (e: any) {
      await dialog.error({
        title: "에러",
        message: `${e?.message || e}\n반 정보를 가져오는데 실패했습니다.`,
      });
      setSelectedTeacher("");
    } finally {
      setInfoLoading(false);
    }
  })();
  // eslint-disable-next-line react-hooks/exhaustive-deps
}, [selectedClass]);

  useEffect(() => {
    loadClasses();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleSubmit = async () => {
    if (!canSubmit || running) return;

    const yes = await dialog.warning({
      title: "선생님을 변경할까요?",
      message: `선택 반: ${selectedClass}\n${selectedTeacher} 선생님 → ${teacherName.trim()} 선생님`,
      confirmText: "변경",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      setRunning(true);
      setDone(false);

      // target_class_name, target_teacher_name
      const res = await rpc.call("change_class_info", {
        target_class_name:   selectedClass,
        target_teacher_name: teacherName.trim(),
      });

      if (res?.ok) {
        setDone(true);
        await dialog.confirm({ title: "성공", message: "교사명이 변경되었습니다." });
      } else {
        await dialog.error({ title: "실패", message: res?.error || "변경에 실패했습니다." });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    } finally {
      setRunning(false);
    }
  };

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="mt-1 text-sm text-muted-foreground">{meta.guide}</p>
        </div>
        <Separator className="mb-4" />

        {/* 본문: 좌(반 선택) → 우(새 교사명 입력) */}
        <div className="h-full grid grid-cols-[1fr_auto_1fr] items-center gap-6">
          {/* 좌측: 반 선택 */}
          <Card className="rounded-2xl shadow-sm">
            <CardContent className="flex h-60 flex-col justify-between p-4">
              <div className="flex flex-col items-center gap-2 text-center">
                <School className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">
                  {selectedClass || "반 미선택"}
                </div>
              </div>

              <div className="grid gap-2">
                <Select value={selectedClass} onValueChange={setSelectedClass}>
                  <SelectTrigger className="w-full rounded-xl" disabled={loading}>
                    <SelectValue placeholder={loading ? "불러오는 중…" : "반 선택"} />
                  </SelectTrigger>
                  <SelectContent>
                    {classList.map((name) => (
                      <SelectItem key={name} value={name}>
                        {name}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>

                {/* ✅ 현재 선생님 표시 바 (우측 버튼과 동일한 크기: h-10, w-full, rounded-xl) */}
                <div
                  className="h-9 w-full rounded-xl border bg-muted/50 px-3 flex items-center justify-between"
                  aria-live="polite"
                >
                  <span className="text-sm text-muted-foreground">현재 선생님</span>
                  <span className="text-sm font-medium flex items-center gap-2">
                    {infoLoading ? <Spinner className="h-4 w-4" /> : null}
                    {selectedTeacher || (selectedClass ? "지정되지 않았습니다" : "반을 선택하세요")}
                  </span>
                </div>
              </div>
            </CardContent>
          </Card>


          {/* 가운데 아이콘(장식) */}
          <div className="flex h-60 items-center justify-center">
            {/* 비워두거나 화살표/아이콘을 넣어도 됨 */}
          </div>

          {/* 우측: 새 교사명 + 실행 버튼 */}
          <Card className="rounded-2xl shadow-sm">
            <CardContent className="flex h-60 flex-col justify-between p-4">
              <div className="flex flex-col items-center gap-2 text-center">
                <User className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">변경할 담당 선생님</div>
              </div>

              <div className="grid gap-2">
                <Input
                  className="rounded-xl"
                  value={teacherName}
                  onChange={(e) => setTeacherName(e.target.value)}
                  placeholder="교사명 입력"
                  disabled={loading}
                />

                {/* 실행 버튼 */}
                <Button
                  className={`w-full rounded-xl text-white transition-colors ${
                    done ? "bg-green-600 hover:bg-green-600/90" : "bg-black hover:bg-black/90"
                  }`}
                  onClick={handleSubmit}
                  disabled={!canSubmit || running}
                  title={
                    !selectedClass
                      ? "반을 먼저 선택하세요"
                      : teacherName.trim().length === 0
                      ? "교사명을 입력하세요"
                      : undefined
                  }
                >
                  {running ? (
                    <Spinner className="h-4 w-4" />
                  ) : done ? (
                    <Check className="h-4 w-4" />
                  ) : (
                    <Play className="h-4 w-4" />
                  )}
                  <span className="ml-2">{done ? "변경 완료" : "변경"}</span>
                </Button>
              </div>
            </CardContent>
          </Card>
        </div>

        {/* 우하단: 새로고침 */}
        <div className="mt-auto flex items-center justify-end">
          <Button
            className="rounded-xl"
            variant="outline"
            onClick={loadClasses}
            disabled={loading}
            title="반 목록을 다시 불러옵니다."
          >
            {loading ? "불러오는 중…" : "새로고침"}
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}
