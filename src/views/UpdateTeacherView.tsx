import { useEffect, useMemo, useState } from "react";
import type { ViewProps } from "@/types/omikron";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { Input } from "@/components/ui/input";
import { Check, Play } from "lucide-react";
import { rpc } from "pyloid-js";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";
import { Spinner } from "@/components/ui/spinner";
import { ScrollArea } from "@/components/ui/scroll-area";

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
  const [classQuery, setClassQuery] = useState<string>("");

  const canSubmit = !!selectedClass && teacherName.trim().length > 0;

  const loadClasses = async () => {
    try {
      setLoading(true);
      const res = await rpc.call("get_class_list", {});
      if (res?.ok) {
        const next = Array.isArray(res.data) ? (res.data as string[]) : [];
        setClassList(next);
        if (selectedClass && !next.includes(selectedClass)) {
          setSelectedClass("");
          setSelectedTeacher("");
        }
      } else {
        await dialog.error({ title: "반 정보 파일 데이터 수집 실패", message: res?.error || "" });
      }
    } catch (e: any) {
      await dialog.error({
        title: "에러",
        message: `반 목록 불러오기 실패: ${e?.message || e}`,
      });
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
        setSelectedTeacher(res?.[1] ?? "");
      } catch (e: any) {
        await dialog.error({
          title: "에러",
          message: `반 정보 불러오기 실패: ${e?.message || e}`,
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
      const res = await rpc.call("change_class_info", {
        target_class_name: selectedClass,
        target_teacher_name: teacherName.trim(),
      });

      if (res?.ok) {
        await dialog.confirm({ title: "성공", message: "담당 선생님이 변경되었습니다." });
        setDone(true);
      } else {
        await dialog.error({ title: "담당 선생님 변경 실패", message: res?.error || "" });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    } finally {
      setRunning(false);
      setTimeout(async () => {
        await handleRefresh();
      }, 5000);
    }
  };

  const handleRefresh = async () => {
    setDone(false);
    setSelectedClass("");
    setSelectedTeacher("");
    setTeacherName("");
    setClassQuery("");
    loadClasses();
  };

  const filteredClasses = useMemo(() => {
    const q = classQuery.trim().toLowerCase();
    if (!q) return classList;
    return classList.filter((name) => name.toLowerCase().includes(q));
  }, [classList, classQuery]);

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="mt-1 text-sm text-muted-foreground">{meta.guide}</p>
        </div>
        <Separator className="mb-4" />

        <div className="grid flex grid-cols-1 gap-4 lg:grid-cols-[1.2fr_1fr]">
          {/* Left: search + list */}
          <Card className="flex h-full flex-col rounded-2xl shadow-sm">
            <CardHeader className="space-y-0 pb-0 pt-2 gap-0 py-0">
              <CardTitle className="text-base font-semibold">반 목록</CardTitle>
            </CardHeader>
            <CardContent className="flex min-h-0 flex-1 flex-col gap-2">
              <Input
                type="search"
                value={classQuery}
                onChange={(e) => setClassQuery(e.target.value)}
                placeholder="반 검색"
                className="h-9 w-full rounded-xl"
                disabled={loading}
              />
              <div className="relative min-h-0 flex-1 rounded-lg border">
                {loading && (
                  <div className="absolute inset-0 z-10 flex items-center justify-center bg-background/60">
                    <div className="flex items-center gap-2 text-sm text-muted-foreground">
                      <Spinner />
                      불러오는 중...
                    </div>
                  </div>
                )}
                <ScrollArea className="h-[375px] w-full p-2">
                  {filteredClasses.length === 0 ? (
                    <div className="p-2 text-xs text-muted-foreground">반이 없습니다</div>
                  ) : (
                    <ul className="space-y-1">
                      {filteredClasses.map((name) => {
                        const isSel = selectedClass === name;
                        return (
                          <li key={name}>
                            <button
                              type="button"
                              onClick={() => setSelectedClass(name)}
                              disabled={loading}
                              className={`group flex w-full items-start gap-2 rounded-md border px-2 py-1 text-left text-xs transition ${
                                isSel ? "bg-blue-50 border-blue-200" : "hover:bg-accent border-transparent"
                              } ${loading ? "opacity-60 cursor-not-allowed" : ""}`}
                            >
                              <span className="flex-1 min-w-0 break-all">{name}</span>
                            </button>
                          </li>
                        );
                      })}
                    </ul>
                  )}
                </ScrollArea>
              </div>
            </CardContent>
          </Card>

          {/* Right: details */}
          <Card className="flex h-full flex-col rounded-2xl shadow-sm">
            <CardHeader className="space-y-0 pb-0 pt-2">
              <CardTitle className="text-base font-semibold">선택된 반</CardTitle>
            </CardHeader>
            <CardContent className="flex min-h-0 flex-1 flex-col justify-between">
              <div className="space-y-3">
                <div className="rounded-lg border p-3 text-sm">
                  <div className="text-xs text-muted-foreground">반명</div>
                  <div className="font-medium">{selectedClass || "-"}</div>
                </div>
                <div className="rounded-lg border p-3 text-sm">
                  <div className="text-xs text-muted-foreground">현재 선생님</div>
                  <div className="flex items-center gap-2 font-medium">
                    {infoLoading ? <Spinner className="h-4 w-4" /> : null}
                    {selectedClass ? (selectedTeacher || "지정되지 않음") : "-"}
                  </div>
                </div>
                <div className="rounded-lg border p-3 text-sm">
                  <div className="text-xs text-muted-foreground">변경할 선생님</div>
                  <Input
                    className="mt-2 rounded-lg"
                    value={teacherName}
                    onChange={(e) => setTeacherName(e.target.value)}
                    placeholder="선생님 이름 입력"
                    disabled={loading}
                  />
                </div>
              </div>

              <div className="mt-4 grid gap-2">
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

        <div className="mt-auto pt-4 flex items-center justify-end">
          <Button
            className="rounded-xl"
            variant="outline"
            onClick={handleRefresh}
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
