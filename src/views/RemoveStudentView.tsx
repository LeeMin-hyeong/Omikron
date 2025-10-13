// src/views/AddStudentView.tsx
import { useEffect, useState } from "react";
import type { ViewProps } from "@/types/omikron";
import { rpc } from "pyloid-js";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";

import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import {
  Select, SelectContent, SelectItem, SelectTrigger, SelectValue,
} from "@/components/ui/select";
import { School, User2, UserMinus, UserX } from "lucide-react";
import { Input } from "@/components/ui/input";
import { Spinner } from "@/components/ui/spinner";

export default function RemoveStudentView({ onAction }: ViewProps) {
  const dialog = useAppDialog();

  const [classes, setClasses] = useState<Array<{ id?: string; name: string }>>([]);
  const [fromClass, setFromClass] = useState<string>("");
  const [student, setStudent] = useState<string>("");
  const [loading, setLoading] = useState(false);
  const [running, setRunning] = useState(false);
  const [students, setStudents] = useState<Array<{ id: string; name: string }>>([]);

  const loadClasses = async () => {
    try {
      setLoading(true);
      const res = await rpc.call("list_classes", {}); // [{id?, name}]
      const list = Array.isArray(res) ? res : [];
      setClasses(list);
      // 선택된 반이 목록에 없으면 초기화
      if (fromClass && !list.some((c) => (c.id ?? c.name) === fromClass)) {
        setFromClass("");
      }
    } catch {
      setClasses([]);
    } finally {
      setLoading(false);
    }
  };

  // 반 선택 시: 해당 반의 학생 목록 로드
  useEffect(() => {
    setStudents([]);
    setStudent("");
    if (!fromClass) return;
    (async () => {
      try {
        const res = await rpc.call("list_students_by_class", { class_name: fromClass }); // [{id,name}]
        setStudents(Array.isArray(res) ? res : []);
      } catch {
        setStudents([]);
      }
    })();
  }, [fromClass]);

  useEffect(() => {
    loadClasses();
  }, []);

  const canSubmit = Boolean(fromClass && student.trim());

  const handleSubmit = async () => {
    if (!canSubmit) return;

    const yes = await dialog.warning({
      title: "학생을 퇴원 처리할까요?",
      message: `‘${fromClass}’반 ${student.trim()} 학생을\n 퇴원 처리합니다.`,
      confirmText: "퇴원 처리",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      setRunning(true);
      onAction?.("remove-student");
      const res = await rpc.call("remove_student_class", {
        student_name: student.trim(),
        from_class: fromClass,
      }); // 서버: {ok:true}
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: "학생이 반에서 퇴원 처리되었습니다." });
        setStudent("");
      } else {
        await dialog.error({ title: "실패", message: res?.error || "퇴원 처리에 실패했습니다." });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    } finally {
      setRunning(false);
    }
  };

  const handleRefresh = async () => {
    await loadClasses();
    await dialog.confirm({ title: "새로고침", message: "반 목록을 다시 불러왔습니다." });
  };

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="mt-1 text-sm text-muted-foreground">
            학생을 퇴원 처리합니다.
          </p>
        </div>
        <Separator className="mb-4" />

        {/* ✅ 중앙 정렬 컨테이너 */}
        <div className="flex flex-1 items-center justify-center">
          <Card className="h-80 w-[420px] rounded-2xl shadow-sm">
            <CardContent className="flex h-full flex-col justify-between p-4">
              {/* 중앙: 아이콘 + 타이틀 */}
              <div className="flex flex-1 flex-col items-center justify-center gap-2 text-center">
                <UserMinus className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">퇴원 학생 선택</div>
              </div>

              {/* 하단 컨트롤 */}
              <div className="grid gap-2">
                <Select value={fromClass} onValueChange={(v) => setFromClass(v)}>
                  <SelectTrigger className="w-full rounded-xl" disabled={loading}>
                    <SelectValue placeholder={loading ? "불러오는 중…" : "반 선택"} />
                  </SelectTrigger>
                  <SelectContent>
                    {classes.map((c) => {
                      const value = c.id ?? c.name;
                      return (
                        <SelectItem key={value} value={value}>
                          {c.name}
                        </SelectItem>
                      );
                    })}
                  </SelectContent>
                </Select>

                <Select
                  value={student}
                  onValueChange={(v) => setStudent(v)}
                  disabled={!fromClass || students.length === 0}
                >
                  <SelectTrigger className="w-full rounded-xl">
                    <SelectValue placeholder="학생 선택" />
                  </SelectTrigger>
                  <SelectContent>
                    {students.map((s) => (
                      <SelectItem key={s.id} value={s.id}>
                        {s.name}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>

                <div className="mt-2 flex items-center gap-2">
                  <Button
                    className="w-full rounded-xl bg-red-600 text-white"
                    disabled={!canSubmit}
                    onClick={handleSubmit}
                  >
                    {running ? <Spinner /> : "퇴원 처리"}
                  </Button>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>

        {/* 하단: 새로고침 */}
        <div className="mt-auto flex items-center justify-end">
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
