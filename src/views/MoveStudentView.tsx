// src/views/MoveStudentView.tsx
import { useEffect, useMemo, useState } from "react";
import type { ViewProps } from "@/types/omikron";
import { rpc } from "pyloid-js";
import { usePrereq } from "@/contexts/prereq";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";

import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import {
  Select, SelectContent, SelectItem, SelectTrigger, SelectValue,
} from "@/components/ui/select";
import { UserCircle2, ChevronsRight, School, User } from "lucide-react";

type ClassInfo = { id: string; name: string };
type StudentInfo = { id: string; name: string };

export default function MoveStudentView({ onAction }: ViewProps) {
  const { enforcePrereq } = usePrereq();
  const dialog = useAppDialog();

  // 데이터
  const [classes, setClasses] = useState<ClassInfo[]>([]);
  const [fromClass, setFromClass] = useState<string>("");
  const [toClass, setToClass] = useState<string>("");
  const [students, setStudents] = useState<StudentInfo[]>([]);
  const [studentId, setStudentId] = useState<string>("");

  const selectedStudent = useMemo(
    () => students.find((s) => s.id === studentId)?.name ?? "",
    [students, studentId]
  );

  // 최초: 반 목록 로드
  useEffect(() => {
    (async () => {
      try {
        const res = await rpc.call("list_classes", {}); // 서버에서 [{id,name}] 반환하도록 구현
        setClasses(Array.isArray(res) ? res : []);
      } catch {
        setClasses([]);
      }
    })();
  }, []);

  // 반 선택 시: 해당 반의 학생 목록 로드
  useEffect(() => {
    setStudents([]);
    setStudentId("");
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

  const canSubmit = fromClass && toClass && studentId && fromClass !== toClass;

  const handleSubmit = async () => {
    if (!canSubmit) return;

    // 실행 전 파일 프리체크
    const ok = await enforcePrereq();
    if (!ok) return;

    const studentName = selectedStudent;
    const fromName = classes.find((c) => c.id === fromClass || c.name === fromClass)?.name ?? fromClass;
    const toName = classes.find((c) => c.id === toClass || c.name === toClass)?.name ?? toClass;

    const yes = await dialog.warning({
      title: "학생 반을 변경할까요?",
      message: `${studentName} 학생을\n‘${fromName}’ → ‘${toName}’ 로 이동합니다.`,
      confirmText: "변경",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      onAction?.("move-student");
      const res = await rpc.call("move_student_class", {
        student_id: studentId,
        from_class: fromClass,
        to_class: toClass,
      }); // 서버에서 {ok:true} 반환하도록 구현
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: "학생 반이 변경되었습니다." });
        // 초기화
        setToClass("");
      } else {
        await dialog.error({ title: "실패", message: res?.error || "변경에 실패했습니다." });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    }
  };

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="mt-1 text-sm text-muted-foreground">
            학생의 반을 변경합니다. 아이소식에서의 작업이 선행되어야 합니다.
          </p>
        </div>
        <Separator className="mb-4" />

        {/* 본문: 좌(학생 선택) → 우(이동 반) */}
        <div className="h-full grid grid-cols-[1fr_auto_1fr] items-center gap-6">
          {/* 좌측 타일 */}
          <Card className="rounded-2xl shadow-sm">
            <CardContent className="flex h-60 flex-col justify-between p-4">
              {/* 중앙: 아이콘 + 학생명 */}
              <div className="flex flex-col items-center gap-2 text-center">
                <User className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">
                  {selectedStudent || "학생 미선택"}
                </div>
              </div>

              {/* 하단: 반 선택 + 학생 선택 */}
              <div className="grid gap-2">
                <Select
                  value={fromClass}
                  onValueChange={(v) => setFromClass(v)}
                >
                  <SelectTrigger className="w-full rounded-xl">
                    <SelectValue placeholder="반 선택" />
                  </SelectTrigger>
                  <SelectContent>
                    {classes.map((c) => (
                      <SelectItem key={c.id || c.name} value={c.id || c.name}>
                        {c.name}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>

                <Select
                  value={studentId}
                  onValueChange={(v) => setStudentId(v)}
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
              </div>
            </CardContent>
          </Card>

          {/* 화살표 */}
          <div className="flex h-60 items-center justify-center">
            <ChevronsRight className="h-6 w-6 text-muted-foreground" />
          </div>

          {/* 우측 타일 (이동할 반) */}
          <Card className="rounded-2xl shadow-sm">
            <CardContent className="flex h-60 flex-col justify-between p-4">
              {/* 중앙: 아이콘 + 타이틀 */}
              <div className="flex flex-col items-center gap-2 text-center">
                <School className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">이동할 반</div>
              </div>

              {/* 하단: 반 선택 */}
              <div className="grid gap-2">
                <Select
                  value={toClass}
                  onValueChange={(v) => setToClass(v)}
                  disabled={!fromClass || !studentId}
                >
                  <SelectTrigger className="w-full rounded-xl">
                    <SelectValue placeholder="반 선택" />
                  </SelectTrigger>
                  <SelectContent>
                    {classes.map((c) => (
                      <SelectItem
                        key={`to-${c.id || c.name}`}
                        value={c.id || c.name}
                        disabled={c.id === fromClass || c.name === fromClass}
                      >
                        {c.name}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </CardContent>
          </Card>
        </div>

        {/* 하단 실행 버튼 */}
        <div className="mt-auto flex items-center justify-end">
          <Button
            className="rounded-xl bg-black text-white"
            disabled={!canSubmit}
            onClick={handleSubmit}
            title={
              !fromClass
                ? "반을 먼저 선택하세요"
                : !studentId
                ? "학생을 선택하세요"
                : !toClass
                ? "이동할 반을 선택하세요"
                : fromClass === toClass
                ? "같은 반으로는 이동할 수 없습니다"
                : undefined
            }
          >
            이동
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}
