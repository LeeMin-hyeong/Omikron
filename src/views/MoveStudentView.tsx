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
import { ChevronsRight, School, User } from "lucide-react";

type ClassInfo = { id: string; name: string };
type StudentInfo = { id: string; name: string };

type ClassStudentDict = Record<string, Record<string, number>>; // {class: {studentName: rowIndex}}

export default function MoveStudentView({ onAction }: ViewProps) {
  const { enforcePrereq } = usePrereq();
  const dialog = useAppDialog();

  // 서버 원본 맵
  const [classStudentMap, setClassStudentMap] = useState<ClassStudentDict>({});

  // 데이터
  const [classes, setClasses] = useState<ClassInfo[]>([]);
  const [fromClass, setFromClass] = useState<string>("");
  const [toClass, setToClass] = useState<string>("");
  const [students, setStudents] = useState<StudentInfo[]>([]);
  const [studentId, setStudentId] = useState<string>("");

  const [loading, setLoading] = useState(false);

  const selectedStudent = useMemo(
    () => students.find((s) => s.id === studentId)?.name ?? "",
    [students, studentId]
  );

  // ✅ 공통 로더: 반/학생 맵 로드
  const loadData = async () => {
    try {
      setLoading(true);
      const res = await rpc.call("get_datafile_data", {}); // [class_student_dict, _]
      let csd: ClassStudentDict = {};
      if (Array.isArray(res) && res.length >= 1 && typeof res[0] === "object") {
        csd = res[0] as ClassStudentDict;
      } else if (res?.class_student_dict) {
        csd = res.class_student_dict as ClassStudentDict;
      }
      setClassStudentMap(csd);

      const classNames = Object.keys(csd).sort();
      setClasses(classNames.map((name) => ({ id: name, name })));

      // 기존 선택값 보정
      if (fromClass && !csd[fromClass]) {
        setFromClass("");
        setStudents([]);
        setStudentId("");
      }
      if (toClass && !csd[toClass]) setToClass("");
    } catch {
      setClassStudentMap({});
      setClasses([]);
      setFromClass("");
      setToClass("");
      setStudents([]);
      setStudentId("");
    } finally {
      setLoading(false);
    }
  };

  // 최초 로드
  useEffect(() => {
    loadData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 반 선택 시: 맵에서 바로 학생 목록 계산
  useEffect(() => {
    setStudents([]);
    setStudentId("");
    if (!fromClass) return;
    const dict = classStudentMap[fromClass] || {};
    const list: StudentInfo[] = Object.entries(dict).map(([name, row]) => ({
      id: String(row), // 서버가 row index를 student_id로 받는다고 가정
      name,
    }));
    setStudents(list);
  }, [fromClass, classStudentMap]);

  const canSubmit = fromClass && toClass && studentId && fromClass !== toClass;

  const handleSubmit = async () => {
    if (!canSubmit) return;

    // 실행 전 파일 프리체크
    const ok = await enforcePrereq();
    if (!ok) return;

    const studentName = selectedStudent;
    const fromName = fromClass;
    const toName = toClass;

    const yes = await dialog.warning({
      title: "학생 반을 변경할까요?",
      message: `${studentName} 학생을\n‘${fromName}’ → ‘${toName}’ 로 이동합니다.`,
      confirmText: "변경",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      onAction?.("move-student");
      // target_student_name, target_class_name, current_class_name
      const res = await rpc.call("move_student", {
        target_student_name: studentName,   // row index string
        current_class_name:  fromClass,
        target_class_name:   toClass,
      }); // {ok:true} 기대
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: "학생 반이 변경되었습니다." });
        setToClass("");
      } else {
        await dialog.error({ title: "실패", message: res?.error || "변경에 실패했습니다." });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    }
  };

  const handleRefresh = async () => {
    setToClass("");
    setFromClass("");
    setStudentId("");
    await loadData();
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
                <Select value={fromClass} onValueChange={setFromClass}>
                  <SelectTrigger className="w-full rounded-xl" disabled={loading}>
                    <SelectValue placeholder={loading ? "불러오는 중…" : "반 선택"} />
                  </SelectTrigger>
                  <SelectContent>
                    {classes.map((c) => (
                      <SelectItem key={c.id} value={c.id}>
                        {c.name}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>

                <Select
                  value={studentId}
                  onValueChange={setStudentId}
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

          {/* 우측 타일 (이동할 반 + 이동 버튼) */}
          <Card className="rounded-2xl shadow-sm">
            <CardContent className="flex h-60 flex-col justify-between p-4">
              {/* 중앙: 아이콘 + 타이틀 */}
              <div className="flex flex-col items-center gap-2 text-center">
                <School className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">이동할 반</div>
              </div>

              {/* 하단: 반 선택 + (바로 아래) 이동 버튼 */}
              <div className="grid gap-2">
                <Select
                  value={toClass}
                  onValueChange={setToClass}
                  disabled={!fromClass || !studentId}
                >
                  <SelectTrigger className="w-full rounded-xl">
                    <SelectValue placeholder="반 선택" />
                  </SelectTrigger>
                  <SelectContent>
                    {classes.map((c) => (
                      <SelectItem
                        key={`to-${c.id}`}
                        value={c.id}
                        disabled={c.id === fromClass}
                      >
                        {c.name}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>

                {/* ✅ 이동 버튼을 여기로 이동 */}
                <Button
                  className="w-full rounded-xl bg-black text-white"
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
        </div>

        {/* ✅ 우하단: 새로고침 버튼 */}
        <div className="mt-auto flex items-center justify-end">
          <Button
            className="rounded-xl"
            variant="outline"
            onClick={handleRefresh}
            disabled={loading}
            title="반/학생 목록을 다시 불러옵니다."
          >
            {loading ? "불러오는 중…" : "새로고침"}
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}
