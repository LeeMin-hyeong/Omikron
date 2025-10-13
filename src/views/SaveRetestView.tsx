// src/views/SaveIndividualExamView.tsx
import { useEffect, useMemo, useState } from "react";
import type { ViewProps } from "@/types/omikron";
import { rpc } from "pyloid-js";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";

import { Card, CardContent } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { Button } from "@/components/ui/button";
import {
  Select, SelectContent, SelectItem, SelectTrigger, SelectValue,
} from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { UserCircle2, FlaskConical, BookCheck, User } from "lucide-react";

type ClassInfo = { id?: string; name: string };
type StudentInfo = { id: string; name: string };
type TestInfo = { id: string; name: string };

export default function SaveRetestView({ onAction, meta }: ViewProps) {
  const dialog = useAppDialog();

  // 목록
  const [classes, setClasses] = useState<ClassInfo[]>([]);
  const [students, setStudents] = useState<StudentInfo[]>([]);
  const [tests, setTests] = useState<TestInfo[]>([]);

  // 선택값
  const [klass, setKlass] = useState<string>("");
  const [studentId, setStudentId] = useState<string>("");
  const [testId, setTestId] = useState<string>("");
  const [score, setScore] = useState<string>("");

  const loadingOk = (arr: any) => Array.isArray(arr) ? arr : [];

  // 초기: 반 목록
  useEffect(() => {
    (async () => {
      try {
        const res = await rpc.call("list_classes", {});
        setClasses(loadingOk(res));
      } catch { setClasses([]); }
    })();
  }, []);

  // 반 선택 시 학생/시험 목록 갱신
  useEffect(() => {
    setStudents([]); setStudentId("");
    setTests([]); setTestId("");
    if (!klass) return;

    (async () => {
      try {
        const [sRes, tRes] = await Promise.all([
          rpc.call("list_students_by_class", { class_name: klass }),
          rpc.call("list_tests", { class_name: klass }),            // 서버에서 [{id,name}] 반환
        ]);
        setStudents(loadingOk(sRes));
        setTests(loadingOk(tRes));
      } catch {
        setStudents([]); setTests([]);
      }
    })();
  }, [klass]);

  // 이름 표시용
  const studentName = useMemo(
    () => students.find(s => s.id === studentId)?.name ?? "",
    [students, studentId]
  );
  const testName = useMemo(
    () => tests.find(t => t.id === testId)?.name ?? "",
    [tests, testId]
  );

  const scoreNum = Number(score);
  const scoreValid = score.trim() !== "" && !Number.isNaN(scoreNum);
  const canSave = klass && studentId && testId && scoreValid;

  const handleSave = async () => {
    if (!canSave) return;

    // 2) 확인
    const yes = await dialog.warning({
      title: "개별 시험 결과 저장",
      message: `${studentName} / ${testName}\n점수: ${scoreNum}`,
      confirmText: "저장",
      cancelText: "취소",
    });
    if (!yes) return;

    // 3) RPC 실행
    try {
      onAction?.("save-individual-exam");
      const res = await rpc.call("save_individual_exam", {
        class_name: klass,
        student_id: studentId,
        test_id: testId,
        score: scoreNum,
      }); // 서버: {ok:true} 기대
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: "점수가 저장되었습니다." });
        setScore("");
      } else {
        await dialog.error({ title: "실패", message: res?.error || "저장에 실패했습니다." });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    }
  };

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          {meta?.guide && (
            <>
              <p className="mt-1 text-sm text-muted-foreground">{meta.guide}</p>
            </>
          )}
        </div>
        <Separator className="mb-4" />

        {/* 한 장의 카드 내부 레이아웃 */}
        <div className="flex-1 h-full justify-center items-center">
          <div className="grid grid-cols-2 h-full gap-6 place-items-center">
            {/* 좌측: 학생 쪽 */}
            <div className="h-[260px] w-full rounded-2xl border bg-card p-4 pt-18">
              <div className="mb-4 flex flex-col items-center gap-2 text-center">
                <User className="h-8 w-8 text-black pt" />
                <div className="text-sm font-medium">학생</div>
              </div>
              <div className="grid gap-3 w-full justify-c">
                {/* 반 선택 */}
                <Select value={klass} onValueChange={setKlass}>
                  <SelectTrigger className="rounded-xl w-full">
                    <SelectValue placeholder="반 선택" />
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
                {/* 학생 선택 */}
                <Select
                  value={studentId}
                  onValueChange={setStudentId}
                  disabled={!klass || students.length === 0}
                >
                  <SelectTrigger className="rounded-xl w-full">
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
            </div>

            {/* 우측: 시험/점수 쪽 */}
            <div className="h-[260px] w-full rounded-2xl border bg-card p-4 pt-18">
              <div className="mb-4 flex flex-col items-center gap-2 text-center">
                <BookCheck className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">재시험</div>
              </div>
              <div className="grid gap-3">
                {/* 시험 선택 */}
                <Select
                  value={testId}
                  onValueChange={setTestId}
                  disabled={!klass || tests.length === 0}
                >
                  <SelectTrigger className="rounded-xl w-full">
                    <SelectValue placeholder="시험 선택" />
                  </SelectTrigger>
                  <SelectContent>
                    {tests.map((t) => (
                      <SelectItem key={t.id} value={t.id}>
                        {t.name}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                {/* 점수 입력 */}
                <Input
                  className="rounded-xl"
                  type="number"
                  inputMode="decimal"
                  step="1"
                  min="0"
                  max="100"
                  placeholder="점수 입력 (맞은 개수/문제 개수)"
                  value={score}
                  onChange={(e) => setScore(e.target.value)}
                  disabled={!testId || !studentId}
                />
              </div>
            </div>
          </div>
        </div>

        {/* 우하단 저장 버튼 */}
        <div className="mt-6 flex items-center justify-end">
          <Button
            className="rounded-xl bg-black text-white"
            disabled={!canSave}
            onClick={handleSave}
            title={
              !klass ? "반을 선택하세요"
              : !studentId ? "학생을 선택하세요"
              : !testId ? "시험을 선택하세요"
              : !scoreValid ? "올바른 점수를 입력하세요"
              : undefined
            }
          >
            저장
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}
