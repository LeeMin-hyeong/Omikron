// src/views/SaveRetestView.tsx (재시험 화면)
import { useEffect, useMemo, useState } from "react";
import type { ViewProps } from "@/types/omikron";
import { rpc } from "pyloid-js";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";

import { Card, CardContent } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { BookCheck, Loader2, Play, User } from "lucide-react";

// 화면에서 쓰는 타입
type ClassInfo = { id: string; name: string };
type StudentInfo = { id: string; name: string }; // id = (class 시트) 학생 행 인덱스(문자열)
type TestInfo = { id: string; name: string };    // id = (재시험 시트) 시험 행 인덱스(문자열)

// 서버 맵 타입
type ClassStudentDict = Record<string, Record<string, number>>; // {반: {학생이름: row}}
type MakeUpMap        = Record<string, Record<string, number>>; // {학생이름: {시험명: row}}

export default function SaveRetestView({ onAction, meta }: ViewProps) {
  const dialog = useAppDialog();

  // 원본 맵
  const [classStudentMap, setClassStudentMap] = useState<ClassStudentDict>({});
  const [makeupMap, setMakeupMap] = useState<MakeUpMap>({});

  // 목록
  const [classes, setClasses]   = useState<ClassInfo[]>([]);
  const [students, setStudents] = useState<StudentInfo[]>([]);
  const [tests, setTests]       = useState<TestInfo[]>([]);

  const [classQuery, setClassQuery] = useState("");
  const [studentQuery, setStudentQuery] = useState("");
  const [testQuery, setTestQuery] = useState("");

  // 선택값
  const [klass, setKlass]         = useState<string>("");
  const [studentId, setStudentId] = useState<string>("");
  const [testId, setTestId]       = useState<string>("");
  const [score, setScore]         = useState<string>("");

  const [loading, setLoading] = useState(false);
  const [running, setRunning] = useState(false);

  const loadData = async () => {
    try {
      setLoading(true);
      // 1) 반/학생(클래스 시트) + 2) 재시험(학생→시험) 동시 로드
      const [datafileRes, makeupRes] = await Promise.all([
        rpc.call("get_datafile_data", {}),   // [class_student_dict, class_test_dict]
        rpc.call("get_makeuptest_data", {}), // {학생: {시험: row}}
      ]);

      if(datafileRes.ok && makeupRes.ok){
        // class_student_dict 파싱
        let csd: ClassStudentDict = {};
        if (Array.isArray(datafileRes.data) && typeof datafileRes.data[0] === "object") {
          csd = datafileRes.data[0] as ClassStudentDict;
        } else if (datafileRes.data?.class_student_dict) {
          csd = datafileRes.data.class_student_dict as ClassStudentDict;
        }
        setClassStudentMap(csd);
  
        // 재시험 맵 파싱
        const makeup: MakeUpMap = (makeupRes.data ?? {}) as MakeUpMap;
        setMakeupMap(makeup);
  
        // 반 목록 세팅
        const names = Object.keys(csd).sort();
        setClasses(names.map((n) => ({ id: n, name: n })));
  
        // 선택값 보정
        // if (klass && !csd[klass]) {
        setKlass("");
        setStudents([]); setStudentId("");
        setTests([]); setTestId("");
        setScore("")
        // }
      } else if (!datafileRes.ok) {
        await dialog.error({ title: "데이터 파일 데이터 수집 실패", message: datafileRes?.error || "" })
      } else if (!makeupRes.ok) {
        await dialog.error({ title: "재시험 명단 파일 데이터 수집 실패", message: makeupRes?.error || "" })
      }
    } catch {
      setClassStudentMap({});
      setMakeupMap({});
      setClasses([]);
      setStudents([]); setTests([]);
      setKlass(""); setStudentId(""); setTestId("");
      setScore("")
    } finally {
      setLoading(false);
    }
  }

  // 초기 로드: 반/학생 맵 + 재시험 맵
  useEffect(() => {
    loadData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 반 선택 → 학생 목록 계산
  useEffect(() => {
    setStudents([]); setStudentId("");
    setTests([]); setTestId("");
    setStudentQuery("");
    setTestQuery("");

    if (!klass) return;

    const sDict = classStudentMap[klass] || {};
    const sList: StudentInfo[] = Object.entries(sDict).map(([name, row]) => ({
      id: String(row), // (클래스 시트) 학생 행 인덱스
      name,
    }));
    setStudents(sList);
  }, [klass, classStudentMap]);

  // 선택된 학생 이름 (재시험 맵 조회에 필요)
  const studentName = useMemo(
    () => students.find((s) => s.id === studentId)?.name ?? "",
    [students, studentId]
  );

  // 학생 선택 → 재시험 대상 시험 목록 계산
  useEffect(() => {
    setTests([]); setTestId("");
    setTestQuery("");
    if (!studentName) return;

    const tDict = makeupMap[studentName] || {};
    const tList: TestInfo[] = Object.entries(tDict).map(([testName, row]) => ({
      id: String(row), // (재시험 시트) 시험 행 인덱스
      name: testName,
    }));
    setTests(tList);
  }, [studentName, makeupMap]);

  // 표시용
  const testName = useMemo(
    () => tests.find((t) => t.id === testId)?.name ?? "",
    [tests, testId]
  );

  const filteredClasses = useMemo(() => {
    const q = classQuery.trim().toLowerCase();
    if (!q) return classes;
    return classes.filter((c) => c.name.toLowerCase().includes(q));
  }, [classes, classQuery]);

  const filteredStudents = useMemo(() => {
    const q = studentQuery.trim().toLowerCase();
    if (!q) return students;
    return students.filter((s) => s.name.toLowerCase().includes(q));
  }, [students, studentQuery]);

  const filteredTests = useMemo(() => {
    const q = testQuery.trim().toLowerCase();
    if (!q) return tests;
    return tests.filter((t) => t.name.toLowerCase().includes(q));
  }, [tests, testQuery]);

  const scoreValid = score.trim() !== "";
  const canSave = klass && studentId && testId && scoreValid;

  const handleSave = async () => {
    if (!canSave) return;

    const yes = await dialog.warning({
      title: "재시험 점수 저장",
      message: `${studentName} / ${testName}\n점수: ${score}`,
      confirmText: "저장",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      setRunning(true);
      onAction?.("save-retest");
      // target_row:int, makeup_test_score:str
      const res = await rpc.call("save_retest_result", {
        target_row:        Number(testId), // (재시험 시트) 시험 행 인덱스 (문자열)
        makeup_test_score: score,
      });
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: "점수가 저장되었습니다." });
        setScore("");
      } else {
        await dialog.error({ title: "재시험 결과 저장 실패", message: res?.error || "" });
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
          {meta?.guide && (
            <p className="mt-1 text-sm text-muted-foreground">{meta.guide}</p>
          )}
        </div>
        <Separator className="mb-4" />

        {/* 한 장의 카드 내부 레이아웃 */}
        <div className="flex-1 h-full">
          <div className="grid grid-cols-2 h-full gap-6 items-stretch">
            {/* 좌측: 학생 쪽 */}
            <div className="h-full w-full rounded-2xl border bg-card p-4 pt-6">
              <div className="mb-4 flex flex-col items-center gap-2 text-center">
                <User className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">학생</div>
              </div>
              <div className="grid gap-2 w-full">
                <Input
                  type="search"
                  value={classQuery}
                  onChange={(e) => setClassQuery(e.target.value)}
                  placeholder="반 검색"
                  className="h-9 w-full rounded-lg"
                  disabled={loading || running}
                />
                <div className="rounded-lg border">
                  <ScrollArea className="h-[140px] w-full p-1">
                    {filteredClasses.length === 0 ? (
                      <div className="p-2 text-xs text-muted-foreground">반이 없습니다</div>
                    ) : (
                      <ul className="space-y-1">
                        {filteredClasses.map((c) => {
                          const isSel = klass === c.id;
                          return (
                            <li key={c.id}>
                              <button
                                type="button"
                                onClick={() => setKlass(c.id)}
                                disabled={loading || running}
                                className={`group flex w-full items-start gap-2 rounded-md border px-2 py-1 text-left text-xs transition ${
                                  isSel ? "bg-blue-50 border-blue-200" : "hover:bg-accent border-transparent"
                                } ${loading || running ? "opacity-60 cursor-not-allowed" : ""}`}
                              >
                                <span className="flex-1 min-w-0 break-all">{c.name}</span>
                              </button>
                            </li>
                          );
                        })}
                      </ul>
                    )}
                  </ScrollArea>
                </div>

                <Input
                  type="search"
                  value={studentQuery}
                  onChange={(e) => setStudentQuery(e.target.value)}
                  placeholder="학생 검색"
                  className="h-9 w-full rounded-lg"
                  disabled={!klass || loading || running}
                />
                <div className="rounded-lg border">
                  <ScrollArea className="h-[140px] w-full p-1">
                    {filteredStudents.length === 0 ? (
                      klass ? 
                        <div className="p-2 text-xs text-muted-foreground">학생이 없습니다</div> :
                        <div className="p-2 text-xs text-muted-foreground">반을 선택하세요</div>
                    ) : (
                      <ul className="space-y-1">
                        {filteredStudents.map((s) => {
                          const isSel = studentId === s.id;
                          return (
                            <li key={s.id}>
                              <button
                                type="button"
                                onClick={() => setStudentId(s.id)}
                                disabled={!klass || loading || running}
                                className={`group flex w-full items-start gap-2 rounded-md border px-2 py-1 text-left text-xs transition ${
                                  isSel ? "bg-blue-50 border-blue-200" : "hover:bg-accent border-transparent"
                                } ${!klass || loading || running ? "opacity-60 cursor-not-allowed" : ""}`}
                              >
                                <span className="flex-1 min-w-0 break-all">{s.name}</span>
                              </button>
                            </li>
                          );
                        })}
                      </ul>
                    )}
                  </ScrollArea>
                </div>
              </div>            </div>

            {/* 우측: 재시험/점수 쪽 */}
            <div className="h-full w-full rounded-2xl border bg-card p-4 pt-6">
              <div className="mb-4 flex flex-col items-center gap-2 text-center">
                <BookCheck className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">재시험</div>
              </div>
              <div className="grid gap-2">
                <div className="rounded-lg border">
                  <ScrollArea className="h-[335px] w-full p-1">
                    {filteredTests.length === 0 ? (
                      studentId ?
                        <div className="p-2 text-xs text-muted-foreground">시험이 없습니다</div> :
                        <div className="p-2 text-xs text-muted-foreground">학생을 선택하세요</div>
                    ) : (
                      <ul className="space-y-1">
                        {filteredTests.map((t) => {
                          const isSel = testId === t.id;
                          return (
                            <li key={t.id}>
                              <button
                                type="button"
                                onClick={() => setTestId(t.id)}
                                disabled={!studentId || loading || running}
                                className={`group flex w-full items-start gap-2 rounded-md border px-2 py-1 text-left text-xs transition ${
                                  isSel ? "bg-blue-50 border-blue-200" : "hover:bg-accent border-transparent"
                                } ${!studentId || loading || running ? "opacity-60 cursor-not-allowed" : ""}`}
                              >
                                <span className="flex-1 min-w-0 break-all">{t.name}</span>
                              </button>
                            </li>
                          );
                        })}
                      </ul>
                    )}
                  </ScrollArea>
                </div>

                <Input
                  className="rounded-lg"
                  placeholder="점수 입력 (맞은 개수/문제 개수)"
                  value={score}
                  onChange={(e) => setScore(e.target.value)}
                  disabled={!testId || !studentId || loading || running}
                />
              </div>
            </div>
          </div>
        </div>

        {/* 우하단 저장 버튼 */}
        <div className="mt-6 flex items-center justify-end gap-2">
          <Button
            className="rounded-xl"
            variant="outline"
            onClick={loadData}
            disabled={loading}
            title="반 목록을 다시 불러옵니다."
          >
            {loading ? "불러오는 중…" : "새로고침"}
          </Button>
          <Button
            className="rounded-xl bg-black text-white"
            disabled={!canSave || running}
            onClick={handleSave}
            title={
              !klass ? "반을 선택하세요"
              : !studentId ? "학생을 선택하세요"
              : !testId ? "시험을 선택하세요"
              : !scoreValid ? "올바른 점수를 입력하세요"
              : undefined
            }
          >
            {running ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Play className="mr-2 h-4 w-4" />}
            {running ? "저장 중..." : "저장"}
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}
