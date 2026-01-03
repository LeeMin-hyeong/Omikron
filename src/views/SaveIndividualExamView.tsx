// src/views/SaveIndividualExamView.tsx
import { useEffect, useMemo, useState } from "react";
import type { ViewProps } from "@/types/omikron";
import { rpc } from "pyloid-js";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";

import { Card, CardContent } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Checkbox } from "@/components/ui/checkbox";
import { User, BookCheck, Play, Loader2 } from "lucide-react";
import useHolidayDialog from "@/components/holiday-dialog/useHolidayDialog";

type ClassInfo = { id?: string; name: string };
type StudentInfo = { id: string; name: string }; // id = rowIndex(string)
type TestInfo = { id: string; name: string };    // id = colIndex(string)

type ClassStudentDict = Record<string, Record<string, number>>; // {class: {studentName: row}}
type ClassTestDict    = Record<string, Record<string, number>>; // {class: {testLabel: col}}

export default function SaveIndividualExamView({ onAction, meta }: ViewProps) {
  const dialog = useAppDialog();
  const { openHolidayDialog, lastHolidaySelection } = useHolidayDialog()

  // 서버 맵(그대로 보관)
  const [classStudentMap, setClassStudentMap] = useState<ClassStudentDict>({});
  const [classTestMap, setClassTestMap]       = useState<ClassTestDict>({});

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

  const [makeupChecked, setMakeupChecked] = useState(true);

  const [loading, setLoading] = useState(false);
  const [running, setRunning] = useState(false);

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

  const scoreNum = Number(score);
  const scoreValid = score.trim() !== "" && !Number.isNaN(scoreNum);
  const canSave = klass && studentId && testId && scoreValid;

  const loadData = async () => {
    try {
      setLoading(true);
      const res = await rpc.call("get_datafile_data", { mocktest: true }); // [class_student_dict, class_test_dict]
      if(res?.ok){
        let csd: ClassStudentDict = {};
        let ctd: ClassTestDict = {};
        
        if (Array.isArray(res.data)) {
          csd = (res.data[0] ?? {}) as ClassStudentDict;
          ctd = (res.data[1] ?? {}) as ClassTestDict;
        } else if (res.data?.class_student_dict) {
          csd = res.data.class_student_dict as ClassStudentDict;
          ctd = res.data.class_test_dict as ClassTestDict;
        }
        
        setClassStudentMap(csd);
        setClassTestMap(ctd);
        
        // 반 목록
        const classNames = Object.keys(csd).sort();
        setClasses(classNames.map((name) => ({ id: name, name })));
        
        // 기존 선택 유지/보정
        // if (klass && !csd[klass]) {
        setKlass("");
        setStudents([]); setStudentId("");
        setTests([]); setTestId("");
        setScore("")
        // }
      } else {
        await dialog.error({ title: "데이터 파일 데이터 수집 실패", message: res?.error || "" })
      }
    } catch {
      setClassStudentMap({});
      setClassTestMap({});
      setClasses([]);
      setScore("")
    } finally {
      setLoading(false);
    }
  }

  // 초기 로드: 한번에 반/학생/시험 사전 전체 받기
  useEffect(() => {
    loadData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 반 선택 시 학생/시험 목록을 맵에서 바로 계산
  useEffect(() => {
    setStudents([]); setStudentId("");
    setTests([]); setTestId("");
    setStudentQuery("");
    setTestQuery("");

    if (!klass) return;

    const sDict = classStudentMap[klass] || {};
    const tDict = classTestMap[klass] || {};

    const sList: StudentInfo[] = Object.entries(sDict).map(([name, row]) => ({
      id: String(row),
      name,
    }));
    const tList: TestInfo[] = Object.entries(tDict).map(([label, col]) => ({
      id: String(col),
      name: label, // 예: "24.09.27 중간고사"
    }));

    setStudents(sList);
    setTests(tList);
  }, [klass, classStudentMap, classTestMap]);

  // 표시용 이름
  const studentName = useMemo(
    () => students.find((s) => s.id === studentId)?.name ?? "",
    [students, studentId],
  );
  const testName = useMemo(
    () => tests.find((t) => t.id === testId)?.name ?? "",
    [tests, testId],
  );

  const handleSave = async () => {
    if (!canSave) return;

    let sel = lastHolidaySelection;
    if (!sel) {
      sel = await openHolidayDialog();
      if(!sel) return
    }

    const yes = await dialog.warning({
      title: "개별 시험 결과 저장",
      message: `${studentName} / ${testName}\n점수: ${scoreNum}`,
      confirmText: "저장",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      const cell = await rpc.call("is_cell_empty", {
        row: Number(studentId),
        col: Number(testId),
      });
      if(!cell.empty){
        const yes = await dialog.warning({
          title: "시험 결과 중복 경고",
          message: `${studentName} 학생의 ${testName} 결과가 이미 있습니다 (점수: ${cell.value})\n시험 결과를 덮어씌우겠습니까?\n${cell.value}점 → ${scoreNum}점`,
          confirmText: "저장",
          cancelText: "취소",
        })
        if(!yes) return
      }
      setRunning(true);
      onAction?.("save-individual-exam");
      //student_name:str, class_name:str, test_name:str, target_row:int, target_col:int, test_score:int|float, makeup_test_check:bool, makeup_test_date:dict
      const res = await rpc.call("save_individual_result", {
        student_name:      studentName,
        class_name:        klass,
        test_name:         testName.slice(11),
        target_row:        Number(studentId), // = row index (string)
        target_col:        Number(testId),    // = col index (string)
        test_score:        scoreNum,
        makeup_test_check: !makeupChecked, //
        makeup_test_date:  sel,
      }); // {ok:true} 기대
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: "점수가 저장되었습니다.\n시험 결과 메시지를 확인하고 전송해주세요." });
        setScore("");
      } else {
        await dialog.error({ title: "개별 시험 결과 저장 실패", message: res?.error || "" });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    } finally {
      setRunning(false);
      setTimeout(() => {
        loadData()
      }, 5000);
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
                          const value = c.id ?? c.name;
                          const isSel = klass === value;
                          return (
                            <li key={value}>
                              <button
                                type="button"
                                onClick={() => setKlass(value)}
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
              </div>
            </div>

            {/* 우측: 시험/점수 쪽 */}
            <div className="h-full w-full rounded-2xl border bg-card p-4 pt-6">
              <div className="mb-4 flex flex-col items-center gap-2 text-center">
                <BookCheck className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">시험</div>
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
                                disabled={!klass || loading || running}
                                className={`group flex w-full items-start gap-2 rounded-md border px-2 py-1 text-left text-xs transition ${
                                  isSel ? "bg-blue-50 border-blue-200" : "hover:bg-accent border-transparent"
                                } ${!klass || loading || running ? "opacity-60 cursor-not-allowed" : ""}`}
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
                  className="rounded-xl"
                  type="number"
                  inputMode="numeric"
                  step="1"
                  min="0"
                  max="100"
                  placeholder="점수 입력"
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
          <label htmlFor="makeup-check" className={`flex items-center justify-between rounded-xl border px-3 py-[7px] text-sm w-35 ${
                  makeupChecked ? "bg-blue-50 border-blue-200" : "hover:bg-accent"
                }`}>
            
            <Checkbox
              id="makeup-check"
              checked={makeupChecked}
              onCheckedChange={(v) => setMakeupChecked(Boolean(v))}
              disabled={running}
            />
            재시험 {makeupChecked ? "응시" : "미응시"}
          </label>
          <Button
            className="rounded-xl bg-black text-white"
            disabled={!canSave || loading || running}
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
