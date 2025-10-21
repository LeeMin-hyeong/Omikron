// src/views/RemoveStudentView.tsx
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
import { UserMinus } from "lucide-react";
import { Spinner } from "@/components/ui/spinner";

/** 서버 반환 타입 헬퍼
 * class_student_dict: { [className]: { [studentName]: rowIndexNumber } }
 * class_test_dict   : { [className]: { [testLabel]: colIndexNumber } }
 */
type ClassStudentDict = Record<string, Record<string, number>>;
type ClassTestDict = Record<string, Record<string, number>>;

export default function RemoveStudentView({ onAction }: ViewProps) {
  const dialog = useAppDialog();

  // 서버에서 받은 전체 맵을 들고 있음
  const [classStudentMap, setClassStudentMap] = useState<ClassStudentDict>({});
  const [classTestMap, setClassTestMap] = useState<ClassTestDict>({}); // 현재 화면에선 사용 X (보존)
  // UI 상태
  const [classes, setClasses] = useState<Array<{ id: string; name: string }>>([]);
  const [fromClass, setFromClass] = useState<string>("");
  const [students, setStudents] = useState<Array<{ id: string; name: string }>>([]);
  const [student, setStudent] = useState<string>("");

  const [loading, setLoading] = useState(false);
  const [running, setRunning] = useState(false);

  /** 서버에서 (class_student_dict, class_test_dict) 받아오기 */
  const loadClasses = async () => {
    try {
      setLoading(true);
      const res = await rpc.call("get_datafile_data", {});
      // res가 [class_student_dict, class_test_dict] 형태로 들어옴
      let csd: ClassStudentDict = {};
      let ctd: ClassTestDict = {};

      if (Array.isArray(res) && res.length >= 1 && typeof res[0] === "object") {
        csd = res[0] as ClassStudentDict;
        if (res.length >= 2 && typeof res[1] === "object") {
          ctd = res[1] as ClassTestDict;
        }
      } else if (res?.class_student_dict) {
        // 혹시 서버가 객체 키로 감싸서 보낼 수도 있으니 호환 처리
        csd = res.class_student_dict as ClassStudentDict;
        ctd = res.class_test_dict as ClassTestDict;
      }

      setClassStudentMap(csd);
      setClassTestMap(ctd);

      // 클래스 목록 갱신
      const cls = Object.keys(csd).sort();
      const list = cls.map((name) => ({ id: name, name }));
      setClasses(list);

      // 선택 유지/초기화
      if (fromClass && !csd[fromClass]) {
        setFromClass("");
        setStudents([]);
        setStudent("");
      } else if (fromClass) {
        // 이미 선택된 반이 여전히 존재하면 학생 목록만 리프레시
        const dict = csd[fromClass] || {};
        const studs = Object.keys(dict).map((name) => ({ id: name, name }));
        setStudents(studs);
        if (student && !dict[student]) setStudent("");
      }
    } catch (e) {
      setClassStudentMap({});
      setClassTestMap({});
      setClasses([]);
      setFromClass("");
      setStudents([]);
      setStudent("");
    } finally {
      setLoading(false);
    }
  };

  /** 최초 로드 */
  useEffect(() => {
    loadClasses();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /** 반 선택 시 해당 반의 학생 목록을 classStudentMap에서 바로 채움 */
  useEffect(() => {
    setStudents([]);
    setStudent("");
    if (!fromClass) return;
    const dict = classStudentMap[fromClass] || {};
    const studs = Object.keys(dict).map((name) => ({ id: name, name }));
    setStudents(studs);
  }, [fromClass, classStudentMap]);

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
      //TODO
      // 서버측 remove_student_class는 이름과 반 이름을 사용
      const res = await rpc.call("remove_student_class", {
        student_name: student.trim(),
        from_class: fromClass,
      }); // { ok: true } 기대
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: `${student.trim()} 학생이 반에서 퇴원 처리되었습니다.` });
        // 학생 목록에서 제거
        setStudents((prev) => prev.filter((s) => s.name !== student));
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
    await dialog.confirm({ title: "새로고침", message: "반/학생 목록을 다시 불러왔습니다." });
  };

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="mt-1 text-sm text-muted-foreground">학생을 퇴원 처리합니다.</p>
        </div>
        <Separator className="mb-4" />

        {/* 중앙 정렬 컨테이너 */}
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
                {/* 반 선택 */}
                <Select value={fromClass} onValueChange={(v) => setFromClass(v)}>
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

                {/* 학생 선택: classStudentMap에서 바로 채움 */}
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
                      // 서버 remove_student_class가 학생 이름을 받으므로 value=이름으로 사용
                      <SelectItem key={s.id} value={s.name}>
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
            title="반/학생 목록을 다시 불러옵니다."
          >
            {loading ? "불러오는 중…" : "새로고침"}
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}
