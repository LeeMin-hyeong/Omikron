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
import { School } from "lucide-react";
import { Input } from "@/components/ui/input";
import { Spinner } from "@/components/ui/spinner";

export default function AddStudentView({ onAction }: ViewProps) {
  const dialog = useAppDialog();

  const [classes, setClasses] = useState<Array<{ id?: string; name: string }>>([]);
  const [toClass, setToClass] = useState<string>("");
  const [student, setStudent] = useState<string>("");
  const [loading, setLoading] = useState(false);
  const [running, setRunning] = useState(false);

  const loadClasses = async () => {
    try {
      setLoading(true);
      const res = await rpc.call("get_datafile_data", {}); // [class_student_dict, class_test_dict]

      // 응답 파싱(호환 처리)
      let classStudentDict: Record<string, unknown> = {};
      if (Array.isArray(res) && res.length >= 1 && typeof res[0] === "object") {
        classStudentDict = res[0] as Record<string, unknown>;
      } else if (res?.class_student_dict) {
        classStudentDict = res.class_student_dict as Record<string, unknown>;
      }

      // 반 목록 생성(이름 기준)
      const names = Object.keys(classStudentDict).sort();
      const list = names.map((name) => ({ id: name, name }));

      setClasses(list);

      // 선택값이 목록에 없으면 초기화
      if (toClass && !list.some((c) => (c.id ?? c.name) === toClass)) {
        setToClass("");
      }
    } catch {
      setClasses([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadClasses();
  }, []);

  const canSubmit = Boolean(toClass && student.trim());

  const handleSubmit = async () => {
    if (!canSubmit) return;

    const yes = await dialog.warning({
      title: "학생을 추가할까요?",
      message: `${student.trim()} 학생을\n‘${toClass}’ 반에 추가합니다.`,
      confirmText: "추가",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      setRunning(true);
      onAction?.("add-student");
      // TODO
      const res = await rpc.call("add_student_class", {
        student_name: student.trim(),
        to_class: toClass,
      }); // 서버: {ok:true}
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: "학생이 반에 추가되었습니다." });
        setStudent("");
      } else {
        await dialog.error({ title: "실패", message: res?.error || "추가에 실패했습니다." });
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
            학생을 추가합니다. 아이소식에서의 작업이 선행되어야 합니다.
          </p>
        </div>
        <Separator className="mb-4" />

        {/* ✅ 중앙 정렬 컨테이너 */}
        <div className="flex flex-1 items-center justify-center">
          <Card className="h-80 w-[420px] rounded-2xl shadow-sm">
            <CardContent className="flex h-full flex-col justify-between p-4">
              {/* 중앙: 아이콘 + 타이틀 */}
              <div className="flex flex-1 flex-col items-center justify-center gap-2 text-center">
                <School className="h-8 w-8 text-black" />
                <div className="text-sm font-medium">추가할 반</div>
              </div>

              {/* 하단 컨트롤 */}
              <div className="grid gap-2">
                <Select value={toClass} onValueChange={(v) => setToClass(v)}>
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

                <Input
                  className="rounded-xl"
                  value={student}
                  onChange={(e) => setStudent(e.target.value)}
                  placeholder="학생 이름"
                />

                <div className="mt-2 flex items-center gap-2">
                  <Button
                    className="w-full rounded-xl bg-black text-white"
                    disabled={!canSubmit}
                    onClick={handleSubmit}
                  >
                    {running ? <Spinner /> : "추가"}
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
