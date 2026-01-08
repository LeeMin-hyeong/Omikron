import { useEffect, useMemo, useState } from "react";
import type { ViewProps } from "@/types/omikron";
import { rpc } from "pyloid-js";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle-group";
import { Input } from "@/components/ui/input";
import { Spinner } from "@/components/ui/spinner";
import { Check } from "lucide-react";
import { ScrollArea } from "@/components/ui/scroll-area";

type ClassStudentDict = Record<string, Record<string, number>>;
type AisosicClassStudentDict = Record<string, string[]>;

type StudentStatus = "ok" | "missing" | "other-class" | "data-only";

type StudentItem = {
  id: string;
  name: string;
  className: string;
  status: StudentStatus;
  dataClassName?: string;
};

const displayClassName = (name?: string) =>
  !name || name === "undefined" || name === "null" ? "미지정" : name;

const STATUS_LABELS: Record<StudentStatus, string> = {
  ok: "OK",
  missing: "아이소식에 추가됨",
  "other-class": "이동됨",
  "data-only": "퇴원 처리됨",
};

const statusTextClass = (status: StudentStatus) => {
  if (status === "missing") return "text-emerald-600";
  if (status === "other-class") return "text-amber-600";
  if (status === "data-only") return "text-red-600";
  return "";
};

const statusDotClass = (status: StudentStatus) => {
  if (status === "missing") return "bg-emerald-500";
  if (status === "other-class") return "bg-amber-500";
  if (status === "data-only") return "bg-red-500";
  return "bg-slate-300";
};

const normalizeDatafile = (data: unknown): ClassStudentDict => {
  if (Array.isArray(data) && data.length >= 1 && typeof data[0] === "object" && data[0] !== null) {
    return data[0] as ClassStudentDict;
  }
  if (data && typeof data === "object") {
    const maybe = data as { class_student_dict?: unknown };
    if (maybe.class_student_dict && typeof maybe.class_student_dict === "object") {
      return maybe.class_student_dict as ClassStudentDict;
    }
  }
  return {};
};

const normalizeAisosic = (data: unknown): AisosicClassStudentDict => {
  if (!data || typeof data !== "object" || Array.isArray(data)) return {};
  const dict: AisosicClassStudentDict = {};
  Object.entries(data as Record<string, unknown>).forEach(([className, students]) => {
    if (Array.isArray(students)) {
      dict[className] = students.map((s) => String(s)).filter(Boolean);
    }
  });
  return dict;
};

const buildStudentsByClass = (
  datafileMap: ClassStudentDict,
  aisosicMap: AisosicClassStudentDict
) => {
  const dataClassNames = Object.keys(datafileMap);
  const classNames = dataClassNames.filter((name) => aisosicMap[name]).sort();

  const dataStudentClassMap = new Map<string, string>();
  dataClassNames.forEach((className) => {
    Object.keys(datafileMap[className] || {}).forEach((studentName) => {
      if (!dataStudentClassMap.has(studentName)) {
        dataStudentClassMap.set(studentName, className);
      }
    });
  });

  const aisosicAllSet = new Set<string>();
  const aisosicStudentClassCount = new Map<string, number>();
  Object.values(aisosicMap).forEach((students) => {
    students.forEach((name) => aisosicAllSet.add(name));
  });
  Object.values(aisosicMap).forEach((students) => {
    students.forEach((name) => {
      aisosicStudentClassCount.set(name, (aisosicStudentClassCount.get(name) ?? 0) + 1);
    });
  });

  const studentsByClass: Record<string, StudentItem[]> = {};

  classNames.forEach((className) => {
    const dataStudents = Object.keys(datafileMap[className] || {});
    const dataSet = new Set(dataStudents);
    const aisosicStudents = aisosicMap[className] || [];

    const items: StudentItem[] = [];

    aisosicStudents.forEach((name, index) => {
      let status: StudentStatus = "ok";
      let dataClassName: string | undefined;
      if (!dataSet.has(name)) {
        const otherClass = dataStudentClassMap.get(name);
        const duplicatedInAisosic = (aisosicStudentClassCount.get(name) ?? 0) > 1;
        if (otherClass && !duplicatedInAisosic) {
          status = "other-class";
          dataClassName = otherClass;
        } else {
          status = "missing";
        }
      }
      items.push({
        id: `${className}::${name}::aisosic::${index}`,
        name,
        className,
        status,
        dataClassName,
      });
    });

    dataStudents.forEach((name, index) => {
      if (!aisosicAllSet.has(name)) {
        items.push({
          id: `${className}::${name}::data::${index}`,
          name,
          className,
          status: "data-only",
        });
      }
    });

    items.sort((a, b) => a.name.localeCompare(b.name));
    studentsByClass[className] = items;
  });

  return { classNames, studentsByClass };
};

function StudentList({
  title,
  classes,
  itemsByClass,
  selectedId,
  onSelect,
  loading,
}: {
  title: string;
  classes: string[];
  itemsByClass: Record<string, StudentItem[]>;
  selectedId: string;
  onSelect: (id: string) => void;
  loading: boolean;
}) {
  return (
    <Card className="flex h-full flex-col pt-2 gap-1 pb-1">
      <CardHeader className="space-y-0 py-1 p-0 justify-center my-0">
        <CardTitle className="text-base font-semibold p-0">{title}</CardTitle>
      </CardHeader>
      <CardContent className="flex min-h-0 flex-1 flex-col px-1">
        <div className="relative min-h-0 flex-1 rounded-lg border">
          {loading && (
            <div className="absolute inset-0 z-10 flex items-center justify-center bg-background/60">
              <div className="flex items-center gap-2 text-sm text-muted-foreground">
                <Spinner />
                불러오는 중...
              </div>
            </div>
          )}
          <ScrollArea className="h-[410px] w-full overflow-y-auto p-2">
            <div className="space-y-2 w-full min-w-0">
              {!loading && classes.length === 0 && (
                <div className="absolute inset-0 flex items-center justify-center text-sm text-muted-foreground">
                  표시할 항목이 없습니다
                </div>
              )}
              {!loading &&
                classes.map((className) => {
                  const items = itemsByClass[className] ?? [];
                  const classLabel = displayClassName(className);
                  return (
                    <div key={className} className="rounded-md border border-transparent">
                      <div className="px-2 py-2 text-xs font-semibold text-muted-foreground">
                        {classLabel}
                      </div>
                      {items.length === 0 ? (
                        <div className="px-4 pb-2 text-xs text-muted-foreground">
                          표시할 학생이 없습니다
                        </div>
                      ) : (
                        <ul className="space-y-1 pb-2">
                          {items.map((item) => {
                            const isSel = selectedId === item.id;
                            const detail =
                              item.status === "other-class" && item.dataClassName
                                ? ` - ${displayClassName(item.dataClassName)} 반으로 이동됨`
                                : item.status === "missing"
                                ? " - 추가됨"
                                : item.status === "data-only"
                                ? " - 퇴원 처리됨"
                                : "";
                            return (
                              <li className="px-2" key={item.id}>
                                <button
                                  type="button"
                                  onClick={() => onSelect(item.id)}
                                  disabled={loading}
                                  className={`group flex w-full items-start gap-2 rounded-md border px-2 py-1 text-left text-sm transition ${
                                    isSel ? "bg-blue-50 border-blue-200" : "hover:bg-accent border-transparent"
                                  } ${loading ? "opacity-60 cursor-not-allowed" : ""}`}
                                >
                                  <span
                                    className={[
                                      "flex-1 min-w-0 break-all text-xs leading-5",
                                      statusTextClass(item.status),
                                    ].join(" ")}
                                  >
                                    {item.name}
                                    {detail}
                                  </span>
                                  <span className="h-4 w-4 shrink-0 grid place-items-center">
                                    {isSel ? <Check className="h-4 w-4 text-blue-600" /> : null}
                                  </span>
                                </button>
                              </li>
                            );
                          })}
                        </ul>
                      )}
                    </div>
                  );
                })}
            </div>
          </ScrollArea>
        </div>
      </CardContent>
    </Card>
  );
}

export default function ManageStudentView({ meta }: ViewProps) {
  const dialog = useAppDialog();

  const [classes, setClasses] = useState<string[]>([]);
  const [studentsByClass, setStudentsByClass] = useState<Record<string, StudentItem[]>>({});
  const [selectedStudentId, setSelectedStudentId] = useState<string>("");
  const [colorFilters, setColorFilters] = useState<string[]>(["green", "orange", "red"]);
  const [query, setQuery] = useState<string>("");
  const [loading, setLoading] = useState(false);
  const [actionRunning, setActionRunning] = useState<null | "add" | "move" | "remove">(null);

  const allStudents = useMemo(
    () => classes.flatMap((className) => studentsByClass[className] || []),
    [classes, studentsByClass]
  );
  const selectedStudent = allStudents.find((s) => s.id === selectedStudentId);

  const filteredStudentsByClass = useMemo(() => {
    const result: Record<string, StudentItem[]> = {};
    const q = query.trim().toLowerCase();
    classes.forEach((className) => {
      const items = studentsByClass[className] || [];
      const classMatch = q && className.toLowerCase().includes(q);
      const byQuery = !q || classMatch
        ? items
        : items.filter((item) => item.name.toLowerCase().includes(q));
      if (colorFilters.length === 0) {
        result[className] = byQuery;
        return;
      }
      result[className] = byQuery.filter((item) => {
        if (item.status === "missing") return colorFilters.includes("green");
        if (item.status === "other-class") return colorFilters.includes("orange");
        if (item.status === "data-only") return colorFilters.includes("red");
        return false;
      });
    });
    return result;
  }, [classes, studentsByClass, colorFilters, query]);

  const visibleClasses = useMemo(() => {
    if (colorFilters.length === 0 && !query.trim()) return classes;
    return classes.filter((className) => (filteredStudentsByClass[className]?.length ?? 0) > 0);
  }, [classes, filteredStudentsByClass, colorFilters, query]);

  const loadData = async () => {
    try {
      setLoading(true);
      const [dfRes, aisosicRes] = await Promise.all([
        rpc.call("get_datafile_data", {}),
        rpc.call("get_aisosic_student_data", {}),
      ]);

      if (dfRes?.ok && aisosicRes?.ok) {
        const datafileMap = normalizeDatafile(dfRes.data);
        const aisosicMap = normalizeAisosic(aisosicRes.data);
        const { classNames, studentsByClass: nextStudents } = buildStudentsByClass(
          datafileMap,
          aisosicMap
        );

        setClasses(classNames);
        setStudentsByClass(nextStudents);
        setSelectedStudentId("");
      } else if (!dfRes?.ok) {
        await dialog.error({ title: "Datafile load failed", message: dfRes?.error || "" });
      } else if (!aisosicRes?.ok) {
        await dialog.error({ title: "Aisosic load failed", message: aisosicRes?.error || "" });
      }
    } catch (e) {
      setClasses([]);
      setStudentsByClass({});
      setSelectedStudentId("");
      console.error(e);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (!selectedStudentId) {
      return;
    }
    if (!allStudents.some((i) => i.id === selectedStudentId)) {
      setSelectedStudentId("");
    }
  }, [selectedStudentId, allStudents]);

  const canAdd = selectedStudent?.status === "missing";
  const canMove = selectedStudent?.status === "other-class";
  const canRemove = selectedStudent?.status === "data-only";
  const busy = loading || actionRunning !== null;

  const handleAdd = async () => {
    if (!selectedStudent || !canAdd) return;

    const yes = await dialog.warning({
      title: "학생을 추가할까요?",
      message: `${selectedStudent.name} 학생을\n‘${displayClassName(selectedStudent.className)}’ 반에 추가합니다.`,
      confirmText: "추가",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      setActionRunning("add");
      // target_student_name, target_class_name
      const res = await rpc.call("add_student", { 
        target_student_name: selectedStudent.name,
        target_class_name: selectedStudent.className,
      }); // 서버: {ok:true}
      if (res?.ok) {
        const warnings: string[] = Array.isArray(res?.warnings) ? res.warnings : [];
        if (warnings.length > 0) {
          await dialog.warning({
            title: `완료 (경고 ${warnings.length}건)`,
            message: warnings.join("\n"),
          });
        } else {
          await dialog.confirm({ title: "완료", message: "학생이 반에 추가되었습니다." });
        }
      } else {
        await dialog.error({ title: "학생 추가 실패", message: res?.error || "" });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    } finally {
      setActionRunning(null);
      handleRefresh()
    }
  };

  const handleMove = async () => {
    if (!selectedStudent || !canMove || !selectedStudent.dataClassName) return;

    const yes = await dialog.warning({
      title: "학생 반을 변경할까요?",
      message: `${selectedStudent.name} 학생을\n‘${displayClassName(selectedStudent.dataClassName)}’ → ‘${displayClassName(selectedStudent.className)}’ 로 이동합니다.`,
      confirmText: "변경",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      setActionRunning("move");
      // target_student_name, target_class_name, current_class_name
      const res = await rpc.call("move_student", {
        target_student_name: selectedStudent.name,   // row index string
        current_class_name:  selectedStudent.dataClassName,
        target_class_name:   selectedStudent.className,
      }); // {ok:true} 기대
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: `${selectedStudent.name} 학생을 ${selectedStudent.className} 반으로 이동하였습니다.` });
      } else {
        await dialog.error({ title: "학생 반 이동 실패", message: res?.error || "" });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    } finally {
      setActionRunning(null);
      handleRefresh()
    }
  };

  const handleRemove = async () => {
    if (!selectedStudent || !canRemove) return;

    const yes = await dialog.warning({
      title: "학생을 삭제할까요?",
      message: `‘${displayClassName(selectedStudent.className)}’반 ${selectedStudent.name} 학생을\n 삭제합니다.`,
      confirmText: "삭제",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      setActionRunning("remove");
      // target_student_name
      const res = await rpc.call("remove_student", {
        target_student_name: selectedStudent.name,
      }); // { ok: true } 기대
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: `${selectedStudent.name} 학생이 삭제되었습니다.` });
      } else {
        await dialog.error({ title: "학생 삭제 실패", message: res?.error || "" });
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    } finally {
      setActionRunning(null);
      handleRefresh()
    }
  };

  const handleRefresh = async () => {
    setSelectedStudentId("");
    setColorFilters(["green", "orange", "red"]);
    setQuery("");
    await loadData();
  };

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm pb-2">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="mt-1 text-sm text-muted-foreground">{meta.guide}</p>
        </div>
        <Separator className="mb-2" />

        <div className="mb-2 flex flex-row items-center gap-2">
          <Input
            type="search"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="반 검색 / 학생 검색"
            className="h-9 w-full rounded-xl"
            disabled={loading}
          />
          <ToggleGroup
            type="multiple"
            variant="outline"
            spacing={1}
            value={colorFilters}
            onValueChange={setColorFilters}
            className="shrink-0 rounded-xl border-border/80 gap-1"
          >
            <ToggleGroupItem
              value="green"
              className="px-3 w-27 data-[state=on]:bg-emerald-50 data-[state=on]:text-emerald-700 rounded-xl"
              disabled={loading}
            >
              추가됨
            </ToggleGroupItem>
            <ToggleGroupItem
              value="orange"
              className="px-3 w-27 data-[state=on]:bg-amber-50 data-[state=on]:text-amber-700 rounded-xl"
              disabled={loading}
            >
              반 이동됨
            </ToggleGroupItem>
            <ToggleGroupItem
              value="red"
              className="px-3 w-28 data-[state=on]:bg-red-50 data-[state=on]:text-red-700 rounded-xl"
              disabled={loading}
            >
              퇴원 처리됨
            </ToggleGroupItem>
          </ToggleGroup>
        </div>

        <div className="grid grid-cols-1 gap-2 lg:grid-cols-[2fr_1fr] flex-1 pb-2">
          <StudentList
            title="학생 목록"
            classes={visibleClasses}
            itemsByClass={filteredStudentsByClass}
            selectedId={selectedStudentId}
            onSelect={setSelectedStudentId}
            loading={loading}
          />

          <Card className="flex h-full flex-col rounded-2xl shadow-sm">
            <CardHeader className="space-y-0 pb-0 pt-2">
              <CardTitle className="text-base font-semibold">선택된 학생</CardTitle>
            </CardHeader>
            <CardContent className="flex min-h-0 flex-1 flex-col justify-between">
              <div className="space-y-3">
                <div className="rounded-lg border p-3 text-sm">
                  <div className="text-xs text-muted-foreground">학생</div>
                  <div className="font-medium">
                    {selectedStudent?.name || "-"}
                  </div>
                </div>
                <div className="rounded-lg border p-3 text-sm">
                  <div className="text-xs text-muted-foreground">반</div>
                  <div className="font-medium">
                    {selectedStudent ? displayClassName(selectedStudent.className) : "-"}
                  </div>
                </div>
                <div className="rounded-lg border p-3 text-sm">
                  <div className="text-xs text-muted-foreground">학생 상태</div>

                  {selectedStudent ? (
                    <div className="flex items-center gap-2">
                      <span className={`h-2 w-2 rounded-full ${statusDotClass(selectedStudent.status)}`} />
                      <div className="flex flex-row">
                        <span>{STATUS_LABELS[selectedStudent.status]}</span>
                        {selectedStudent?.status === "other-class" && selectedStudent.dataClassName ? (
                          <span>: {displayClassName(selectedStudent.dataClassName)}</span>
                        ) : null}
                      </div>
                    </div>
                  ) : (
                    <div>-</div>
                  )}
                </div>
              </div>

              <div className="mt-4 grid gap-2">
                <Button
                  className="rounded-xl bg-emerald-600 text-white disabled:bg-black disabled:text-white"
                  disabled={!canAdd || busy}
                  onClick={handleAdd}
                >
                  {actionRunning === "add" ? <><Spinner className="h-4 w-4" />작업 중...</> : "학생 추가"}
                </Button>
                <Button
                  className="rounded-xl bg-amber-500 text-white disabled:bg-black disabled:text-white"
                  disabled={!canMove || busy}
                  onClick={handleMove}
                >
                  {actionRunning === "move" ? <><Spinner className="h-4 w-4" />작업 중...</> : "반 이동"}
                </Button>
                <Button
                  className="rounded-xl bg-red-600 text-white disabled:bg-black disabled:text-white"
                  disabled={!canRemove || busy}
                  onClick={handleRemove}
                >
                  {actionRunning === "remove" ? <><Spinner className="h-4 w-4" />작업 중...</> : "학생 삭제"}
                </Button>
              </div>
            </CardContent>
          </Card>
        </div>

        <div className="mt-auto flex items-center justify-between">
          <div className="flex flex-col gap-1 text-sm text-muted-foreground">
            <div className="flex items-center gap-2">
              <span className="h-3 w-3 rounded bg-emerald-500" />
              <span>아이소식에 추가되었지만 데이터 파일에 존재하지 않는 학생</span>
            </div>
            <div className="flex items-center gap-2">
              <span className="h-3 w-3 rounded bg-amber-500" />
              <span>아이소식 및 데이터 파일에 존재하지만 실제 수강중인 반이 아닌 다른 반에 속해 있는 학생</span>
            </div>
            <div className="flex items-center gap-2">
              <span className="h-3 w-3 rounded bg-red-500" />
              <span>데이터 파일에만 존재하지만 아이소식에 존재하지 않는 학생</span>
            </div>
          </div>

          <Button
            className="rounded-xl"
            variant="outline"
            onClick={handleRefresh}
            disabled={busy}
            title="Reload"
          >
            {loading ? "불러오는 중…" : "새로고침"}
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}
