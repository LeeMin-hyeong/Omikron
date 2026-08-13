import { useEffect, useMemo, useState } from "react";
import type { ViewProps } from "@/shared/types/tdm";
import { generalRpc } from "@/api/rpc";
import { useAppDialog } from "@/shared/components/dialogs/app/useAppDialog";

import { Button } from "@/shared/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/shared/components/ui/card";
import { Separator } from "@/shared/components/ui/separator";
import { ToggleGroup, ToggleGroupItem } from "@/shared/components/ui/toggle-group";
import { Input } from "@/shared/components/ui/input";
import { Spinner } from "@/shared/components/ui/spinner";
import { Check } from "lucide-react";
import { ScrollArea } from "@/shared/components/ui/scroll-area";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/shared/components/ui/select";

type ClassStudentDict = Record<string, Record<string, number>>;
type AisosikClassStudentDict = Record<string, string[]>;

type StudentStatus = "ok" | "missing" | "other-class" | "data-only";

type StudentItem = {
  id: string;
  name: string;
  className: string;
  status: StudentStatus;
  dataClassCandidates?: string[];
  allowAdd?: boolean;
};

const displayClassName = (name?: string) =>
  !name || name === "undefined" || name === "null" ? "미지정" : name;

const errorMessage = (error: unknown) =>
  error instanceof Error ? error.message : String(error);

const statusTextClass = (status: StudentStatus) => {
  if (status === "missing") return "text-emerald-600";
  if (status === "other-class") return "text-amber-600";
  if (status === "data-only") return "text-red-600";
  return "";
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

const normalizeAisosik = (data: unknown): AisosikClassStudentDict => {
  if (!data || typeof data !== "object" || Array.isArray(data)) return {};
  const dict: AisosikClassStudentDict = {};
  Object.entries(data as Record<string, unknown>).forEach(([className, students]) => {
    if (Array.isArray(students)) {
      dict[className] = students.map((s) => String(s)).filter(Boolean);
    }
  });
  return dict;
};

const buildStudentsByClass = (
  datafileMap: ClassStudentDict,
  aisosikMap: AisosikClassStudentDict
) => {
  const dataClassNames = Object.keys(datafileMap).sort();
  const dataClassSet = new Set(dataClassNames);
  const filteredAisosikMap: AisosikClassStudentDict = {};
  Object.entries(aisosikMap).forEach(([className, students]) => {
    if (dataClassSet.has(className)) {
      filteredAisosikMap[className] = students;
    }
  });
  const aisosikClassNames = Object.keys(filteredAisosikMap);
  const classNames = dataClassNames;

  const dataClassesByStudent = new Map<string, Set<string>>();
  dataClassNames.forEach((className) => {
    Object.keys(datafileMap[className] || {}).forEach((studentName) => {
      const next = dataClassesByStudent.get(studentName) ?? new Set<string>();
      next.add(className);
      dataClassesByStudent.set(studentName, next);
    });
  });

  const aisosikClassesByStudent = new Map<string, Set<string>>();
  aisosikClassNames.forEach((className) => {
    (filteredAisosikMap[className] || []).forEach((studentName) => {
      const name = String(studentName);
      if (!name) return;
      const next = aisosikClassesByStudent.get(name) ?? new Set<string>();
      next.add(className);
      aisosikClassesByStudent.set(name, next);
    });
  });

  const allStudents = new Set<string>([
    ...dataClassesByStudent.keys(),
    ...aisosikClassesByStudent.keys(),
  ]);

  type StatusInfo = {
    status: StudentStatus;
    dataClassCandidates?: string[];
    allowAdd?: boolean;
  };
  type DiffInfo = { aisosikStatusByClass: Map<string, StatusInfo>; dataOnlyClasses: Set<string> };
  const diffByStudent = new Map<string, DiffInfo>();

  const toSortedArray = (set: Set<string>) => Array.from(set).sort();

  allStudents.forEach((studentName) => {
    const dataSet = dataClassesByStudent.get(studentName) ?? new Set<string>();
    const aisosikSet = aisosikClassesByStudent.get(studentName) ?? new Set<string>();

    const overlap = new Set<string>();
    dataSet.forEach((className) => {
      if (aisosikSet.has(className)) overlap.add(className);
    });

    const dataOnly = new Set<string>();
    dataSet.forEach((className) => {
      if (!aisosikSet.has(className)) dataOnly.add(className);
    });

    const aisosikOnly = new Set<string>();
    aisosikSet.forEach((className) => {
      if (!dataSet.has(className)) aisosikOnly.add(className);
    });

    const aisosikStatusByClass = new Map<string, StatusInfo>();
    const dataOnlyClasses = new Set<string>();

    overlap.forEach((className) => aisosikStatusByClass.set(className, { status: "ok" }));

    const dataList = toSortedArray(dataOnly);
    const aisosikList = toSortedArray(aisosikOnly);
    if (dataList.length === 0) {
      aisosikList.forEach((className) => {
        aisosikStatusByClass.set(className, { status: "missing", allowAdd: true });
      });
    } else if (aisosikList.length === 0) {
      dataList.forEach((className) => dataOnlyClasses.add(className));
    } else {
      // 어느 출발 반과 도착 반이 실제 이동 관계인지는 데이터만으로 확정할 수 없다.
      // 모든 가능한 출발 반을 UI에 제공하고 사용자가 직접 연결한다.
      aisosikList.forEach((className) => {
        aisosikStatusByClass.set(className, {
          status: "other-class",
          dataClassCandidates: dataList,
          // 아이소식 쪽 잔여 인원이 더 많으면 그 차이만큼은 신규 추가다.
          // 어떤 도착 반이 신규인지 알 수 없으므로 해당 항목에서 추가도 선택 가능하게 한다.
          allowAdd: aisosikList.length > dataList.length,
        });
      });
      // 데이터 쪽 잔여 인원이 더 많으면 그 차이만큼은 삭제 대상이다.
      // 이동을 먼저 수행하면 새 비교 결과에서 실제 삭제 대상만 남는다.
      if (dataList.length > aisosikList.length) {
        dataList.forEach((className) => dataOnlyClasses.add(className));
      }
    }

    diffByStudent.set(studentName, { aisosikStatusByClass, dataOnlyClasses });
  });

  const studentsByClass: Record<string, StudentItem[]> = {};

  classNames.forEach((className) => {
    studentsByClass[className] = [];
  });

  classNames.forEach((className) => {
    const aisosikStudents = filteredAisosikMap[className] || [];
    aisosikStudents.forEach((name, index) => {
      const studentName = String(name);
      if (!studentName) return;
      const diff = diffByStudent.get(studentName);
      const statusInfo = diff?.aisosikStatusByClass.get(className);
      const status = statusInfo?.status ?? "missing";
      studentsByClass[className].push({
        id: `${className}::${studentName}::aisosik::${index}`,
        name: studentName,
        className,
        status,
        dataClassCandidates: statusInfo?.dataClassCandidates,
        allowAdd: statusInfo?.allowAdd,
      });
    });
  });

  dataClassNames.forEach((className) => {
    const dataStudents = Object.keys(datafileMap[className] || {});
    dataStudents.forEach((name, index) => {
      const diff = diffByStudent.get(name);
      if (!diff?.dataOnlyClasses.has(className)) return;
      studentsByClass[className].push({
        id: `${className}::${name}::data::${index}`,
        name,
        className,
        status: "data-only",
      });
    });
  });

  classNames.forEach((className) => {
    studentsByClass[className].sort((a, b) => a.name.localeCompare(b.name));
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
    <Card className="flex h-full min-h-0 flex-col gap-1 overflow-hidden pb-1 pt-2">
      <CardHeader className="space-y-0 py-1 p-0 justify-center my-0">
        <CardTitle className="text-base font-semibold p-0">{title}</CardTitle>
      </CardHeader>
      <CardContent className="flex min-h-0 flex-1 flex-col px-1">
        <div className="relative min-h-0 flex-1 overflow-hidden rounded-lg border">
          {loading && (
            <div className="absolute inset-0 z-10 flex items-center justify-center bg-background/60">
              <div className="flex items-center gap-2 text-sm text-muted-foreground">
                <Spinner />
                불러오는 중...
              </div>
            </div>
          )}
          <ScrollArea className="h-full w-full overflow-y-auto p-2">
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
                              item.status === "other-class"
                                ? " - 반 이동 확인 필요"
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
  const [moveSourceClass, setMoveSourceClass] = useState("");

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
      const [dfRes, aisosikRes] = await Promise.all([
        generalRpc.call("get_datafile_data", {}),
        generalRpc.call("get_aisosik_student_data", {}),
      ]);

      if (dfRes?.ok && aisosikRes?.ok) {
        const datafileMap = normalizeDatafile(dfRes.data);
        const aisosikMap = normalizeAisosik(aisosikRes.data);
        const { classNames, studentsByClass: nextStudents } = buildStudentsByClass(
          datafileMap,
          aisosikMap
        );

        setClasses(classNames);
        setStudentsByClass(nextStudents);
        setSelectedStudentId("");
      } else if (!dfRes?.ok) {
        await dialog.error({ title: "Datafile load failed", message: dfRes?.error || "", detail: dfRes?.detail });
      } else if (!aisosikRes?.ok) {
        await dialog.error({ title: "아이소식 데이터 수집 실패", message: aisosikRes?.error || "", detail: aisosikRes?.detail });
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

  useEffect(() => {
    setMoveSourceClass("");
  }, [selectedStudentId]);

  const canAdd = selectedStudent?.status === "missing" || selectedStudent?.allowAdd === true;
  const canMove =
    selectedStudent?.status === "other-class" &&
    (selectedStudent.dataClassCandidates?.includes(moveSourceClass) ?? false);
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
      const res = await generalRpc.call("add_student", {
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
        await dialog.error({ title: "학생 추가 실패", message: res?.error || "", detail: res?.detail });
      }
    } catch (error: unknown) {
      await dialog.error({ title: "오류", message: errorMessage(error) });
    } finally {
      setActionRunning(null);
      handleRefresh()
    }
  };

  const handleMove = async () => {
    if (!selectedStudent || !canMove || !moveSourceClass) return;

    const yes = await dialog.warning({
      title: "학생 반을 변경할까요?",
      message: `${selectedStudent.name} 학생을\n‘${displayClassName(moveSourceClass)}’ → ‘${displayClassName(selectedStudent.className)}’ 로 이동합니다.`,
      confirmText: "변경",
      cancelText: "취소",
    });
    if (!yes) return;

    try {
      setActionRunning("move");
      // target_student_name, target_class_name, current_class_name
      const res = await generalRpc.call("move_student", {
        target_student_name: selectedStudent.name,   // row index string
        current_class_name: moveSourceClass,
        target_class_name:   selectedStudent.className,
      }); // {ok:true} 기대
      if (res?.ok) {
        const warnings: string[] = Array.isArray(res?.warnings) ? res.warnings : [];
        if (warnings.length > 0) {
          await dialog.warning({
            title: `이동 완료 (경고 ${warnings.length}건)`,
            message: warnings.join("\n"),
          });
        } else {
          await dialog.confirm({ title: "완료", message: `${selectedStudent.name} 학생을 ${selectedStudent.className} 반으로 이동하였습니다.` });
        }
      } else {
        await dialog.error({ title: "학생 반 이동 실패", message: res?.error || "", detail: res?.detail });
      }
    } catch (error: unknown) {
      await dialog.error({ title: "오류", message: errorMessage(error) });
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
      const res = await generalRpc.call("remove_student", {
        target_class_name: selectedStudent.className,
        target_student_name: selectedStudent.name,
      }); // { ok: true } 기대
      if (res?.ok) {
        await dialog.confirm({ title: "완료", message: `${selectedStudent.name} 학생이 삭제되었습니다.` });
      } else {
        await dialog.error({ title: "학생 삭제 실패", message: res?.error || "", detail: res?.detail });
      }
    } catch (error: unknown) {
      await dialog.error({ title: "오류", message: errorMessage(error) });
    } finally {
      setActionRunning(null);
      handleRefresh()
    }
  };

  const handleRefresh = async () => {
    setSelectedStudentId("");
    setMoveSourceClass("");
    setColorFilters(["green", "orange", "red"]);
    setQuery("");
    await loadData();
  };

  return (
    <Card className="h-full min-h-0 gap-3 overflow-hidden rounded-2xl border-border/80 py-4 shadow-sm">
      <CardContent className="flex h-full min-h-0 flex-col px-4">
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

        <div className="grid min-h-0 flex-1 grid-cols-1 gap-2 pb-2 lg:grid-cols-[2fr_1fr]">
          <StudentList
            title="학생 목록"
            classes={visibleClasses}
            itemsByClass={filteredStudentsByClass}
            selectedId={selectedStudentId}
            onSelect={setSelectedStudentId}
            loading={loading}
          />

          <Card className="flex h-full min-h-0 gap-3 overflow-hidden rounded-2xl py-3 shadow-sm">
            {/* <CardHeader className="space-y-0 pb-0 pt-2">
              <CardTitle className="text-base font-semibold">선택된 학생</CardTitle>
            </CardHeader> */}
            <CardContent className="flex min-h-0 flex-1 flex-col justify-between px-4">
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
                {selectedStudent?.status === "other-class" ? (
                  <div className="rounded-lg border p-3 text-sm">
                    <div className="text-xs text-muted-foreground">
                      이동할 반 선택
                    </div>
                    <Select
                      value={moveSourceClass}
                      onValueChange={setMoveSourceClass}
                      disabled={busy}
                    >
                      <SelectTrigger className="mt-2">
                        <SelectValue placeholder="이동할 반 선택" />
                      </SelectTrigger>
                      <SelectContent>
                        {(selectedStudent.dataClassCandidates ?? []).map((className) => (
                          <SelectItem key={className} value={className}>
                            {displayClassName(selectedStudent.className)}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    {selectedStudent.allowAdd ? (
                      <p className="mt-2 text-xs text-muted-foreground">
                        아이소식 인원이 더 많아 이 학생은 이동 또는 신규 추가 중 하나를 선택할 수 있습니다.
                      </p>
                    ) : null}
                  </div>
                ) : null}
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

        <div className="flex shrink-0 items-center justify-between pt-1">
          <div className="flex flex-col gap-1 text-sm text-muted-foreground">
            <div className="flex items-center gap-2">
              <span className="h-3 w-3 rounded bg-emerald-500" />
              <span>동일 반을 제외한 뒤 아이소식 쪽에만 남아 신규 추가가 필요한 학생</span>
            </div>
            <div className="flex items-center gap-2">
              <span className="h-3 w-3 rounded bg-amber-500" />
              <span>양쪽에 다른 반이 남아 출발 반을 선택하여 이동해야 하는 학생</span>
            </div>
            <div className="flex items-center gap-2">
              <span className="h-3 w-3 rounded bg-red-500" />
              <span>동일 반과 이동 후보를 제외한 뒤 데이터 쪽에만 남아 삭제가 필요한 학생</span>
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
