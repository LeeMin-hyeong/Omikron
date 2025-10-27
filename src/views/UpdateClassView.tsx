// src/views/UpdateClassView.tsx
import { useEffect, useState, useMemo } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { ScrollArea } from "@/components/ui/scroll-area";
import { ArrowLeft, ArrowRight, Check, FileSpreadsheet, } from "lucide-react";
import { ViewProps } from "@/types/omikron";
import { Separator } from "@/components/ui/separator";
import { Toggle } from "@/components/ui/toggle";
import { rpc } from "pyloid-js";
import { Spinner } from "@/components/ui/spinner";
import { Input } from "@/components/ui/input";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription, DialogFooter,
} from "@/components/ui/dialog";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle-group";

export type ClassItem = {
  id: string;
  name: string;
  aisosicOnly?: boolean; // 아이소식에는 있고 데이터파일엔 없음 → 초록
  dataOnly?: boolean;    // 데이터파일엔 있고 아이소식엔 없음 → 주황
};

type Col = "left" | "center" | "right";

function useSelectableList(initial: ClassItem[]) {
  const [items, setItems] = useState<ClassItem[]>(initial);
  const [selected, setSelected] = useState<Set<string>>(new Set());

  const toggle = (id: string) =>
    setSelected((p) => {
      const n = new Set(p);
      n.has(id) ? n.delete(id) : n.add(id);
      return n;
    });

  const clearSelection = () => setSelected(new Set());
  const selectAll = () => setSelected(new Set(items.map((i) => i.id)));
  const selectSome = (ids: string[]) =>
    setSelected((p) => {
      const n = new Set(p);
      ids.forEach((id) => n.add(id));
      return n;
    });
  const clearSome = (ids: string[]) =>
    setSelected((p) => {
      const n = new Set(p);
      ids.forEach((id) => n.delete(id));
      return n;
    });

  return { items, setItems, selected, setSelected, toggle, clearSelection, selectAll, selectSome, clearSome } as const;
}

function canMove(item: ClassItem, from: Col, to: Col) {
  if (item.aisosicOnly) {
    return (from === "left" && to === "center") || (from === "center" && to === "left");
  }
  // 나머지 색상(주황/기본): center ↔ right만
  return (from === "center" && to === "right") || (from === "right" && to === "center");
}

function guardedMove(
  from: Col,
  to: Col,
  lists: {
    left: ReturnType<typeof useSelectableList>;
    center: ReturnType<typeof useSelectableList>;
    right: ReturnType<typeof useSelectableList>;
  }
) {
  const src = lists[from];
  const dst = lists[to];

  // 선택 중 허용된 것만 이동
  const toMove = src.items.filter(i => src.selected.has(i.id) && canMove(i, from, to));
  if (toMove.length === 0) return; // 전부 금지면 아무 일도 없음

  const moveIdSet = new Set(toMove.map(i => i.id));
  const remain = src.items.filter(i => !moveIdSet.has(i.id));

  const dstIds = new Set(dst.items.map(d => d.id));
  const merged = [...dst.items, ...toMove.filter(i => !dstIds.has(i.id))];

  src.setItems(remain);
  dst.setItems(merged);

  // 이동된 항목만 선택 해제, 나머지는 그대로 유지
  const newSel = new Set(src.selected);
  toMove.forEach(i => newSel.delete(i.id));
  src.setSelected(newSel);
}

function ListBox({
  title,
  list,
  selected,
  onToggle,
  onSelectVisible,
  onClearVisible,
  loading,
}: {
  title: string;
  list: ClassItem[];
  selected: Set<string>;
  onToggle: (id: string) => void;
  onSelectVisible: (ids: string[]) => void;
  onClearVisible: (ids: string[]) => void;
  loading?: boolean;
}) {
  const visibleIds = useMemo(() => list.map((i) => i.id), [list]);
  const allSelected = list.length > 0 && list.every((i) => selected.has(i.id));

  return (
    <Card className="flex h-[470px] w-full flex-col py-1 gap-1">
      <CardHeader className="space-y-0 pb-0 pt-2 justify-center">
        <CardTitle className="text-base font-semibold pb-0">{title}</CardTitle>
      </CardHeader>
      <CardContent className="flex min-h-0 flex-1 flex-col px-1">
        <div className="flex items-center justify-end">
          <Toggle
            pressed={allSelected}
            onPressedChange={(pressed) => (pressed ? onSelectVisible(visibleIds) : onClearVisible(visibleIds))}
            aria-label="전체 선택 토글"
            className="h-8 px-3"
            disabled={loading}
          >
            {allSelected ? "전체 해제" : "전체 선택"}
          </Toggle>
        </div>

        <div className="relative min-h-0 flex-1 rounded-lg border">
          {loading && (
            <div className="absolute inset-0 z-10 flex items-center justify-center bg-background/60">
              <div className="flex items-center gap-2 text-sm text-muted-foreground">
                <Spinner />
                불러오는 중…
              </div>
            </div>
          )}

          <ScrollArea className="h-full w-full p-1">
            <ul className="space-y-1 w-full min-w-0">
              {!loading && list.length === 0 && (
                <div className="absolute inset-0 z-10 flex items-center justify-center bg-background/60">
                  <div className="flex items-center gap-2 text-sm text-muted-foreground">
                    항목이 없습니다
                  </div>
                </div>
              )}
              {!loading &&
                list.map((item) => {
                  const isSel = selected.has(item.id);
                  return (
                    <li className="pr-2" key={item.id}>
                      <button
                        type="button"
                        onClick={() => onToggle(item.id)}
                        disabled={loading}
                        className={`group flex w-full items-start gap-2 rounded-md border px-2 py-1.5 text-left text-sm transition ${
                          isSel ? "bg-blue-50 border-blue-200" : "hover:bg-accent border-transparent"
                        } ${loading ? "opacity-60 cursor-not-allowed" : ""}`}
                      >
                        <span
                          className={[
                            "flex-1 min-w-0 break-all text-xs leading-5",
                            item.aisosicOnly ? "text-emerald-600" : "",
                            item.dataOnly ? "text-amber-600" : "",
                          ].join(" ")}
                        >
                          {item.name}
                        </span>
                        <span className="h-4 w-4 shrink-0 grid place-items-center">
                          {isSel ? <Check className="h-4 w-4 text-blue-600" /> : null}
                        </span>
                      </button>
                    </li>
                  );
                })}
            </ul>
          </ScrollArea>
        </div>
      </CardContent>
    </Card>
  );
}

export default function UpdateClassView({ meta }: ViewProps) {
  const dialog = useAppDialog();
  const left = useSelectableList([] as ClassItem[]);
  const center = useSelectableList([] as ClassItem[]);
  const right = useSelectableList([] as ClassItem[]);

  const [loading, setLoading] = useState<boolean>(false);

  // 검색
  const [query, setQuery] = useState<string>("");
  const [colorFilters, setColorFilters] = useState<string[]>([]); // ["green"], ["orange"], ["green","orange"], []
  const q = query.trim().toLowerCase();
  const matchesColor = (i: ClassItem) => {
    const greenOn = colorFilters.includes("green");
    const orangeOn = colorFilters.includes("orange");
    if (!greenOn && !orangeOn) return true; // 색상 필터 미사용
    const isGreen = !!i.aisosicOnly; // 아이소식 전용 → 초록
    const isOrange = !!i.dataOnly;   // 데이터만 존재 → 주황
    return (greenOn && isGreen) || (orangeOn && isOrange);
  };

  const filterByQueryAndColor = (arr: ClassItem[]) => {
    const byQuery = !q ? arr : arr.filter((i) => i.name.toLowerCase().includes(q));
    return byQuery.filter(matchesColor);
  };

  // 기존 filteredLeft/Center/Right 교체
  const filteredLeft = useMemo(() => filterByQueryAndColor(left.items), [left.items, q, colorFilters]);
  const filteredCenter = useMemo(() => filterByQueryAndColor(center.items), [center.items, q, colorFilters]);
  const filteredRight = useMemo(() => filterByQueryAndColor(right.items), [right.items, q, colorFilters]);

  // 다이얼로그 상태 (3단계)
  const [step1Open, setStep1Open] = useState(false);
  const [step1Creating, setStep1Creating] = useState(false);
  const [step1Path, setStep1Path] = useState<string>();
  const [step1Checked, setStep1Checked] = useState(false);

  const [step2Open, setStep2Open] = useState(false);
  const [step2Checked, setStep2Checked] = useState(false);
  const [step2Loading, setStep2Loading] = useState(false);
  const [tempClasses, setTempClasses] = useState<string[]>([]);

  const [progressOpen, setProgressOpen] = useState(false);
  const [progressMsg, setProgressMsg] = useState<string>("");

  useEffect(() => {
    loadData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const loadData = async () => {
    try {
      setLoading(true);
      const [dfRes, aisosicRes] = await Promise.all([
        rpc.call("get_datafile_data", {}),
        rpc.call("get_aisosic_data", {}), // ★ 아이소식 전체 목록(string[])
      ]);

      // 데이터 파일의 반 목록
      let classStudentDict: Record<string, unknown> = {};
      if (Array.isArray(dfRes) && typeof dfRes[0] === "object") {
        classStudentDict = dfRes[0] as Record<string, unknown>;
      } else if (dfRes?.class_student_dict) {
        classStudentDict = dfRes.class_student_dict as Record<string, unknown>;
      }
      const currentNames = Object.keys(classStudentDict).sort();

      // 아이소식 전체 목록
      const aisosicAll: string[] = Array.isArray(aisosicRes) ? aisosicRes : [];
      const aisosicSet = new Set(aisosicAll);

      // 아이소식에는 있는데 데이터파일엔 없는 목록(=추가 후보 → 좌측, 초록)
      const aisosicOnlyNames = aisosicAll.filter((n) => !classStudentDict[n]);

      // 데이터파일에는 있지만 아이소식엔 없는 목록(=주황)
      const dataOnlyNames = currentNames.filter((n) => !aisosicSet.has(n));
      const dataOnlySet = new Set(dataOnlyNames);

      // 리스트 주입 (플래그를 아이템에 넣어서 이동해도 유지)
      center.setItems(
        currentNames.map((n) => ({
          id: n,
          name: n,
          dataOnly: dataOnlySet.has(n),  // 주황
        }))
      );

      left.setItems(
        aisosicOnlyNames.map((n) => ({
          id: n,
          name: n,
          aisosicOnly: true,             // 초록
        }))
      );

      right.setItems([]); // 초기엔 비움
      left.clearSelection(); center.clearSelection(); right.clearSelection();
    } catch (e) {
      left.setItems([]); center.setItems([]); right.setItems([]);
      left.clearSelection(); center.clearSelection(); right.clearSelection();
      console.error(e);
    } finally {
      setLoading(false);
    }
  };

  // 버튼 → 1단계 다이얼로그 오픈 & 임시파일 생성
  const openStep1File = async (path?: string) => {
    const targetPath = path ?? step1Path;
    if (!targetPath) return;
    try {
      const openRes = await rpc.call("open_path", { path: targetPath });
      if (!openRes?.ok) {
        const errorMessage = openRes?.error ?? "??? ? ? ????.";
        throw new Error(errorMessage);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      await dialog.error({ title: "??", message: `??? ? ? ????: ${message}` });
    }
  };

  const handleOpenStep1 = async () => {
    setStep1Open(true);
    setStep1Checked(false);
    setStep1Creating(true);
    setStep1Path(undefined);
    try {
      const classNames = center.items.map((i) => i.name);
      const res = await rpc.call("make_temp_class_info", { new_class_list: classNames }); // {ok, path}
      if (res?.ok) {
        setStep1Path(res.path);
        // 바로 열기
        // if (res.path) {
        //   await openStep1File(res.path);
        // }
      } else if (res?.error) {
        throw new Error(res.error);
      } else {
        throw new Error("임시 반 정보 생성 중 에러가 발생했습니다.");
      }
    } catch (e) {
      setStep1Open(false);
      const message = e instanceof Error ? e.message : String(e);
      await dialog.error({ title: "??", message: `임시 반 정보 생성 중 에러가 발생했습니다: ${message}` })
    } finally {
      setStep1Creating(false);
    }
  };

  // 1단계 다음 → 2단계 목록 로드
  const goStep2 = async () => {
    if (!step1Checked) return;
    setStep2Open(true);
    setStep2Checked(false);
    setStep2Loading(true);
    try {
      // 임시 파일에서 반 이름 리스트 가져오기
      const res = await rpc.call("get_new_class_list", {}); // string[]
      setTempClasses(Array.isArray(res) ? res : []);
    } catch (e) {
      setTempClasses([]);
      setStep2Open(false);
      await dialog.error({ title: "에러", message: `에러가 발생하였습니다: ${e}` })
    } finally {
      setStep2Loading(false);
      setStep1Open(false);
    }
  };

  // 2단계 최종 적용
  const applyUpdate = async () => {
    if (!step2Checked) return;

    setStep2Open(false);
    setProgressMsg("반 업데이트를 적용하는 중입니다…");
    setProgressOpen(true);

    try {
      const res = await rpc.call("update_class", {});
      if (res?.ok) {
        setProgressOpen(false);
        await dialog.confirm({title: "성공", message: "반 업데이트가 완료되었습니다."})
      } else {
        throw new Error(res?.error ?? "반 업데이트 적용 중 오류가 발생했습니다.");
      }
    } catch (e) {
      console.error(e);
      await dialog.error({ title: "에러", message: `반 업데이트 적용 중 오류가 발생했습니다: ${e}` });
    } finally {
      setProgressOpen(false);
      loadData();
    }
  };

  const handleCancle = async () => {
    try {
      const res = await rpc.call("delete_class_info_temp", {});
      if(res?.ok) {
        loadData();
      }
    } catch (e) {
      dialog.error({title: "에러", message: `${e}`})
    }
  }

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="mt-1 text-sm text-muted-foreground">{meta.guide}</p>
        </div>
        <Separator className="mb-4" />

        {/* 검색 */}
        <div className="mb-3 flex flex-row gap-1">
          <Input
            type="search"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="반 이름 검색"
            className="rounded-xl"
            disabled={loading}
          />
          <ToggleGroup
            type="multiple"
            variant="outline"
            spacing={1}
            value={colorFilters}
            onValueChange={setColorFilters}
            className="shrink-0 rounded-xl border-border/80"
          >
            <ToggleGroupItem
              value="green"
              // aria-label="초록만"
              className="px-3 data-[state=on]:bg-emerald-50 data-[state=on]:text-emerald-700 rounded-xl"
              disabled={loading}
            >
              추가할 반 보기
            </ToggleGroupItem>
            <ToggleGroupItem
              value="orange"
              // aria-label="주황만"
              className="px-3 data-[state=on]:bg-amber-50 data-[state=on]:text-amber-700 rounded-xl"
              disabled={loading}
            >
              삭제할 반 보기
            </ToggleGroupItem>
          </ToggleGroup>
        </div>

        <div className="grid grid-cols-14 gap-1 h-full">
          {/* Left */}
          <div className="col-span-4 h-full">
            <ListBox
              title="추가되지 않은 반"
              list={filteredLeft}
              selected={left.selected}
              onToggle={left.toggle}
              onSelectVisible={(ids) => left.selectSome(ids)}
              onClearVisible={(ids) => left.clearSome(ids)}
              loading={loading}
            />
          </div>

          {/* controls */}
          <div className="col-span-1 flex flex-col items-center justify-center gap-2">
            {/* Left → Center */}
            <Button
              variant="outline"
              onClick={() => guardedMove("left", "center", { left, center, right })}
              className="w-10"
              title="선택 → 가운데"
              disabled={loading}
            >
              <ArrowRight className="h-4 w-4" />
            </Button>

            {/* Center → Left */}
            <Button
              variant="outline"
              onClick={() => guardedMove("center", "left", { left, center, right })}
              className="w-10"
              title="선택 ← 왼쪽"
              disabled={loading}
            >
              <ArrowLeft className="h-4 w-4" />
            </Button>
          </div>

          {/* Center */}
          <div className="col-span-4 h-full">
            <ListBox
              title="현행(유지할) 반"
              list={filteredCenter}
              selected={center.selected}
              onToggle={center.toggle}
              onSelectVisible={(ids) => center.selectSome(ids)}
              onClearVisible={(ids) => center.clearSome(ids)}
              loading={loading}
            />
          </div>

          {/* controls */}
          <div className="col-span-1 flex flex-col items-center justify-center gap-2">
            {/* Center → Right */}
            <Button
              variant="outline"
              onClick={() => guardedMove("center", "right", { left, center, right })}
              className="w-10"
              title="선택 → 오른쪽"
              disabled={loading}
            >
              <ArrowRight className="h-4 w-4" />
            </Button>

            {/* Right → Center */}
            <Button
              variant="outline"
              onClick={() => guardedMove("right", "center", { left, center, right })}
              className="w-10"
              title="선택 ← 가운데"
              disabled={loading}
            >
              <ArrowLeft className="h-4 w-4" />
            </Button>
          </div>

          {/* Right */}
          <div className="col-span-4 h-full">
            <ListBox
              title="삭제할 반"
              list={filteredRight}
              selected={right.selected}
              onToggle={right.toggle}
              onSelectVisible={(ids) => right.selectSome(ids)}
              onClearVisible={(ids) => right.clearSome(ids)}
              loading={loading}
            />
          </div>
        </div>

        {/* 하단 버튼: '반 업데이트' */}
        <div className="flex items-center justify-between">
          <div className="flex flex-col gap-1">
            <div className="text-sm text-muted-foreground flex flex-row justify-start items-center">
              <div className="bg-green-500 w-4 h-4 rounded mr-1"/> <p>: 아이소식에는 존재하지만 데이터 파일에 존재하지 않는 반 (추가할 반)</p>
            </div>
            <div className="text-sm text-muted-foreground flex flex-row justify-start items-center">
              <div className="bg-amber-500 w-4 h-4 rounded mr-1"/> <p>: 아이소식에 존재하지 않지만 데이터 파일에는 존재하는 반 (삭제할 반)</p>
            </div>
          </div>
          <div className="flex gap-2">
            <Button
              className="rounded-xl"
              variant="outline"
              onClick={loadData}
              disabled={loading}
              title="반 목록을 다시 불러옵니다."
            >
              {loading ? "불러오는 중…" : "새로고침"}
            </Button>
            <Button className="rounded-xl bg-black text-white" onClick={handleOpenStep1} disabled={loading}>
              반 업데이트
            </Button>
          </div>
        </div>
      </CardContent>

      {/* 1단계: 임시 파일 생성 & 편집 안내 */}
      <Dialog open={step1Open} onOpenChange={(o) => setStep1Open(o)}>
        <DialogContent className="sm:max-w-[520px]" onInteractOutside={(e) => { e.preventDefault(); }}>
          <DialogHeader>
            <DialogTitle>'반 정보(임시).xlsx' 반 정보 수정</DialogTitle>
            <DialogDescription>
              {step1Creating ? "임시 파일을 생성 중입니다…" : "임시 반 정보 파일을 열어 수정하고 저장 및 종료해 주세요."}
            </DialogDescription>
          </DialogHeader>

          <div className="grid gap-3">
            {step1Creating ? (
              <div className="flex items-center justify-center py-6 text-sm text-muted-foreground">
                <Spinner /> <span className="ml-2">생성 중…</span>
              </div>
            ) : (
              <>
                <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border px-3 py-2 text-xs text-muted-foreground">
                  <div className="min-w-0 flex flex-row items-center">
                    <FileSpreadsheet className="text-green-600 mr-2"/> <span className="font-mono break-all">{step1Path}</span>
                  </div>
                  <Button size="sm" variant="outline" onClick={() => void openStep1File()}>
                    파일 열기
                  </Button>
                </div>
                <label className="flex items-center gap-2 rounded-xl border px-3 py-2">
                  <Checkbox checked={step1Checked} onCheckedChange={(v) => setStep1Checked(Boolean(v))} />
                  <span className="text-sm">반 정보를 <b>수정</b>한 후 <b>저장</b>하고 <b>종료</b>했습니다.</span>
                </label>
              </>
            )}
          </div>

          <DialogFooter>
            <Button variant="outline" onClick={() => {handleCancle(); setStep1Open(false)}} disabled={step1Creating}>취소</Button>
            <Button onClick={goStep2} disabled={!step1Checked || step1Creating}>
              다음
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* 2단계: 임시 파일 목록 확인 + 동의 체크 → 적용 */}
      <Dialog open={step2Open} onOpenChange={(o) => setStep2Open(o)}>
        <DialogContent className="sm:max-w-[560px]" onInteractOutside={(e) => { e.preventDefault(); }}>
          <DialogHeader>
            <DialogTitle>반 목록 확인</DialogTitle>
            <DialogDescription>임시 파일에 작성된 반 목록입니다. 확인 후 동의하시면 적용됩니다.</DialogDescription>
          </DialogHeader>

          <div className="rounded-lg border">
            <ScrollArea className="h-[260px] w-full p-2">
              {step2Loading ? (
                <div className="flex h-[240px] items-center justify-center text-sm text-muted-foreground">
                  <Spinner /> <span className="ml-2">불러오는 중…</span>
                </div>
              ) : tempClasses.length === 0 ? (
                <div className="p-3 text-sm text-muted-foreground">임시 파일에 반이 없습니다.</div>
              ) : (
                <ul className="space-y-1">
                  {tempClasses.map((name) => (
                    <li key={name} className="text-sm px-2 py-1 rounded hover:bg-accent/50">
                      {name}
                    </li>
                  ))}
                </ul>
              )}
            </ScrollArea>
          </div>

          <label className="flex items-center gap-2 rounded-xl border px-3 py-2">
            <Checkbox checked={step2Checked} onCheckedChange={(v) => setStep2Checked(Boolean(v))} />
            <span className="text-sm leading-5">
              <i>'반 정보(임시).xlsx'</i>에 <b>작성되어 있지 않은 반</b>의 시험 기록은 <b>데이터 파일에서 제거</b>되고
              <i> '지난 데이터.xlsx'</i>로 <b>이관</b>됨에 동의합니다.
            </span>
          </label>

          <DialogFooter>
            <Button variant="outline" onClick={() => {handleCancle(); setStep2Open(false)}}>취소</Button>
            <Button onClick={applyUpdate} disabled={!step2Checked || step2Loading}>
              반 업데이트 적용
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
      {/* 3단계: 반 업데이트 진행중 */}
      <Dialog open={progressOpen} onOpenChange={(o) => setProgressOpen(o)}>
        <DialogContent
          className="sm:max-w-[440px]"
          onInteractOutside={(e) => { e.preventDefault(); }} // 진행중엔 바깥 클릭으로 닫히지 않게
          showCloseButton={false}
        >
          <DialogHeader>
            <DialogTitle>반 업데이트 진행 중</DialogTitle>
            <DialogDescription>데이터 파일을 갱신하고 있습니다. 이 작업은 오래 걸립니다.</DialogDescription>
          </DialogHeader>

          <div className="flex items-center justify-center py-6 text-sm text-muted-foreground">
            <Spinner />
            <span className="ml-2">{progressMsg || "작업 중…"}</span>
          </div>
        </DialogContent>
      </Dialog>
    </Card>
  );
}
