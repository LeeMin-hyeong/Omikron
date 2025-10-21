import { useEffect, useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { ScrollArea } from "@/components/ui/scroll-area";
import { ArrowLeft, ArrowRight, Check } from "lucide-react";
import { ViewProps } from "@/types/omikron";
import { Separator } from "@/components/ui/separator";
import { Toggle } from "@/components/ui/toggle";
import { rpc } from "pyloid-js";
import { Spinner } from "@/components/ui/spinner"; // ✅ 추가

export type ClassItem = { id: string; name: string };

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
  return { items, setItems, selected, setSelected, toggle, clearSelection, selectAll } as const;
}

function move(
  srcItems: ClassItem[],
  setSrc: (v: ClassItem[]) => void,
  dstItems: ClassItem[],
  setDst: (v: ClassItem[]) => void,
  selected: Set<string>,
  clearSel: () => void,
) {
  if (selected.size === 0) return;
  const toMove = srcItems.filter((i) => selected.has(i.id));
  if (!toMove.length) return;
  const remain = srcItems.filter((i) => !selected.has(i.id));
  const dstIds = new Set(dstItems.map((d) => d.id));
  const merged = [...dstItems, ...toMove.filter((i) => !dstIds.has(i.id))];
  setSrc(remain);
  setDst(merged);
  clearSel();
}

function ListBox({
  title,
  list,
  selected,
  onToggle,
  onSelectAll,
  onClear,
  loading, // ✅ 추가
}: {
  title: string;
  hint?: string;
  list: ClassItem[];
  selected: Set<string>;
  onToggle: (id: string) => void;
  onSelectAll: () => void;
  onClear: () => void;
  loading?: boolean; // ✅ 추가
}) {
  const allSelected = list.length > 0 && selected.size === list.length;

  return (
    <Card className="flex h-[495px] w-full flex-col py-1 gap-1">
      <CardHeader className="space-y-0 pb-0 pt-2 justify-center">
        <CardTitle className="text-base font-semibold pb-0">{title}</CardTitle>
      </CardHeader>

      <CardContent className="flex min-h-0 flex-1 flex-col px-1">
        <div className="flex items-center justify-end">
          <Toggle
            pressed={allSelected}
            onPressedChange={(pressed) => (pressed ? onSelectAll() : onClear())}
            aria-label="전체 선택 토글"
            className="h-8 px-3"
            disabled={loading} // ✅ 로딩 중 비활성화
          >
            {allSelected ? "전체 해제" : "전체 선택"}
          </Toggle>
        </div>

        <div className="min-h-0 flex-1 rounded-lg border">
          <ScrollArea className="h-full w-full p-1">
            {loading ? (
              // ✅ 로딩 상태 UI
              <div className="flex h-full w-full items-center justify-center gap-2 text-sm text-muted-foreground">
                <Spinner />
                불러오는 중…
              </div>
            ) : (
              <ul className="space-y-1 w-full min-w-0">
                {list.length === 0 && (
                  <li className="p-3 text-sm text-muted-foreground">항목이 없습니다.</li>
                )}
                {list.map((item) => {
                  const isSel = selected.has(item.id);
                  return (
                    <li className="pr-2" key={item.id}>
                      <button
                        type="button"
                        onClick={() => onToggle(item.id)}
                        disabled={loading} // ✅ 로딩 중 선택 비활성화
                        className={`group flex w-full items-start gap-2 rounded-md border px-2 py-1.5 text-left text-sm transition ${
                          isSel ? "bg-blue-50 border-blue-200" : "hover:bg-accent border-transparent"
                        } ${loading ? "opacity-60 cursor-not-allowed" : ""}`}
                      >
                        {/* 긴 이름도 영역 안 줄바꿈 */}
                        <span className="flex-1 min-w-0 break-all text-xs leading-5">{item.name}</span>
                        <span className="h-4 w-4 shrink-0 grid place-items-center">
                          {isSel ? <Check className="h-4 w-4 text-blue-600" /> : null}
                        </span>
                      </button>
                    </li>
                  );
                })}
              </ul>
            )}
          </ScrollArea>
        </div>
      </CardContent>
    </Card>
  );
}

export default function UpdateClassView({ meta }: ViewProps) {
  const left = useSelectableList([] as ClassItem[]); // 추가되지 않은 반
  const center = useSelectableList([] as ClassItem[]); // 현행(유지)
  const right = useSelectableList([] as ClassItem[]); // 삭제할 반

  const [loading, setLoading] = useState<boolean>(false); // ✅ 로딩 상태

  // 데이터 로드
  useEffect(() => {
    (async () => {
      try {
        setLoading(true); // ✅ 시작
        // 1) 파일 기준 현행 반 목록
        const dfRes = await rpc.call("get_datafile_data", {}); // [class_student_dict, class_test_dict]
        let classStudentDict: Record<string, unknown> = {};
        if (Array.isArray(dfRes) && typeof dfRes[0] === "object") {
          classStudentDict = dfRes[0] as Record<string, unknown>;
        } else if (dfRes?.class_student_dict) {
          classStudentDict = dfRes.class_student_dict as Record<string, unknown>;
        }
        const currentNames = Object.keys(classStudentDict).sort();

        // 2) 아이소식 기준인데 파일에는 없는 반 목록
        const aisosicRes = await rpc.call("get_aisosic_data", {}); // string[]
        const missingNames: string[] = Array.isArray(aisosicRes) ? aisosicRes : [];

        const currentSet = new Set(currentNames);
        const leftOnly = missingNames.filter((n) => !currentSet.has(n));

        center.setItems(currentNames.map((n) => ({ id: n, name: n })));
        left.setItems(leftOnly.map((n) => ({ id: n, name: n })));
        right.setItems([]);
        left.clearSelection();
        center.clearSelection();
        right.clearSelection();
      } catch (e) {
        left.setItems([]);
        center.setItems([]);
        right.setItems([]);
        left.clearSelection();
        center.clearSelection();
        right.clearSelection();
        console.error(e);
      } finally {
        setLoading(false); // ✅ 종료
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleSave = () => {
    console.log("현행(유지)", center.items);
    console.log("삭제할", right.items);
    console.log("추가되지 않은", left.items);
  };

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="mt-1 text-sm text-muted-foreground">{meta.guide}</p>
        </div>
        <Separator className="mb-4" />

        <div className="grid grid-cols-14 gap-1 h-full">
          {/* Left List */}
          <div className="col-span-4 h-full">
            <ListBox
              title="추가되지 않은 반"
              list={left.items}
              selected={left.selected}
              onToggle={left.toggle}
              onSelectAll={left.selectAll}
              onClear={left.clearSelection}
              loading={loading} // ✅ 전달
            />
          </div>

          {/* Controls between Left and Center */}
          <div className="col-span-1 flex flex-col items-center justify-center gap-2">
            <Button
              variant="outline"
              onClick={() =>
                move(left.items, left.setItems, center.items, center.setItems, left.selected, left.clearSelection)
              }
              className="w-10"
              title="선택 → 가운데"
              disabled={loading} // ✅ 로딩 중 비활성화
            >
              <ArrowRight className="h-4 w-4" />
            </Button>
            <Button
              variant="outline"
              onClick={() =>
                move(center.items, center.setItems, left.items, left.setItems, center.selected, center.clearSelection)
              }
              className="w-10"
              title="선택 ← 왼쪽"
              disabled={loading}
            >
              <ArrowLeft className="h-4 w-4" />
            </Button>
          </div>

          {/* Center List */}
          <div className="col-span-4 h-full">
            <ListBox
              title="현행(유지할) 반"
              list={center.items}
              selected={center.selected}
              onToggle={center.toggle}
              onSelectAll={center.selectAll}
              onClear={center.clearSelection}
              loading={loading}
            />
          </div>

          {/* Controls between Center and Right */}
          <div className="col-span-1 flex flex-col items-center justify-center gap-2">
            <Button
              variant="outline"
              onClick={() =>
                move(center.items, center.setItems, right.items, right.setItems, center.selected, center.clearSelection)
              }
              className="w-10"
              title="선택 → 오른쪽"
              disabled={loading}
            >
              <ArrowRight className="h-4 w-4" />
            </Button>
            <Button
              variant="outline"
              onClick={() =>
                move(right.items, right.setItems, center.items, center.setItems, right.selected, right.clearSelection)
              }
              className="w-10"
              title="선택 ← 가운데"
              disabled={loading}
            >
              <ArrowLeft className="h-4 w-4" />
            </Button>
          </div>

          {/* Right List */}
          <div className="col-span-4 h-full">
            <ListBox
              title="삭제할 반"
              list={right.items}
              selected={right.selected}
              onToggle={right.toggle}
              onSelectAll={right.selectAll}
              onClear={right.clearSelection}
              loading={loading}
            />
          </div>
        </div>

        {/* 하단 액션 */}
        <div className="flex items-center justify-end">
          <Button className="rounded-xl bg-black text-white" onClick={handleSave} disabled={loading}>
            {loading ? <><Spinner /> 불러오는 중…</> : "반 정보 수정"}
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}
