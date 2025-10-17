import React, { useMemo, useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { ScrollArea } from "@/components/ui/scroll-area";
import { ArrowLeft, ArrowRight, Check, X } from "lucide-react";
import { ViewProps } from "@/types/omikron";
import { Separator } from "@/components/ui/separator";
import { Toggle } from "@/components/ui/toggle";



export type ClassItem = {
  id: string;
  name: string;
};

function useSelectableList(initial: ClassItem[]) {
  const [items, setItems] = useState<ClassItem[]>(initial);
  const [selected, setSelected] = useState<Set<string>>(new Set());

  const toggle = (id: string) => {
    setSelected((prev) => {
      const n = new Set(prev);
      if (n.has(id)) n.delete(id);
      else n.add(id);
      return n;
    });
  };

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
  if (toMove.length === 0) return;
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
}: {
  title: string;
  hint?: string;
  list: ClassItem[];
  selected: Set<string>;
  onToggle: (id: string) => void;
  onSelectAll: () => void;
  onClear: () => void;
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
            onPressedChange={(pressed) => {
              if (pressed) onSelectAll();
              else onClear();
            }}
            aria-label="전체 선택 토글"
            className="h-8 px-3"
          >
            { allSelected ? "전체 해제" : "전체 선택" }
          </Toggle>
        </div>

        <div className="min-h-0 flex-1 rounded-lg border">
          <ScrollArea className="h-full w-full p-1">
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
                      className={`group flex w-full items-start gap-2 rounded-md border px-2 py-1.5 text-left text-sm transition ${
                        isSel
                          ? "bg-blue-50 border-blue-200"
                          : "hover:bg-accent border-transparent"
                      }`}
                    >
                      <span className="flex-1 min-w-0 break-normal text-xs leading-5">
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
  // 더미 데이터
  const initialLeft = useMemo<ClassItem[]>(
    () => [
      { id: "c1", name: "특강)중2 서술형특강2차_토요일(9/27,11,18)" },
      { id: "c2", name: "가나다라마바사아자차카타파하가나다라마바사아자차카" },
      { id: "c3", name: "3반" },
      { id: "c4", name: "4반" },
      { id: "c5", name: "5반" },
      { id: "c6", name: "5반" },
      { id: "c7", name: "5반" },
      { id: "c8", name: "5반" },
      { id: "c9", name: "5반" },
      { id: "c10", name: "5반" },
      { id: "c11", name: "5반" },
      { id: "c12", name: "5반" },
      { id: "c13", name: "5반" },
      { id: "c14", name: "5반" },
      { id: "c15", name: "5반" },
      { id: "c16", name: "5반" },
      { id: "c17", name: "5반" },
      { id: "c18", name: "5반" },
      { id: "c19", name: "5반" },
    ],
    [],
  );

  const left = useSelectableList(initialLeft);
  const center = useSelectableList([] as ClassItem[]);
  const right = useSelectableList([] as ClassItem[]);

  // 저장(우측 하단 버튼)
  const handleSave = () => {
    // 실서비스: API 호출로 center.items(유지), right.items(삭제) 전송
    console.log("현행(유지)", center.items);
    console.log("삭제할", right.items);
    console.log("추가되지 않은", left.items);
    // toast/알림을 사용해도 좋습니다.
  };

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="mt-1 text-sm text-muted-foreground">
            {meta.guide}
          </p>
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
          />
        </div>

        {/* Controls between Left and Center */}
        <div className="col-span-1 flex flex-col items-center justify-center gap-2">
          <Button
            variant="outline"
            onClick={() => move(left.items, left.setItems, center.items, center.setItems, left.selected, left.clearSelection)}
            className="w-10"
            title="선택 → 가운데"
          >
            <ArrowRight className="h-4 w-4" />
          </Button>
          <Button
            variant="outline"
            onClick={() => move(center.items, center.setItems, left.items, left.setItems, center.selected, center.clearSelection)}
            className="w-10"
            title="선택 ← 왼쪽"
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
          />
        </div>

        {/* Controls between Center and Right */}
        <div className="col-span-1 flex flex-col items-center justify-center gap-2">
          <Button
            variant="outline"
            onClick={() => move(center.items, center.setItems, right.items, right.setItems, center.selected, center.clearSelection)}
            className="w-10"
            title="선택 → 오른쪽"
          >
            <ArrowRight className="h-4 w-4" />
          </Button>
          <Button
            variant="outline"
            onClick={() => move(right.items, right.setItems, center.items, center.setItems, right.selected, right.clearSelection)}
            className="w-10"
            title="선택 ← 가운데"
          >
            <ArrowLeft className="h-4 w-4" />
          </Button>
        </div>

        {/* Right List */}
        <div className="col-span-4 h-full">
          <ListBox
            title="삭제할 반"
            hint="잘못 넣었으면 가운데로 되돌리세요"
            list={right.items}
            selected={right.selected}
            onToggle={right.toggle}
            onSelectAll={right.selectAll}
            onClear={right.clearSelection}
          />
        </div>
      </div>

      {/* 하단 고정 액션 영역 */}
      <div className="flex items-center justify-end">
        <Button
          className="rounded-xl bg-black text-white"
          onClick={handleSave}
        >
          반 정보 수정
        </Button>
      </div>
      </CardContent>
    </Card>
  );
}
