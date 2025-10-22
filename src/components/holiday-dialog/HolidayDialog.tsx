// src/components/holiday-dialog/HolidayDialog.tsx
import * as React from "react";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";

export type WeekdayKr = "월" | "화" | "수" | "목" | "금" | "토" | "일";
export type WeekdayKRMap = Record<WeekdayKr, string>; // 'YYYY-MM-DD'

const WEEKDAYS: WeekdayKr[] = ["월", "화", "수", "목", "금", "토", "일"];
const toMondayZero = (jsDay: number) => (jsDay + 6) % 7; // 월=0 … 일=6

const fmt = (d: Date) => {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${dd}`;
};

function nextDateForWeekday(from: Date, weekdayMonZero: number) {
  const d = new Date(from); d.setHours(0, 0, 0, 0);
  const today = toMondayZero(d.getDay());
  let delta = (weekdayMonZero - today + 7) % 7;
  if (delta === 0) delta = 7; // 오늘 요일이면 다음 주
  const r = new Date(d);
  r.setDate(d.getDate() + delta);
  return r;
}

export default function HolidayDialog({
  open,
  onOpenChange,
  onConfirm,
  title = "학원 휴일 선택",
  confirmText = "휴일 설정",
  baseDate,
}: {
  open: boolean;
  onOpenChange: (v: boolean) => void;
  onConfirm: (result: WeekdayKRMap) => void;
  title?: string;
  confirmText?: string;
  baseDate?: Date;
}) {
  const today = React.useMemo(() => {
    const d = baseDate ? new Date(baseDate) : new Date();
    d.setHours(0, 0, 0, 0);
    return d;
  }, [baseDate]);

  const initialMap = React.useMemo(() => {
    const map = {} as Record<WeekdayKr, Date>;
    WEEKDAYS.forEach((w, idx) => (map[w] = nextDateForWeekday(today, idx)));
    return map;
  }, [today]);

  const order = React.useMemo(() => {
    const start = (toMondayZero(today.getDay()) + 1) % 7; // 내일부터
    return Array.from({ length: 7 }, (_, i) => (start + i) % 7);
  }, [today]);

  const [checked, setChecked] = React.useState<Set<number>>(new Set());
  const toggle = (idx: number, v?: boolean) =>
    setChecked((prev) => {
      const n = new Set(prev);
      const on = v ?? !n.has(idx);
      on ? n.add(idx) : n.delete(idx);
      return n;
    });

  const handleConfirm = () => {
    const out = {} as WeekdayKRMap;
    WEEKDAYS.forEach((w, idx) => {
      const d = new Date(initialMap[w]);
      if (checked.has(idx)) d.setDate(d.getDate() + 7);
      out[w] = fmt(d); // ← 서버에 주기 좋은 문자열 포맷
    });
    onConfirm(out);
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-[420px]">
        <DialogHeader>
          <DialogTitle>{title}</DialogTitle>
        </DialogHeader>

        <DialogDescription>
          <div>
            <p>학원 휴일을 체크하여 해당 날짜에 예정된 <b style={{color: "red"}}>재시험</b>을 </p>
            <p><b style={{color: "red"}}>일주일(+7일) 연기</b>합니다.</p>
            <p>학생의 가능한 재시험 일정 중 가장 가까운 일정을</p>
            <p>자동으로 선택합니다.</p>
            <p>학원 휴일은 데스크에 문의해주세요</p>
          </div>
        </DialogDescription>

        <div className="grid gap-2">
          {order.map((idx) => {
            const label = WEEKDAYS[idx];
            const date = initialMap[label];
            const id = `holiday-${idx}`;
            const isChecked = checked.has(idx);
            return (
              <label
                key={id}
                htmlFor={id}
                className={`flex items-center justify-between rounded-xl border px-3 py-2 text-sm ${
                  isChecked ? "bg-blue-50 border-blue-200" : "hover:bg-accent"
                }`}
              >
                <div className="flex items-center gap-3">
                  <Checkbox id={id} checked={isChecked} onCheckedChange={(v) => toggle(idx, Boolean(v))} />
                  <span className="font-medium">{label}</span>
                </div>
                <div className="tabular-nums text-muted-foreground">{fmt(date)}</div>
              </label>
            );
          })}
        </div>

        <DialogFooter className="mt-4">
          <Button variant="outline" onClick={() => onOpenChange(false)}>취소</Button>
          <Button onClick={handleConfirm}>{confirmText}</Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
