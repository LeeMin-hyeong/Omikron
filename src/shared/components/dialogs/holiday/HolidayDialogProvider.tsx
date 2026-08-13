import {
  useCallback,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import HolidayDialog, { type WeekdayKRMap } from "./HolidayDialog";
import {
  HolidayDialogContext,
  type HolidayDialogContextValue,
  type HolidayDialogOptions,
} from "./holidayDialogContext";

export function HolidayDialogProvider({ children }: { children: ReactNode }) {
  const resolverRef = useRef<((value: WeekdayKRMap | null) => void) | null>(null);
  const [open, setOpen] = useState(false);
  const [dialogOptions, setDialogOptions] = useState<HolidayDialogOptions>();
  const [selection, setSelection] = useState<WeekdayKRMap | null>(null);

  const settle = useCallback((result: WeekdayKRMap | null) => {
    const resolver = resolverRef.current;
    resolverRef.current = null;
    resolver?.(result);
    setOpen(false);
    setDialogOptions(undefined);
  }, []);

  const handleConfirm = useCallback((map: WeekdayKRMap) => {
    setSelection(map);
    settle(map);
  }, [settle]);

  const handleOpenChange = useCallback((nextOpen: boolean) => {
    if (!nextOpen && resolverRef.current) settle(null);
  }, [settle]);

  const openHolidayDialog = useCallback((opts?: HolidayDialogOptions) => (
    new Promise<WeekdayKRMap | null>((resolve) => {
      resolverRef.current = resolve;
      setDialogOptions(opts);
      setOpen(true);
    })
  ), []);

  const clearHolidaySelection = useCallback(() => setSelection(null), []);
  const value = useMemo<HolidayDialogContextValue>(() => ({
    openHolidayDialog,
    lastHolidaySelection: selection,
    clearHolidaySelection,
  }), [clearHolidaySelection, openHolidayDialog, selection]);

  return (
    <HolidayDialogContext.Provider value={value}>
      {children}
      <HolidayDialog
        open={open}
        onOpenChange={handleOpenChange}
        onConfirm={handleConfirm}
        title={dialogOptions?.title}
        confirmText={dialogOptions?.confirmText}
        baseDate={dialogOptions?.baseDate}
      />
    </HolidayDialogContext.Provider>
  );
}
