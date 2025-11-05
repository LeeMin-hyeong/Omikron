// src/components/holiday-dialog/useHolidayDialog.tsx
import {
  createContext,
  useCallback,
  useContext,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import HolidayDialog, { type WeekdayKRMap } from "./HolidayDialog";

type HolidayDialogOptions = {
  title?: string;
  confirmText?: string;
  baseDate?: Date;
};

type HolidayDialogContextValue = {
  openHolidayDialog: (opts?: HolidayDialogOptions) => Promise<WeekdayKRMap | null>;
  lastHolidaySelection: WeekdayKRMap | null;
  clearHolidaySelection: () => void;
};

const HolidayDialogContext = createContext<HolidayDialogContextValue | null>(null);

export function HolidayDialogProvider({ children }: { children: ReactNode }) {
  const resolverRef = useRef<((value: WeekdayKRMap | null) => void) | null>(null);
  const [open, setOpen] = useState(false);
  const [dialogOptions, setDialogOptions] = useState<HolidayDialogOptions | undefined>(undefined);
  const [selection, setSelection] = useState<WeekdayKRMap | null>(null);

  const settle = useCallback((result: WeekdayKRMap | null) => {
    const resolver = resolverRef.current;
    resolverRef.current = null;
    if (resolver) resolver(result);
    setOpen(false);
    setDialogOptions(undefined);
  }, []);

  const handleConfirm = useCallback(
    (map: WeekdayKRMap) => {
      setSelection(map);
      settle(map);
    },
    [settle]
  );

  const handleOpenChange = useCallback(
    (nextOpen: boolean) => {
      if (!nextOpen && resolverRef.current) {
        settle(null);
      }
    },
    [settle]
  );

  const openHolidayDialog = useCallback(
    (opts?: HolidayDialogOptions) => {
      return new Promise<WeekdayKRMap | null>((resolve) => {
        resolverRef.current = resolve;
        setDialogOptions(opts);
        setOpen(true);
      });
    },
    []
  );

  const clearHolidaySelection = useCallback(() => setSelection(null), []);

  const value = useMemo<HolidayDialogContextValue>(
    () => ({
      openHolidayDialog,
      lastHolidaySelection: selection,
      clearHolidaySelection,
    }),
    [clearHolidaySelection, openHolidayDialog, selection]
  );

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

export function useHolidayDialog() {
  const ctx = useContext(HolidayDialogContext);
  if (!ctx) throw new Error("HolidayDialogProvider is missing in the tree.");
  return ctx;
}

export default useHolidayDialog;
