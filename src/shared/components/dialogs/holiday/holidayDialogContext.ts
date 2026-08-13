import { createContext } from "react";
import type { WeekdayKRMap } from "./HolidayDialog";

export type HolidayDialogOptions = {
  title?: string;
  confirmText?: string;
  baseDate?: Date;
};

export type HolidayDialogContextValue = {
  openHolidayDialog: (opts?: HolidayDialogOptions) => Promise<WeekdayKRMap | null>;
  lastHolidaySelection: WeekdayKRMap | null;
  clearHolidaySelection: () => void;
};

export const HolidayDialogContext = createContext<HolidayDialogContextValue | null>(null);
