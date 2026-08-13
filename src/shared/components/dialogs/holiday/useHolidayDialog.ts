import { useContext } from "react";
import { HolidayDialogContext } from "./holidayDialogContext";

export default function useHolidayDialog() {
  const context = useContext(HolidayDialogContext);
  if (!context) throw new Error("HolidayDialogProvider is missing in the tree.");
  return context;
}
