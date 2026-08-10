import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import '@/index.css';
import TdmPanel from '@/app/App.tsx';
import { PrereqProvider } from '@/app/PrereqProvider';
import { AppDialogProvider } from "@/shared/components/dialogs/app/AppDialogProvider";
import { HolidayDialogProvider } from "@/shared/components/dialogs/holiday/useHolidayDialog";

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <PrereqProvider>
      <AppDialogProvider>
        <HolidayDialogProvider>
          <TdmPanel />
        </HolidayDialogProvider>
      </AppDialogProvider>
    </PrereqProvider>
  </StrictMode>
);
