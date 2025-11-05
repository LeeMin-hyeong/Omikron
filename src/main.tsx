import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import '@/index.css';
import OmikronPanel from '@/App.tsx';
import { PrereqProvider } from '@/contexts/prereq.tsx';
import { AppDialogProvider } from "@/components/app-dialog/AppDialogProvider";
import { HolidayDialogProvider } from "@/components/holiday-dialog/useHolidayDialog";

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <PrereqProvider>
      <AppDialogProvider>
        <HolidayDialogProvider>
          <OmikronPanel />
        </HolidayDialogProvider>
      </AppDialogProvider>
    </PrereqProvider>
  </StrictMode>
);
