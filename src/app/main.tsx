import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import '@/index.css';
import TdmPanel from '@/app/App.tsx';
import { PrereqProvider } from '@/app/PrereqProvider';
import { AppDialogProvider } from "@/shared/components/dialogs/app/AppDialogProvider";
import { HolidayDialogProvider } from "@/shared/components/dialogs/holiday/HolidayDialogProvider";
import { JobProvider } from "@/app/JobProvider";
import { JobNotificationProvider } from "@/app/JobNotificationProvider";

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <PrereqProvider>
      <AppDialogProvider>
        <JobProvider>
          <JobNotificationProvider>
            <HolidayDialogProvider>
              <TdmPanel />
            </HolidayDialogProvider>
          </JobNotificationProvider>
        </JobProvider>
      </AppDialogProvider>
    </PrereqProvider>
  </StrictMode>
);
