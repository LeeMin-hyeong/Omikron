export type OmikronActionKey =
  | "welcome"
  | "update-class"
  | "rename-data-file"
  | "update-students"
  | "update-teacher"
  | "generate-daily-form"
  | "save-exam"
  | "send-exam-message"
  | "save-individual-exam"
  | "save-retest"
  | "reapply-conditional-format"
  | "manage-student";

export interface ActionMeta {
  title: string;
  guide: string;
  steps: string[];
}

export interface ViewProps {
  meta: ActionMeta;
  onAction?: (key: OmikronActionKey) => void;
}
