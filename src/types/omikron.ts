export type OmikronActionKey =
  | "update-class"
  | "rename-data-file"
  | "update-students"
  | "generate-daily-form"
  | "save-exam"
  | "send-exam-message"
  | "save-individual-exam"
  | "save-retest"
  | "reapply-conditional-format"
  | "add-student"
  | "remove-student"
  | "move-student";

export interface ActionMeta {
  title: string;
  guide: string;
  steps: string[];
}

export interface ViewProps {
  meta: ActionMeta;
  onAction?: (key: OmikronActionKey) => void;
}
