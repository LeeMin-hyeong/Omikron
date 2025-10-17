import type { OmikronActionKey, ViewProps } from "@/types/omikron";
import type { FC } from "react";
import RenameDataFileView from "./RenameDataFileView";
import SaveExamView from "./SaveExamView";
import GenericComingSoonView from "./GenericComingSoonView";
import GenerateDailyFormView from "./GenerateDailyFromView";
import SendExamMessageView from "./SendExamMessageView";
import ReapplyConditionalFormatView from "./ReapplyConditionalFromatView";
import MoveStudentView from "./MoveStudentView";
import AddStudentView from "./AddStudentView";
import SaveIndividualExamView from "./SaveIndividualExamView";
import SaveRetestView from "./SaveRetestView";
import RemoveStudentView from "./RemoveStudentView";
import UpdateClassView from "./UpdateClassView";
import UpdateStudentView from "./UpdateStudentView";

const viewMap: Partial<Record<OmikronActionKey, FC<ViewProps>>> = {
  "rename-data-file": RenameDataFileView,
  "save-exam": SaveExamView,
  "generate-daily-form": GenerateDailyFormView,
  "send-exam-message": SendExamMessageView,
  "reapply-conditional-format": ReapplyConditionalFormatView,
  "move-student": MoveStudentView,
  "add-student": AddStudentView,
  "save-individual-exam": SaveIndividualExamView,
  "save-retest": SaveRetestView,
  "remove-student": RemoveStudentView,
  "update-class": UpdateClassView,
  "update-students": UpdateStudentView,
};

export function getActionView(action: OmikronActionKey): FC<ViewProps> {
  return viewMap[action] ?? GenericComingSoonView;
}
