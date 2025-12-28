import type { OmikronActionKey, ViewProps } from "@/types/omikron";
import type { FC } from "react";
import RenameDataFileView from "./RenameDataFileView";
import SaveExamView from "./SaveExamView";
import GenerateDailyFormView from "./GenerateDailyFromView";
import SendExamMessageView from "./SendExamMessageView";
import ReapplyConditionalFormatView from "./ReapplyConditionalFromatView";
import SaveIndividualExamView from "./SaveIndividualExamView";
import SaveRetestView from "./SaveRetestView";
import UpdateClassView from "./UpdateClassView";
import UpdateStudentView from "./UpdateStudentView";
import UpdateTeacherView from "./UpdateTeacherView";
import WelcomeView from "./WelcomeView";
import ManageStudentView from "./ManageStudentView";

const viewMap: Partial<Record<OmikronActionKey, FC<ViewProps>>> = {
  "welcome": WelcomeView,
  "rename-data-file": RenameDataFileView,
  "save-exam": SaveExamView,
  "generate-daily-form": GenerateDailyFormView,
  "send-exam-message": SendExamMessageView,
  "reapply-conditional-format": ReapplyConditionalFormatView,
  "save-individual-exam": SaveIndividualExamView,
  "save-retest": SaveRetestView,
  "update-class": UpdateClassView,
  "update-students": UpdateStudentView,
  "update-teacher": UpdateTeacherView,
  "manage-student": ManageStudentView
};

export function getActionView(action: OmikronActionKey): FC<ViewProps> {
  return viewMap[action]!;
}
