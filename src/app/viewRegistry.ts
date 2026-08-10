import type { tdmActionKey, ViewProps } from "@/shared/types/tdm";
import type { FC } from "react";
import UpdateClassView from "@/features/classes/UpdateClassView";
import UpdateTeacherView from "@/features/classes/UpdateTeacherView";
import EditMessageConfigView from "@/features/configuration/EditMessageConfigView";
import RenameDataFileView from "@/features/data-files/RenameDataFileView";
import GenerateDailyFormView from "@/features/exams/GenerateDailyFormView";
import ReapplyConditionalFormatView from "@/features/exams/ReapplyConditionalFormatView";
import SaveExamView from "@/features/exams/SaveExamView";
import SaveIndividualExamView from "@/features/exams/SaveIndividualExamView";
import SaveRetestView from "@/features/exams/SaveRetestView";
import SendExamMessageView from "@/features/exams/SendExamMessageView";
import WelcomeView from "@/features/home/WelcomeView";
import ManageStudentView from "@/features/students/ManageStudentView";
import UpdateStudentView from "@/features/students/UpdateStudentView";

const viewMap: Partial<Record<tdmActionKey, FC<ViewProps>>> = {
  "welcome": WelcomeView,
  "rename-data-file": RenameDataFileView,
  "save-exam": SaveExamView,
  "generate-daily-form": GenerateDailyFormView,
  "send-exam-message": SendExamMessageView,
  "reapply-conditional-format": ReapplyConditionalFormatView,
  "save-individual-exam": SaveIndividualExamView,
  "save-retest": SaveRetestView,
  "update-class": UpdateClassView,
  "edit-message-config": EditMessageConfigView,
  "update-students": UpdateStudentView,
  "update-teacher": UpdateTeacherView,
  "manage-student": ManageStudentView
};

export function getActionView(action: tdmActionKey): FC<ViewProps> {
  return viewMap[action]!;
}
