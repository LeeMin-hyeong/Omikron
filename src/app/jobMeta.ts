export const jobLabels: Record<string, string> = {
  start_save_exam: "시험 결과 저장",
  start_send_exam_message: "시험 결과 메시지 작성",
  start_update_class: "반 업데이트",
};

export const jobSuccessMessages: Record<string, string> = {
  start_save_exam: "데이터 저장을 완료하였습니다.",
  start_send_exam_message: "메시지 작성이 완료되었습니다. 전송 전 내용을 확인하세요.",
  start_update_class: "반 업데이트가 완료되었습니다.",
};

export function getJobLabel(jobType: string) {
  return jobLabels[jobType] ?? "작업";
}
