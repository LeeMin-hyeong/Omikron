import { rpc } from "pyloid-js";
import type { ProgressPayload } from "@/api/progress";

export const RPC_ERROR_CODES = [
  "AISOSIK_STUDENT_NOT_FOUND", "CANCELLED", "CHROME_DRIVER_MISMATCH",
  "CLASS_INFO_MISSING", "DAILY_MESSAGE_REQUIRED", "DATA_DIR_INVALID",
  "DATA_DIR_REQUIRED", "DATA_FILE_NAME_MISSING", "DATA_FILE_NAME_REQUIRED",
  "DATA_FORM_INVALID", "ENVIRONMENT_IO", "EXCEL_REQUIRED", "FILE_ALREADY_EXISTS",
  "FILE_NOT_FOUND", "FILE_OPEN", "INTERNAL_ERROR", "INVALID_INPUT",
  "INVALID_OPERATION", "INVALID_REQUEST", "JOB_ALREADY_RUNNING", "JOB_NOT_CANCELLABLE",
  "JOB_NOT_FOUND", "MAKEUP_DATE_MESSAGE_REQUIRED", "MAKEUP_MESSAGE_REQUIRED",
  "NOTICE_ID_REQUIRED", "NOTICE_REPOSITORY_INVALID", "STORAGE_PERMISSION_DENIED",
  "STORAGE_UNAVAILABLE", "TDM_ERROR", "UNSUPPORTED_FILE_TYPE", "URL_INVALID",
  "URL_REQUIRED", "WORKBOOK_BUSY", "WORKBOOK_COLUMN_MISSING", "WORKBOOK_CONFLICT",
  "WORKBOOK_INVALID", "WORKBOOK_RECOVERY_REQUIRED", "WORKBOOK_SHEET_MISSING",
  "WORKBOOK_TRANSACTION",
] as const;
export type RpcErrorCode = typeof RPC_ERROR_CODES[number];
export type RpcFailure = { ok: false; code: RpcErrorCode; error: string; detail?: string };
export type RpcResponse<T> = { ok: true; data: T } | RpcFailure;
type EmptyParams = Record<string, never>;
type EmptyData = Record<string, never>;
type StudentMap = Record<string, Record<string, number>>;
type TestMap = Record<string, Record<string, number>>;
type DatafileData = [StudentMap, TestMap] | { class_student_dict: StudentMap; class_test_dict: TestMap };
type AisosikStudentMap = Record<string, string[]>;

type ConfigData = {
  exists: boolean;
  ready: boolean;
  termsAccepted: boolean;
  config: {
    url: string;
    dataDir: string;
    dataFileName: string;
    dailyTest: string;
    makeupTest: string;
    makeupTestDate: string;
  };
};

type GeneralRpcContract = {
  check_data_files: { params: EmptyParams; data: { ok: boolean; has_class: boolean; has_data: boolean; has_student: boolean; data_dir_valid: boolean; data_file_name: string; cwd: string; data_dir: string; missing: string[]; recovery_actions: unknown[]; storage_error: string } };
  verify_storage_health: { params: EmptyParams; data: { recovery_actions: unknown[] } };
  get_config_status: { params: EmptyParams; data: ConfigData };
  select_data_dir: { params: EmptyParams; data: { path: string } };
  save_initial_config: { params: { url: string; data_dir: string; data_file_name: string; daily_test_message: string; makeup_test_message: string; makeup_test_date_message: string }; data: EmptyData };
  update_message_templates: { params: { url: string; daily_test_message: string; makeup_test_message: string; makeup_test_date_message: string }; data: EmptyData };
  validate_script_url: { params: { url: string }; data: { warning: boolean } };
  get_terms_text: { params: EmptyParams; data: { title: string; text: string } };
  accept_terms: { params: EmptyParams; data: EmptyData };
  get_startup_notice: { params: EmptyParams; data: { enabled: boolean; title: string; message: string; noticeId?: string } };
  mark_notice_seen: { params: { notice_id: string }; data: EmptyData };
  quit_app: { params: EmptyParams; data: EmptyData };
  confirm_app_exit: { params: EmptyParams; data: EmptyData };
  get_startup_messages: { params: EmptyParams; data: { termsTitle: string; termsMessage: string; noticeTitle: string; noticeMessage: string } };
  change_data_dir: { params: EmptyParams; data: EmptyData };
  change_data_file_name: { params: { new_filename: string }; data: EmptyData };
  open_path: { params: { path: string }; data: EmptyData };
  open_url: { params: { url: string }; data: EmptyData };
  make_class_info: { params: EmptyParams; data: { path: string } };
  make_data_file: { params: EmptyParams; data: EmptyData };
  make_student_info: { params: EmptyParams; data: { path: string } };
  make_data_form: { params: EmptyParams; data: EmptyData };
  reapply_conditional_format: { params: EmptyParams; data: { warnings: string[] } };
  update_student_info: { params: EmptyParams; data: EmptyData };
  add_student: { params: { target_student_name: string; target_class_name: string }; data: { warnings: string[] } };
  remove_student: { params: { target_class_name: string; target_student_name: string }; data: EmptyData };
  move_student: { params: { target_student_name: string; target_class_name: string; current_class_name: string }; data: { warnings: string[] } };
  change_class_info: { params: { target_class_name: string; target_teacher_name: string }; data: EmptyData };
  make_temp_class_info: { params: { new_class_list: string[] }; data: { path: string } };
  update_class: { params: EmptyParams; data: EmptyData };
  delete_class_info_temp: { params: EmptyParams; data: EmptyData };
  save_individual_result: { params: { student_name: string; class_name: string; test_name: string; target_row: number; target_col: number; test_score: number; makeup_test_check: boolean; makeup_test_date: Record<string, string> }; data: { warnings: string[] } };
  save_retest_result: { params: { target_row: number; makeup_test_score: string }; data: EmptyData };
  change_data_file_name_by_select: { params: EmptyParams; data: EmptyData };
  open_file_picker: { params: EmptyParams; data: { path: string; name: string; b64: string } };
  get_datafile_data: { params: { mocktest?: boolean }; data: DatafileData };
  get_aisosik_data: { params: EmptyParams; data: string[] };
  get_aisosik_student_data: { params: EmptyParams; data: AisosikStudentMap };
  check_aisosik_difference: { params: EmptyParams; data: boolean };
  get_makeuptest_data: { params: EmptyParams; data: Record<string, Record<string, number>> };
  get_class_list: { params: EmptyParams; data: string[] };
  get_class_info: { params: { class_name: string }; data: string };
  get_new_class_list: { params: EmptyParams; data: string[] };
  is_cell_empty: { params: { row: number; col: number }; data: { empty: boolean; value: unknown } };
};

export type GeneralRpcMethod = keyof GeneralRpcContract;
export const GENERAL_RPC_METHODS: GeneralRpcMethod[] = [
  "check_data_files", "verify_storage_health", "get_config_status", "select_data_dir",
  "save_initial_config", "update_message_templates", "validate_script_url", "get_terms_text",
  "accept_terms", "get_startup_notice", "mark_notice_seen", "quit_app", "confirm_app_exit",
  "get_startup_messages", "change_data_dir", "change_data_file_name", "open_path", "open_url",
  "make_class_info", "make_data_file", "make_student_info", "make_data_form",
  "reapply_conditional_format", "update_student_info", "add_student", "remove_student",
  "move_student", "change_class_info", "make_temp_class_info", "update_class", "delete_class_info_temp",
  "save_individual_result", "save_retest_result", "change_data_file_name_by_select",
  "open_file_picker", "get_datafile_data", "get_aisosik_data", "get_aisosik_student_data",
  "check_aisosik_difference", "get_makeuptest_data", "get_class_list", "get_class_info",
  "get_new_class_list", "is_cell_empty",
];

type DataPreservingMethod =
  | "get_datafile_data"
  | "get_aisosik_data"
  | "get_aisosik_student_data"
  | "check_aisosik_difference"
  | "get_makeuptest_data"
  | "get_class_list"
  | "get_class_info"
  | "get_new_class_list";

const DATA_PRESERVING_METHODS = new Set<GeneralRpcMethod>([
  "get_datafile_data", "get_aisosik_data", "get_aisosik_student_data",
  "check_aisosik_difference", "get_makeuptest_data", "get_class_list",
  "get_class_info", "get_new_class_list",
]);

type GeneralRpcResult<M extends GeneralRpcMethod> = M extends "check_data_files"
  ? GeneralRpcContract[M]["data"]
  : M extends DataPreservingMethod
    ? RpcResponse<GeneralRpcContract[M]["data"]>
    : RpcFailure | ({ ok: true } & GeneralRpcContract[M]["data"]);

async function callGeneralRpc<M extends GeneralRpcMethod>(
  method: M,
  params: GeneralRpcContract[M]["params"],
): Promise<GeneralRpcResult<M>> {
  const response = await rpc.call(method, params) as RpcFailure | ({ ok: true; data?: unknown } & Record<string, unknown>);
  if (!response.ok) return response as GeneralRpcResult<M>;
  if (DATA_PRESERVING_METHODS.has(method)) return response as GeneralRpcResult<M>;
  if (response.data !== undefined) {
    if (typeof response.data === "object" && response.data !== null && !Array.isArray(response.data)) {
      return { ok: true, ...response.data } as GeneralRpcResult<M>;
    }
    return { ok: true, data: response.data } as GeneralRpcResult<M>;
  }
  return response as GeneralRpcResult<M>;
}

export const generalRpc = { call: callGeneralRpc };
export type JobType = "start_save_exam" | "start_send_exam_message" | "start_update_class";
export type JobStartData = { jobId: string; jobType: JobType; status: "running" };
export type JobUploadRequest = {
  filename: string;
  b64: string;
  makeup_test_date: Record<string, string>;
};
export type JobState = ProgressPayload & {
  jobId: string;
  jobType: JobType;
  createdAt: number;
  updatedAt: number;
  finishedAt?: number | null;
  cancellationRequested: boolean;
};
export type JobBatchItem = {
  jobId: string;
  revision: number;
  changed: boolean;
  state?: Record<string, unknown>;
};

async function callJobRpc<T>(method: string, params: Record<string, unknown>): Promise<T> {
  return rpc.call(method, params) as Promise<T>;
}

export const jobRpc = {
  start(method: JobType, params: JobUploadRequest | Record<string, never>) {
    return callJobRpc<RpcResponse<JobStartData>>(method, params);
  },
  getBatch(jobs: Array<{ jobId: string; revision: number }>) {
    return callJobRpc<RpcResponse<{ jobs: JobBatchItem[] }>>(
      "get_job_progress_batch",
      { jobs },
    );
  },
  get(jobId: string) {
    return callJobRpc<RpcResponse<JobState>>("get_job", { job_id: jobId });
  },
  acknowledge(jobId: string) {
    return callJobRpc<RpcResponse<{ jobId: string; acknowledged: true }>>(
      "acknowledge_job_completion",
      { job_id: jobId },
    );
  },
  cancel(jobId: string) {
    return callJobRpc<RpcResponse<{ jobId: string; cancellationRequested: true }>>(
      "cancel_job",
      { job_id: jobId },
    );
  },
};

export async function fileToBase64(file: File): Promise<string> {
  const buf = await file.arrayBuffer()
  // 빠른 base64 인코딩
  let binary = ""
  const bytes = new Uint8Array(buf)
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk))
  }
  return btoa(binary)
}
