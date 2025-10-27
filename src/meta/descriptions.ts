import type { ActionMeta, OmikronActionKey } from "@/types/omikron";

export const descriptions: Record<OmikronActionKey, ActionMeta> = {
  "welcome": {
    title: "환영합니다",
    guide: "",
    steps: []
  },
  "update-class": {
    title: "반 업데이트",
    guide: "아이소식의 반 정보를 불러와 반 목록을 수정합니다. 수정된 반 정보 파일에서 반의 상세 정보를 작성해 주세요.",
    steps: ["반/학기 파라미터 확인", "기존 명부 백업", "신규 명부 반영", "요약 리포트 생성"],
  },
  "rename-data-file": {
    title: "데이터 파일 이름 변경",
    guide: "엑셀 데이터 파일의 버전 규칙(예: YYYYMMDD-반이름)에 맞춰 안전하게 이름을 변경합니다.",
    steps: ["현재 경로 스캔", "충돌(동일 파일명) 검사", "새 이름 규칙 적용", "결과 확인"],
  },
  "update-students": {
    title: "학생 정보 업데이트",
    guide: "아이소식에 등록된 학생 기준으로 '학생 정보.xlsx'를 최신화합니다.",
    steps: ["학생 이력 로드", "변경사항 계산", "엑셀 반영", "로그 저장"],
  },
  "update-teacher": {
    title: "담당 선생님 변경",
    guide: "특정 반의 담당 선생님을 변경합니다. 반 정보 파일과 데이터 파일에 변경 사항이 반영됩니다.",
    steps: [],
  },
  "generate-daily-form": {
    title: "데일리 테스트 기록 양식 생성",
    guide: "시험 기록 템플릿 시트를 생성하고 기본 항목을 채웁니다.",
    steps: ["템플릿 로드", "시트 생성", "헤더/서식 적용", "파일 저장"],
  },
  "save-exam": {
    title: "시험 결과 저장",
    guide: "시험의 전체 결과를 데이터 파일에 저장합니다.",
    steps: ["데이터 입력 양식 유효성 검사", "백업 생성", "데이터 파일 입력", "데이터 파일 조건부 서식 적용", "재시험 명단 입력", "파일 저장 완료"],
  },
  "send-exam-message": {
    title: "시험 결과 메시지 전송",
    guide: "저장된 결과를 기반으로 학부모 메시지를 생성하고 전송합니다.",
    steps: ["데이터 입력 양식 유효성 검사", "메시지 작성", "작성 완료"],
  },
  "save-individual-exam": {
    title: "개별 시험 결과 저장",
    guide: "특정 학생이 응시하지 않았던 시험의 결과를 저장합니다.",
    steps: ["학생 선택", "개별 스코어 계산", "엑셀 반영", "로그 기록"],
  },
  "save-retest": {
    title: "재시험 결과 저장",
    guide: "재시험 명단에 작성된 학생의 재시험 결과를 저장합니다.",
    steps: ["재시험 스코어 로드", "원점수 대비 비교", "최종 점수 산출", "머지/저장"],
  },
  "reapply-conditional-format": {
    title: "데이터 파일 조건부 서식 재지정",
    guide: "데이터 파일의 조건부 서식을 재지정합니다.",
    steps: ["시트 탐색", "서식 규칙 로드", "규칙 재적용", "검증/완료"],
  },
  "add-student": {
    title: "신규생 추가",
    guide: "신규 학생을 등록하고 초기 반/시험 설정을 적용합니다.",
    steps: ["기본 정보 입력", "반 배정", "초기 시트 생성", "확인/저장"],
  },
  "remove-student": {
    title: "퇴원 처리",
    guide: "퇴원 학생의 기록을 보관 폴더로 이동하고 현행 명부에서 제외합니다.",
    steps: ["학생 선택", "보관 이관", "명부 업데이트", "보고서 생성"],
  },
  "move-student": {
    title: "학생 반 이동",
    guide: "선택한 학생의 반을 이동하고 관련 기록을 업데이트합니다.",
    steps: ["학생/대상반 선택", "기록 이동", "참조 갱신", "완료 알림"],
  },
};
