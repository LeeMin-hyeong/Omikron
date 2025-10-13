import type { ActionMeta, OmikronActionKey } from "@/types/omikron";

export const descriptions: Record<OmikronActionKey, ActionMeta> = {
  "update-class": {
    title: "반 업데이트",
    guide: "선택한 학기/반 정보를 최신 데이터로 동기화합니다. 학생명부와 반 구성 변경 사항을 반영합니다.",
    steps: ["반/학기 파라미터 확인", "기존 명부 백업", "신규 명부 반영", "요약 리포트 생성"],
  },
  "rename-data-file": {
    title: "데이터 파일 이름 변경",
    guide: "엑셀 데이터 파일의 버전 규칙(예: YYYYMMDD-반이름)에 맞춰 안전하게 이름을 변경합니다.",
    steps: ["현재 경로 스캔", "충돌(동일 파일명) 검사", "새 이름 규칙 적용", "결과 확인"],
  },
  "update-students": {
    title: "학생 정보 업데이트",
    guide: "등록/퇴원/반 이동 이력을 기반으로 학생 정보를 업데이트합니다.",
    steps: ["학생 이력 로드", "변경사항 계산", "엑셀 반영", "로그 저장"],
  },
  "generate-daily-form": {
    title: "데일리 테스트 기록 양식 생성",
    guide: "시험 기록 템플릿 시트를 생성하고 기본 항목을 채웁니다.",
    steps: ["템플릿 로드", "시트 생성", "헤더/서식 적용", "파일 저장"],
  },
  "save-exam": {
    title: "시험 결과 저장",
    guide: "시험의 전체 결과를 데이터 파일에 저장합니다.",
    steps: ["시험 데이터 검증", "결측치 처리", "엑셀 쓰기", "요약/차트 갱신", "test", "test", "test", "test"],
  },
  "send-exam-message": {
    title: "시험 결과 메시지 전송",
    guide: "저장된 결과를 기반으로 학부모 메시지를 생성하고 전송합니다.",
    steps: ["메시지 템플릿 생성", "수신 대상 확인", "전송 큐 적재", "전송/재시도"],
  },
  "save-individual-exam": {
    title: "개별 시험 결과 저장",
    guide: "특정 학생의 결과만 부분 저장합니다.",
    steps: ["학생 선택", "개별 스코어 계산", "엑셀 반영", "로그 기록"],
  },
  "save-retest": {
    title: "재시험 결과 저장",
    guide: "재시험 점수를 기존 성적표에 머지합니다.",
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
