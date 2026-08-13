import { useMemo, useState } from "react";
import { generalRpc } from "@/api/rpc";
import { Card, CardContent, CardHeader, CardTitle } from "@/shared/components/ui/card";
import { Input } from "@/shared/components/ui/input";
import { Label } from "@/shared/components/ui/label";
import { Textarea } from "@/shared/components/ui/textarea";
import { Button } from "@/shared/components/ui/button";
import { FolderOpen, HelpCircle } from "lucide-react";
import { useAppDialog } from "@/shared/components/dialogs/app/useAppDialog";

type InitialConfig = {
  url: string;
  dataDir: string;
  dataFileName: string;
  dailyTest: string;
  makeupTest: string;
  makeupTestDate: string;
};

interface Props {
  initial?: Partial<InitialConfig>;
  onComplete: () => void;
}

const STEPS: Array<{ key: keyof InitialConfig; title: string; description: string }> = [
  {
    key: "url",
    title: "학원 아이소식 스크립트 URL",
    description: "아이소식 스크립트 주소 복사 방법: 아이소식 메뉴 > 스크립트 버튼 우클릭 > 링크 주소 복사",
  },
  {
    key: "dataDir",
    title: "데이터 저장 위치",
    description: "데이터 파일을 보관할 폴더를 선택해 주세요.",
  },
  {
    key: "dataFileName",
    title: "데이터 파일 이름",
    description: "저장될 데이터 파일 이름을 입력해 주세요.",
  },
  {
    key: "dailyTest",
    title: "시험 결과 메시지 템플릿",
    description: "시험 결과 전송 시 사용하는 기본 메시지를 입력해 주세요.",
  },
  {
    key: "makeupTest",
    title: "재시험(일정 미정) 메시지 템플릿",
    description: "재시험 일정이 없는 학생에게 보낼 메시지를 입력해 주세요.",
  },
  {
    key: "makeupTestDate",
    title: "재시험(일정 안내) 메시지 템플릿",
    description: "재시험 일정을 안내할 때 사용할 메시지를 입력해 주세요.",
  },
];

export default function InitialConfigView({ initial, onComplete }: Props) {
  const dialog = useAppDialog();
  const HELP_URL = "https://tdm-db.notion.site/instruction?source=copy_link";

  const initialState = useMemo<InitialConfig>(
    () => ({
      url: initial?.url ?? "",
      dataDir: initial?.dataDir ?? "",
      dataFileName: initial?.dataFileName ?? "",
      dailyTest: initial?.dailyTest ?? "",
      makeupTest: initial?.makeupTest ?? "",
      makeupTestDate: initial?.makeupTestDate ?? "",
    }),
    [initial],
  );

  const [form, setForm] = useState<InitialConfig>(initialState);
  const [step, setStep] = useState(0);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const currentStep = STEPS[step];
  const isLastStep = step === STEPS.length - 1;

  const updateField = (key: keyof InitialConfig, value: string) => {
    setForm((prev) => ({ ...prev, [key]: value }));
    setError("");
  };

  const validateCurrentStep = () => {
    const value = form[currentStep.key];
    if (!value || !value.trim()) {
      setError("현재 항목을 입력해 주세요.");
      return false;
    }
    return true;
  };

  const pickDataDir = async () => {
    const res = await generalRpc.call("select_data_dir", {});
    if (res?.ok && res?.path) {
      updateField("dataDir", res.path);
    }
  };

  const goNext = () => {
    if (!validateCurrentStep()) return;
    setStep((prev) => Math.min(prev + 1, STEPS.length - 1));
  };

  const goPrev = () => {
    setError("");
    setStep((prev) => Math.max(prev - 1, 0));
  };

  const save = async () => {
    if (!validateCurrentStep()) return;

    const validateRes = await generalRpc.call("validate_script_url", { url: form.url });
    if (!validateRes?.ok) {
      setError(validateRes?.error ?? "URL을 검증하지 못했습니다.");
      return;
    }
    if (validateRes?.warning) {
      const proceed = await dialog.warning({
        title: "경고",
        message: "url이 정확하지 않은 것 같습니다. 계속 진행하시겠습니까?",
      });
      if (!proceed) return;
    }

    setSaving(true);
    setError("");
    try {
      const res = await generalRpc.call("save_initial_config", {
        url: form.url,
        data_dir: form.dataDir,
        data_file_name: form.dataFileName,
        daily_test_message: form.dailyTest,
        makeup_test_message: form.makeupTest,
        makeup_test_date_message: form.makeupTestDate,
      });
      if (!res?.ok) {
        setError(res?.error ?? "설정 저장에 실패했습니다.");
        return;
      }
      onComplete();
    } catch (e) {
      setError(String(e));
    } finally {
      setSaving(false);
    }
  };

  const handleOpenHelp = async () => {
    try {
      const res = await generalRpc.call("open_url", { url: HELP_URL });
      if (!res?.ok) {
        setError(res?.error ?? "사용법 페이지를 열지 못했습니다.");
      }
    } catch (e) {
      setError(String(e));
    }
  };

  return (
    <div className="flex h-full min-h-0 items-center justify-center p-4">
      <Card className="w-full max-w-3xl rounded-2xl border-border/80 shadow-sm">
        <CardHeader>
          <div className="flex items-start justify-between gap-3">
            <div>
              <CardTitle>초기 설정</CardTitle>
              <p className="text-sm text-muted-foreground">
                {step + 1} / {STEPS.length}
              </p>
            </div>
            <Button type="button" variant="outline" className="rounded-xl" onClick={handleOpenHelp}>
              <HelpCircle className="h-4 w-4" /> 사용법 및 도움말
            </Button>
          </div>
        </CardHeader>
        <CardContent className="space-y-5">
          <div className="space-y-2">
            <Label>{currentStep.title}</Label>
            <p className="text-sm text-muted-foreground">{currentStep.description}</p>

            {currentStep.key === "url" ? (
              <Input
                value={form.url}
                onChange={(e) => updateField("url", e.target.value)}
                placeholder="https://..."
              />
            ) : null}

            {currentStep.key === "dataDir" ? (
              <div className="flex gap-2">
                <Input
                  value={form.dataDir}
                  onChange={(e) => updateField("dataDir", e.target.value)}
                  placeholder="C:\\data\\tdm"
                />
                <Button type="button" variant="outline" onClick={pickDataDir}>
                  <FolderOpen className="h-4 w-4" />
                  폴더 선택
                </Button>
              </div>
            ) : null}

            {currentStep.key === "dataFileName" ? (
              <Input
                value={form.dataFileName}
                onChange={(e) => updateField("dataFileName", e.target.value)}
                placeholder="예: 2026-1"
              />
            ) : null}

            {currentStep.key === "dailyTest" ? (
              <Textarea
                value={form.dailyTest}
                onChange={(e) => updateField("dailyTest", e.target.value)}
                className="min-h-40"
                placeholder="시험 결과 안내 문구를 입력해 주세요."
              />
            ) : null}

            {currentStep.key === "makeupTest" ? (
              <Textarea
                value={form.makeupTest}
                onChange={(e) => updateField("makeupTest", e.target.value)}
                className="min-h-40"
                placeholder="재시험 일정 미정 안내 문구를 입력해 주세요."
              />
            ) : null}

            {currentStep.key === "makeupTestDate" ? (
              <Textarea
                value={form.makeupTestDate}
                onChange={(e) => updateField("makeupTestDate", e.target.value)}
                className="min-h-40"
                placeholder="재시험 일정 안내 문구를 입력해 주세요."
              />
            ) : null}
          </div>

          {error ? <p className="text-sm text-red-600">{error}</p> : null}

          <div className="flex justify-between">
            <Button type="button" variant="outline" onClick={goPrev} disabled={step === 0 || saving}>
              이전
            </Button>
            {isLastStep ? (
              <Button type="button" onClick={save} disabled={saving}>
                {saving ? "저장 중..." : "설정 저장"}
              </Button>
            ) : (
              <Button type="button" onClick={goNext} disabled={saving}>
                다음
              </Button>
            )}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
