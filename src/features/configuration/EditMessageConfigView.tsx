import { useEffect, useState } from "react";
import { useCallback } from "react";
import { generalRpc } from "@/api/rpc";
import { Card, CardContent } from "@/shared/components/ui/card";
import { Label } from "@/shared/components/ui/label";
import { Input } from "@/shared/components/ui/input";
import { Textarea } from "@/shared/components/ui/textarea";
import { Button } from "@/shared/components/ui/button";
import { Spinner } from "@/shared/components/ui/spinner";
import { useAppDialog } from "@/shared/components/dialogs/app/useAppDialog";
import { errorMessage } from "@/shared/utils/errors";

export default function EditMessageConfigView() {
  const dialog = useAppDialog();
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [url, setUrl] = useState("");
  const [dailyTest, setDailyTest] = useState("");
  const [makeupTest, setMakeupTest] = useState("");
  const [makeupTestDate, setMakeupTestDate] = useState("");

  const loadConfig = useCallback(async () => {
    setLoading(true);
    try {
      const res = await generalRpc.call("get_config_status", {});
      if (!res?.ok) {
        await dialog.error({ title: "설정 로드 실패", message: res?.error || "설정을 불러오지 못했습니다." });
        return;
      }
      setUrl(res?.config?.url ?? "");
      setDailyTest(res?.config?.dailyTest ?? "");
      setMakeupTest(res?.config?.makeupTest ?? "");
      setMakeupTestDate(res?.config?.makeupTestDate ?? "");
    } catch (error: unknown) {
      await dialog.error({ title: "설정 로드 실패", message: errorMessage(error) });
    } finally {
      setLoading(false);
    }
  }, [dialog]);

  const saveConfig = async () => {
    if (!url.trim() || !dailyTest.trim() || !makeupTest.trim() || !makeupTestDate.trim()) {
      await dialog.error({ title: "입력 오류", message: "URL과 메시지 템플릿 3종을 모두 입력해 주세요." });
      return;
    }

    const validateRes = await generalRpc.call("validate_script_url", { url });
    if (!validateRes?.ok) {
      await dialog.error({ title: "URL 검증 실패", message: validateRes?.error || "URL을 검증하지 못했습니다." });
      return;
    }
    if (validateRes?.warning) {
      const proceed = await dialog.warning({
        title: "URL 경고",
        message: "URL이 정확하지 않은 것 같습니다. 계속 진행하시겠습니까?",
      });
      if (!proceed) return;
    }

    setSaving(true);
    try {
      const res = await generalRpc.call("update_message_templates", {
        url,
        daily_test_message: dailyTest,
        makeup_test_message: makeupTest,
        makeup_test_date_message: makeupTestDate,
      });
      if (!res?.ok) {
        await dialog.error({ title: "저장 실패", message: res?.error || "설정을 저장하지 못했습니다." });
        return;
      }
      await dialog.confirm({ title: "저장 완료", message: "설정을 저장했습니다." });
      loadConfig();
    } catch (error: unknown) {
      await dialog.error({ title: "저장 실패", message: errorMessage(error) });
    } finally {
      setSaving(false);
    }
  };

  useEffect(() => {
    void loadConfig();
  }, [loadConfig]);

  return (
    <Card className="h-full min-h-0 gap-3 overflow-hidden rounded-2xl border-border/80 py-4 shadow-sm">
      <CardContent className="flex h-full min-h-0 flex-col px-4">
        {loading ? (
          <div className="flex min-h-0 flex-1 items-center justify-center">
            <Spinner />
          </div>
        ) : (
          <>
            <div className="mb-3 shrink-0 space-y-1.5">
              <Label htmlFor="url">아이소식 스크립트 URL</Label>
              <Input id="url" value={url} onChange={(e) => setUrl(e.target.value)} />
            </div>
            <div className="grid min-h-0 flex-1 grid-rows-3 gap-3">
              <div className="flex min-h-0 flex-col gap-1.5">
                <Label htmlFor="daily" className="shrink-0">시험 결과 메시지 템플릿</Label>
                <Textarea
                  id="daily"
                  className="min-h-0 flex-1 resize-none overflow-y-auto"
                  value={dailyTest}
                  onChange={(e) => setDailyTest(e.target.value)}
                />
              </div>
              <div className="flex min-h-0 flex-col gap-1.5">
                <Label htmlFor="makeup" className="shrink-0">재시험(일정 미정) 메시지 템플릿</Label>
                <Textarea
                  id="makeup"
                  className="min-h-0 flex-1 resize-none overflow-y-auto"
                  value={makeupTest}
                  onChange={(e) => setMakeupTest(e.target.value)}
                />
              </div>
              <div className="flex min-h-0 flex-col gap-1.5">
                <Label htmlFor="makeup-date" className="shrink-0">재시험(일정 안내) 메시지 템플릿</Label>
                <Textarea
                  id="makeup-date"
                  className="min-h-0 flex-1 resize-none overflow-y-auto"
                  value={makeupTestDate}
                  onChange={(e) => setMakeupTestDate(e.target.value)}
                />
              </div>
            </div>
            <div className="flex shrink-0 justify-end gap-2 pt-3">
              <Button variant="outline" onClick={loadConfig} disabled={saving}>
                새로고침
              </Button>
              <Button onClick={saveConfig} disabled={saving}>
                {saving ? "저장 중..." : "저장"}
              </Button>
            </div>
          </>
        )}
      </CardContent>
    </Card>
  );
}
