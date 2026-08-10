import { useEffect, useState } from "react";
import { rpc } from "pyloid-js";
import { Card, CardContent } from "@/shared/components/ui/card";
import { Label } from "@/shared/components/ui/label";
import { Input } from "@/shared/components/ui/input";
import { Textarea } from "@/shared/components/ui/textarea";
import { Button } from "@/shared/components/ui/button";
import { Spinner } from "@/shared/components/ui/spinner";
import { useAppDialog } from "@/shared/components/dialogs/app/AppDialogProvider";

export default function EditMessageConfigView() {
  const dialog = useAppDialog();
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [url, setUrl] = useState("");
  const [dailyTest, setDailyTest] = useState("");
  const [makeupTest, setMakeupTest] = useState("");
  const [makeupTestDate, setMakeupTestDate] = useState("");

  const loadConfig = async () => {
    setLoading(true);
    try {
      const res = await rpc.call("get_config_status", {});
      if (!res?.ok) {
        await dialog.error({ title: "설정 로드 실패", message: res?.error || "설정을 불러오지 못했습니다." });
        return;
      }
      setUrl(res?.config?.url ?? "");
      setDailyTest(res?.config?.dailyTest ?? "");
      setMakeupTest(res?.config?.makeupTest ?? "");
      setMakeupTestDate(res?.config?.makeupTestDate ?? "");
    } catch (e: any) {
      await dialog.error({ title: "설정 로드 실패", message: String(e?.message || e) });
    } finally {
      setLoading(false);
    }
  };

  const saveConfig = async () => {
    if (!url.trim() || !dailyTest.trim() || !makeupTest.trim() || !makeupTestDate.trim()) {
      await dialog.error({ title: "입력 오류", message: "URL과 메시지 템플릿 3종을 모두 입력해 주세요." });
      return;
    }

    const validateRes = await rpc.call("validate_script_url", { url });
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
      const res = await rpc.call("update_message_templates", {
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
    } catch (e: any) {
      await dialog.error({ title: "저장 실패", message: String(e?.message || e) });
    } finally {
      setSaving(false);
    }
  };

  useEffect(() => {
    void loadConfig();
  }, []);

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="space-y-4">
        {loading ? (
          <div className="flex h-40 items-center justify-center">
            <Spinner />
          </div>
        ) : (
          <>
            <div className="space-y-2">
              <Label htmlFor="url">아이소식 스크립트 URL</Label>
              <Input id="url" value={url} onChange={(e) => setUrl(e.target.value)} />
            </div>
            <div className="space-y-2">
              <Label htmlFor="daily">시험 결과 메시지 템플릿</Label>
              <Textarea
                id="daily"
                className="h-33 resize-none overflow-y-auto"
                value={dailyTest}
                onChange={(e) => setDailyTest(e.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="makeup">재시험(일정 미정) 메시지 템플릿</Label>
              <Textarea
                id="makeup"
                className="h-33 resize-none overflow-y-auto"
                value={makeupTest}
                onChange={(e) => setMakeupTest(e.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="makeup-date">재시험(일정 안내) 메시지 템플릿</Label>
              <Textarea
                id="makeup-date"
                className="h-33 resize-none overflow-y-auto"
                value={makeupTestDate}
                onChange={(e) => setMakeupTestDate(e.target.value)}
              />
            </div>
            <div className="flex justify-end gap-2">
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
