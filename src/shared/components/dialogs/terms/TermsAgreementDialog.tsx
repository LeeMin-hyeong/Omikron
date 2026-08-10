import { useEffect, useState } from "react";
import { rpc } from "pyloid-js";
import { Dialog, DialogContent, DialogFooter, DialogHeader, DialogTitle } from "@/shared/components/ui/dialog";
import { Button } from "@/shared/components/ui/button";
import { Checkbox } from "@/shared/components/ui/checkbox";
import { Label } from "@/shared/components/ui/label";

interface Props {
  open: boolean;
  onAccepted: () => void;
}

export default function TermsAgreementDialog({ open, onAccepted }: Props) {
  const [checked, setChecked] = useState(false);
  const [loading, setLoading] = useState(false);
  const [title, setTitle] = useState("이용약관");
  const [text, setText] = useState("");
  const [error, setError] = useState("");

  useEffect(() => {
    if (!open) return;

    setChecked(false);
    setError("");
    setText("");

    rpc
      .call("get_terms_text", {})
      .then((res) => {
        if (!res?.ok) {
          setError(res?.error ?? "이용약관을 불러오지 못했습니다.");
          return;
        }
        setTitle(res?.title ?? "이용약관");
        setText(res?.text ?? "");
      })
      .catch((e) => setError(String(e)));
  }, [open]);

  const handleAccept = async () => {
    if (!checked) return;
    setLoading(true);
    setError("");
    try {
      const res = await rpc.call("accept_terms", {});
      if (!res?.ok) {
        setError(res?.error ?? "약관 동의 저장에 실패했습니다.");
        return;
      }
      onAccepted();
    } catch (e) {
      setError(String(e));
    } finally {
      setLoading(false);
    }
  };

  const handleCancel = async () => {
    await rpc.call("quit_app", {});
  };

  return (
    <Dialog open={open}>
      <DialogContent
        showCloseButton={false}
        className="sm:max-w-3xl max-h-[85vh] flex flex-col"
        onInteractOutside={(e) => e.preventDefault()}
        onEscapeKeyDown={(e) => e.preventDefault()}
      >
        <DialogHeader>
          <DialogTitle>{title}</DialogTitle>
        </DialogHeader>
        <div className="min-h-0 flex-1 overflow-auto rounded-md border bg-slate-50 p-4 text-sm whitespace-pre-wrap">
          {text || "약관을 불러오는 중입니다..."}
        </div>
        <div className="flex items-center gap-2">
          <Checkbox
            id="terms-agree"
            checked={checked}
            onCheckedChange={(v) => setChecked(v === true)}
          />
          <Label htmlFor="terms-agree">이용약관에 동의합니다.</Label>
        </div>
        {error ? <p className="text-sm text-red-600">{error}</p> : null}
        <DialogFooter>
          <Button variant="outline" onClick={handleCancel}>
            취소
          </Button>
          <Button onClick={handleAccept} disabled={!checked || loading}>
            확인
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
