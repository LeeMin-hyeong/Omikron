import { useEffect, useState } from "react";
import { rpc } from "pyloid-js";
import { Dialog, DialogContent, DialogFooter, DialogHeader, DialogTitle } from "@/shared/components/ui/dialog";
import { Button } from "@/shared/components/ui/button";
import { Checkbox } from "@/shared/components/ui/checkbox";
import { Label } from "@/shared/components/ui/label";

interface Props {
  open: boolean;
  title: string;
  message: string;
  noticeId: string;
  onClose: () => void;
}

export default function NoticeDialog({ open, title, message, noticeId, onClose }: Props) {
  const [dontShowAgain, setDontShowAgain] = useState(false);
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    if (open) {
      setDontShowAgain(false);
    }
  }, [open]);

  const handleConfirm = async () => {
    if (dontShowAgain && noticeId) {
      setSubmitting(true);
      try {
        await rpc.call("mark_notice_seen", { notice_id: noticeId });
      } catch {
        // Keep UX flowing even when persisting preference fails.
      } finally {
        setSubmitting(false);
      }
    }
    onClose();
  };

  return (
    <Dialog open={open}>
      <DialogContent
        showCloseButton={false}
        className="sm:max-w-2xl max-h-[80vh] flex flex-col"
        onInteractOutside={(e) => e.preventDefault()}
        onEscapeKeyDown={(e) => e.preventDefault()}
      >
        <DialogHeader>
          <DialogTitle>{title}</DialogTitle>
        </DialogHeader>
        <div className="min-h-0 flex-1 overflow-auto rounded-md border bg-slate-50 p-4 text-sm whitespace-pre-wrap">
          {message}
        </div>
        <div className="flex items-center gap-2">
          <Checkbox
            id="notice-hide"
            checked={dontShowAgain}
            onCheckedChange={(v) => setDontShowAgain(v === true)}
          />
          <Label htmlFor="notice-hide">다시 보지 않기</Label>
        </div>
        <DialogFooter>
          <Button onClick={handleConfirm} disabled={submitting}>
            확인
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
