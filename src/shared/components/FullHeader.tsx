import { Badge } from "@/shared/components/ui/badge";

export default function FullHeader({
  title,
  actions,
}: {
  title: string;
  actions?: React.ReactNode;
}) {
  return (
    <div className="mb-2 flex shrink-0 items-center justify-between rounded-2xl border border-border/80 bg-card px-4 py-2 text-card-foreground shadow-sm">
      <div className="flex items-center gap-2">
        <Badge variant="secondary" className="rounded-lg">
          선택된 작업
        </Badge>
        <span className="text-lg font-semibold">{title}</span>
      </div>
      <div className="flex items-center gap-2">{actions}</div>
    </div>
  );
}
