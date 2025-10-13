import type { ViewProps } from "@/types/omikron";
import { Card, CardContent } from "@/components/ui/card";

export default function GenericComingSoonView({ meta }: ViewProps) {
  return (
    <div className="grid h-full grid-rows-[auto_1fr] gap-5 overflow-hidden">
      <Card className="rounded-2xl border-border/80 shadow-sm">
        <CardContent className="flex h-[500px] items-center justify-center text-muted-foreground">
          작업 전용 화면은 준비 중입니다. 
        </CardContent>
      </Card>
    </div>
  );
}
