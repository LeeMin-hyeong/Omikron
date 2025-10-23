import { Card, CardContent } from "@/components/ui/card";
// 경로는 실제 위치에 맞게 조정
import omikron from "@/assets/omikron.png";

export default function WelcomeView() {
  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex flex-col h-full flex items-center justify-center">
        <img
          src={omikron}
          alt="Omikron"
          width={400}
          className="h-auto max-w-full py-3"
        />
        <p>오미크론 프로그램</p>
        <p>왼쪽의 메뉴를 클릭하여 시작하세요</p>
      </CardContent>
    </Card>
  );
}
