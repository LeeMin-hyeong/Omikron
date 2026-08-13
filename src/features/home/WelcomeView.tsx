import { Card, CardContent } from "@/shared/components/ui/card";
import { generalRpc } from "@/api/rpc";
import tdm from "@/assets/tdm.png";

export default function WelcomeView() {
  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="relative flex h-full flex-col text-center">
        <div className="flex flex-1 flex-col items-center justify-center">
          <img
            src={tdm}
            alt="tdm"
            width={96}
            className="h-auto max-w-full py-3"
          />
          <p>테스트 데이터 관리 프로그램</p>
          <p>왼쪽의 메뉴를 클릭하여 시작하세요</p>
        </div>
        <div className="absolute bottom-2 left-0 right-0 text-xs text-muted-foreground">
          Icons by{" "}
          <button
            type="button"
            onClick={() => generalRpc.call("open_url", { url: "https://icons8.com" })}
            className="cursor-pointer underline underline-offset-2"
          >
            Icons8
          </button>
        </div>
      </CardContent>
    </Card>
  );
}
