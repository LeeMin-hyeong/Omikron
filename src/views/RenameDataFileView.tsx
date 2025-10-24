import { useEffect, useState } from "react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { Input } from "@/components/ui/input";
import { Check, ChevronsRight, FileSpreadsheet, Play, } from "lucide-react";
import { rpc } from "pyloid-js";
import { useAppDialog } from "@/components/app-dialog/AppDialogProvider";
import { Spinner } from "@/components/ui/spinner";

export default function RenameDataFileView() {
  const dialog = useAppDialog();
  const [state, setState] = useState<{
    ok: boolean;
    has_class: boolean;
    has_data: boolean;
    has_student: boolean;
    missing: string[];
    data_file_name?: string;
  }>({
    ok: false,
    has_class: false,
    has_data: false,
    has_student: false,
    missing: [],
    data_file_name: undefined,
  });
  const fetchState = async () => {
      try {
        setLoading(true);
        const res = await rpc.call("check_data_files", {});
        setState(res);
      } catch {
        // RPC 사용 불가(브라우저 단독 실행 등) 시엔 통과
        setState({ ok: true, has_class: true, has_data: true, has_student: true, missing: [] });
      } finally {
        setLoading(false);
        setTimeout(() => {
          handleRefresh()
        }, 5000);
      }
    };

  const [running, setRunning] = useState(false);
  const [done, setDone] = useState(false);
  const [dataName, setDataName] = useState(state.data_file_name ?? "");
  const [loading, setLoading] = useState(false);

  const handleRefresh = async () => {
    setLoading(true);
    fetchState();
    setDataName("")
    setDone(false);
    setLoading(false);
  };

  const start = async (dataName: string) => {
    if (running) return;
    setDone(false);

    try {
      setRunning(true);
      const res = await rpc.call("change_data_file_name", {new_filename: dataName});
      if(res?.ok){
        await dialog.confirm({ title: "성공", message: `데이터파일 이름을 ${dataName}(으)로 변경하였습니다.` });
        setDone(true);
      }
    } catch (e: any) {
      await dialog.error({ title: "오류", message: String(e?.message || e) });
    } finally {
      setRunning(false);
    }
  };

  useEffect(() => {
    fetchState();
  }, []);

  return (
    <Card className="h-full rounded-2xl border-border/80 shadow-sm">
      <CardContent className="flex h-full flex-col">
        <div className="mb-3">
          <p className="flex flex-row mt-1 text-sm text-muted-foreground">
            데이터 파일의 이름을 변경합니다.
            {state.data_file_name ? (
              <span className="ml-2">
                (저장된 데이터 파일 이름: <span className="font-medium">{state.data_file_name}</span>)
              </span>
            ) : loading ? null : (
              <span className="ml-2 text-amber-600">
                config.json의 <b>dataFileName</b> 설정이 필요합니다.
              </span>
            )}
          </p>
        </div>
        <Separator className="mb-4" />
        
        <div className="flex flex-col h-full justify-center">
          <div className="h-full w-full grid grid-cols-[1fr_auto_1fr] place-items-center">
            <Card className="rounded-2xl h-[250px] w-[300px] border-border/80 shadow-sm">
              <CardContent className="flex h-full flex-col justify-center">
                <div className="flex flex-col items-center gap-2 text-center">
                  <FileSpreadsheet className="h-8 w-8 text-green-600 mb-2" />
                  <div className="text-sm font-medium mb-2">{
                    loading ? (
                      <div className="flex flex-row items-center gap-2">
                        <Spinner /><p>불러오는 중...</p>
                      </div> 
                    ) : state.data_file_name
                  }</div>
                </div>
              </CardContent>
            </Card>
            <ChevronsRight className="mx-1 h-5 w-5 text-muted-foreground" />
            <Card className="rounded-2xl h-[250px] w-[300px] border-border/80 shadow-sm">
              <CardContent className="flex h-full flex-col justify-center">
                <div className="flex flex-col items-center gap-2 text-center">
                  <FileSpreadsheet className="h-8 w-8 text-green-600" />
                  <Input
                    value={dataName}
                    onChange={(e) => setDataName(e.target.value)}
                    placeholder="데이터 파일 이름"
                  />
                </div>
              </CardContent>
            </Card>
          </div>
        </div>

        <div className="mt-auto flex items-center justify-end">
          <div className="flex items-center gap-2">
            <Button
              className="rounded-xl"
              variant="outline"
              onClick={handleRefresh}
              disabled={loading || loading}
              title="데이터 파일 이름을 다시 불러옵니다."
            >
              {loading ? "불러오는 중…" : "새로고침"}
            </Button>
            <Button className={`rounded-xl text-white ${
              done ? "bg-green-600 hover:bg-green-600/90" : "bg-black hover:bg-black/90"
            }`}
              onClick={async () => {
                if (dataName.trim().length === 0) {
                  await dialog.error({title: "데이터 입력 오류", message: "데이터 파일 이름을 입력하세요."});
                  return;
                }
                if (dataName === state.data_file_name) {
                  await dialog.error({title: "데이터 입력 오류", message: "현재 데이터 파일 이름과 동일합니다."});
                  return;
                }
                start(dataName);
              }}
              disabled={loading}
            >
              {running ? <Spinner className="mr-2 h-4 w-4" /> : done ? <Check className="mr-2 h-4 w-4" /> : <Play className="mr-2 h-4 w-4" /> }
              {running ? "변경 중..." : done ? "완료" : "변경"}
            </Button>
          </div>
        </div>
      </CardContent>
    </Card>
  );
}
