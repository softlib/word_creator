program DocxGen;
{$APPTYPE CONSOLE}

{
  DocxGen.exe —— 基于 OfficeXML4D 的通用 .docx 生成工具（Delphi 11.3）

  用法：
    DocxGen.exe -c config.json -o result.docx
    DocxGen.exe config.json result.docx
    DocxGen.exe --config=config.json --output=result.docx
    DocxGen.exe --help
}

uses
  System.SysUtils,
  System.IOUtils,
  System.StrUtils,
  uDocxTemplate in 'uDocxTemplate.pas';

procedure PrintUsage;
begin
  WriteLn('DocxGen - 基于 OfficeXML4D 的通用 Word(.docx) 生成工具');
  WriteLn;
  WriteLn('用法:');
  WriteLn('  DocxGen.exe -c <config.json> -o <output.docx>');
  WriteLn('  DocxGen.exe <config.json> [<output.docx>]');
  WriteLn;
  WriteLn('参数:');
  WriteLn('  -c, --config    内容配置 JSON 文件');
  WriteLn('  -o, --output    输出 docx 文件（默认 output.docx）');
  WriteLn('  -h, --help      显示本帮助');
end;

function TakeValue(const Args: TArray<string>; var I: Integer;
  const Name: string): string;
begin
  if I >= High(Args) then
    raise Exception.Create('参数 ' + Name + ' 缺少取值');
  Inc(I);
  Result := Args[I];
end;

function Execute: Integer;
var
  Args: TArray<string>;
  ConfigFile, OutputFile, S: string;
  I: Integer;
begin
  SetLength(Args, ParamCount);
  for I := 1 to ParamCount do
    Args[I - 1] := ParamStr(I);

  ConfigFile := '';
  OutputFile := '';

  I := 0;
  while I <= High(Args) do
  begin
    S := Args[I];
    if (S = '-h') or (S = '--help') or (S = '/?') then
    begin
      PrintUsage;
      Exit(0);
    end
    else if (S = '-c') or (S = '--config') then
      ConfigFile := TakeValue(Args, I, S)
    else if (S = '-o') or (S = '--output') then
      OutputFile := TakeValue(Args, I, S)
    else if StartsText('--config=', S) then
      ConfigFile := Copy(S, 9, MaxInt)
    else if StartsText('--output=', S) then
      OutputFile := Copy(S, 9, MaxInt)
    else if S <> '' then
    begin
      if ConfigFile = '' then
        ConfigFile := S
      else if OutputFile = '' then
        OutputFile := S
      else
        raise Exception.Create('无法识别的参数：' + S);
    end;
    Inc(I);
  end;

  if ConfigFile = '' then
  begin
    PrintUsage;
    raise Exception.Create('请提供配置文件（ -c <config.json> ）');
  end;

  if OutputFile = '' then
    OutputFile := TPath.Combine(TPath.GetDirectoryName(TPath.GetFullPath(ConfigFile)),
      'output.docx');

  if not TPath.IsPathRooted(OutputFile) then
    OutputFile := TPath.GetFullPath(OutputFile);

  TDocxTemplate.Generate(ConfigFile, OutputFile);

  WriteLn('已生成: ', OutputFile);
  Result := 0;
end;

var
  ExitCode: Integer;
begin
  try
    ExitCode := Execute;
  except
    on E: Exception do
    begin
      WriteLn('错误: ', E.Message);
      ExitCode := 1;
    end;
  end;
  Halt(ExitCode);
end.
