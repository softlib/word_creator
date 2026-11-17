unit uDocxTemplate;

{
  uDocxTemplate.pas
  ---------------------------------------------------------------------------
  通用 Word 内容生成器（Delphi 11.3 + OfficeXML4D）。

  只依赖工作空间中的 OfficeXML4D 库和系统自带单元（System.JSON 等），
  通过 JSON 配置描述文档内容，生成包含文字、格式、列表、表格、
  分页以及嵌入式图片的 .docx 文件。

  典型用法：

    var Doc: IWordDocument;
    Doc := TDocxTemplate.Generate('config.json', 'output.docx');

  或复用对象：

    var Tpl: TDocxTemplate;
    Tpl := TDocxTemplate.Create;
    try
      Tpl.LoadConfig('config.json');   // 只解析并构建文档
      Tpl.SaveTo('output.docx');
    finally
      Tpl.Free;
    end;

  配置说明见随附的 config.sample.json。
  ---------------------------------------------------------------------------
}

interface

uses
  System.SysUtils,
  System.Classes,
  System.JSON,
  Office4D.Word;

type
  EDocxTemplateError = class(Exception);

  TDocxTemplate = class
  private
    FDoc: IWordDocument;
    FBaseDir: string;

    procedure ApplyParagraphProps(Para: IWordParagraph; Obj: TJSONObject);
    procedure ApplyRunProps(Run: IWordRun; Obj: TJSONObject);

    procedure RenderBlock(Block: TJSONObject);
    procedure RenderTextBlock(Block: TJSONObject);
    procedure RenderRunsBlock(Block: TJSONObject);
    procedure RenderListBlock(Block: TJSONObject);
    procedure RenderTableBlock(Block: TJSONObject);
    procedure RenderImageBlock(Block: TJSONObject);
    procedure RenderPageBreak;

    function ResolveImage(const Block: TJSONObject; out Data: TBytes;
      out Ext: string): Boolean;
    function PrepareImage(const Block: TJSONObject; out Data: TBytes;
      out Ext: string; out W, H: Integer): Boolean;
    class function ReadImageSize(const Data: TBytes; out W, H: Integer): Boolean; static;
  public
    constructor Create;
    procedure LoadConfig(const FileName: string);overload;
    procedure LoadConfig(const AJsonText: string;ABaseDir:string); overload;
    procedure SaveTo(const FileName: string);

    class function Generate(const ConfigFile, OutputFile: string): IWordDocument;overload; static;

    class function Generate(const AJsonText,ABaseDir, OutputFile: string): IWordDocument; overload;static;

    property Document: IWordDocument read FDoc;
  end;

implementation

uses
  System.IOUtils,
  System.Math,
  System.Generics.Collections,
  System.NetEncoding;

const
  DefaultFontName = 'Microsoft YaHei';
  DefaultFontSizePt = 10;
  DefaultMaxImageWidthPx = 560;

type
  TIntPair = record
    A, B: Integer;
  end;

{ ------------------------------ JSON helpers ------------------------------ }

function JValue(Obj: TJSONObject; const Name: string): TJSONValue;
begin
  Result := nil;
  if Obj <> nil then
    Result := Obj.GetValue(Name);
end;

function JStr(Obj: TJSONObject; const Name: string;
  const Default: string = ''): string;
var
  V: TJSONValue;
begin
  Result := Default;
  V := JValue(Obj, Name);
  if (V = nil) or (V is TJSONNull) then
    Exit;
  if V is TJSONString then
    Result := TJSONString(V).Value
  else if V is TJSONNumber then
    Result := TJSONNumber(V).ToString
  else if V is TJSONBool then
    Result := BoolToStr(TJSONBool(V).AsBoolean, True)
  else
    Result := V.ToString;
end;

function JInt(Obj: TJSONObject; const Name: string; Default: Integer): Integer;
var
  V: TJSONValue;
  S: string;
begin
  Result := Default;
  V := JValue(Obj, Name);
  if V = nil then
    Exit;
  if V is TJSONNumber then
    Result := TJSONNumber(V).AsInt
  else if V is TJSONString then
  begin
    S := Trim(TJSONString(V).Value);
    if not TryStrToInt(S, Result) then
      Result := Default;
  end;
end;

function JBool(Obj: TJSONObject; const Name: string; Default: Boolean): Boolean;
var
  V: TJSONValue;
  S: string;
begin
  Result := Default;
  V := JValue(Obj, Name);
  if V = nil then
    Exit;
  if V is TJSONBool then
    Result := TJSONBool(V).AsBoolean
  else if V is TJSONNumber then
    Result := TJSONNumber(V).AsInt <> 0
  else if V is TJSONString then
  begin
    S := LowerCase(Trim(TJSONString(V).Value));
    Result := S = 'true';
  end;
end;

function JDouble(Obj: TJSONObject; const Name: string; Default: Double): Double;
var
  V: TJSONValue;
  S: string;
begin
  Result := Default;
  V := JValue(Obj, Name);
  if V = nil then
    Exit;
  if V is TJSONNumber then
    Result := TJSONNumber(V).AsDouble
  else if V is TJSONString then
  begin
    S := Trim(TJSONString(V).Value);
    if not TryStrToFloat(S, Result, TFormatSettings.Invariant) then
      Result := Default;
  end;
end;

function JArray(Obj: TJSONObject; const Name: string): TJSONArray;
var
  V: TJSONValue;
begin
  Result := nil;
  V := JValue(Obj, Name);
  if V is TJSONArray then
    Result := TJSONArray(V);
end;

function JObj(Obj: TJSONObject; const Name: string): TJSONObject;
var
  V: TJSONValue;
begin
  Result := nil;
  V := JValue(Obj, Name);
  if V is TJSONObject then
    Result := TJSONObject(V);
end;

{ --------------------------- small value helpers -------------------------- }

function FirstNonEmpty(const A, B: string): string;
begin
  if A <> '' then
    Result := A
  else
    Result := B;
end;

function NormalizeColor(const S: string): string;
begin
  Result := Trim(S);
  if (Result <> '') and (Result[1] = '#') then
    Delete(Result, 1, 1);
end;

function PtToTwips(Pt: Integer): Integer;
begin
  Result := Pt * 20;
end;

function PtToHalfPoints(Pt: Integer): Integer;
begin
  Result := Pt * 2;
end;

function ParseAlign(const S: string; Default: TParagraphAlignment): TParagraphAlignment;
var
  L: string;
begin
  L := LowerCase(Trim(S));
  if (L = 'center') or (L = 'centre') or (L = 'middle') or (L = '居中') then
    Result := TParagraphAlignment.Center
  else if (L = 'right') or (L = 'end') or (L = '右') then
    Result := TParagraphAlignment.Right
  else if (L = 'justify') or (L = 'both') or (L = '两端') then
    Result := TParagraphAlignment.Justify
  else if (L = 'left') or (L = 'start') or (L = '左') then
    Result := TParagraphAlignment.Left
  else
    Result := Default;
end;

function ParseListStyle(const S: string): TListStyle;
var
  L: string;
begin
  L := LowerCase(Trim(S));
  if (L = 'bullet') or (L = 'unordered') or (L = 'ul') or (L = '项目') then
    Result := TListStyle.Bullet
  else if (L = 'number') or (L = 'numbered') or (L = 'ordered') or (L = 'ol') or (L = '编号') then
    Result := TListStyle.Numbered
  else
    Result := TListStyle.None;
end;

function ReadBE16(const D: TBytes; Index: Integer): Integer;
begin
  Result := (D[Index] shl 8) or D[Index + 1];
end;

function ReadBE32(const D: TBytes; Index: Integer): Integer;
begin
  Result := (D[Index] shl 24) or (D[Index + 1] shl 16) or
            (D[Index + 2] shl 8) or D[Index + 3];
end;

function ReadLE32(const D: TBytes; Index: Integer): Integer;
begin
  Result := D[Index] or (D[Index + 1] shl 8) or
            (D[Index + 2] shl 16) or (D[Index + 3] shl 24);
end;

{ ------------------------------- TDocxTemplate ---------------------------- }

constructor TDocxTemplate.Create;
begin
  inherited Create;
  FDoc := TWordDocumentFactory.CreateDocument;
  FBaseDir := '';
end;

procedure TDocxTemplate.ApplyParagraphProps(Para: IWordParagraph; Obj: TJSONObject);
var
  L: Double;
  S: string;
begin
  if Obj = nil then
    Exit;

  Para.Alignment := ParseAlign(JStr(Obj, 'align', JStr(Obj, 'alignment', 'left')),
    TParagraphAlignment.Left);

  S := JStr(Obj, 'list', JStr(Obj, 'listStyle', ''));
  if S <> '' then
    Para.ListStyle := ParseListStyle(S);

  L := JDouble(Obj, 'lineSpacing', 0);
  if L > 0 then
    Para.LineSpacing := TLineSpacing.Create(Round(L * 240), TLineSpacingRule.Auto,
      PtToTwips(JInt(Obj, 'spaceBefore', 0)), PtToTwips(JInt(Obj, 'spaceAfter', 0)))
  else
    Para.LineSpacing := TLineSpacing.Create(0, TLineSpacingRule.Auto,
      PtToTwips(JInt(Obj, 'spaceBefore', 0)), PtToTwips(JInt(Obj, 'spaceAfter', 0)));

  Para.Indent := TParagraphIndent.Create(
    PtToTwips(JInt(Obj, 'indentLeft', 0)),
    PtToTwips(JInt(Obj, 'indentRight', 0)),
    PtToTwips(JInt(Obj, 'firstLine', 0)));
end;

procedure TDocxTemplate.ApplyRunProps(Run: IWordRun; Obj: TJSONObject);
var
  S: string;
  Size: Integer;
begin
  if Obj = nil then
    Exit;

  S := NormalizeColor(JStr(Obj, 'color', JStr(Obj, 'fontColor', '')));
  if S <> '' then
    Run.FontColor := S;

  S := FirstNonEmpty(JStr(Obj, 'font', ''), JStr(Obj, 'fontName', ''));
  if S = '' then
    S := DefaultFontName;
  Run.FontName := S;

  Size := JInt(Obj, 'fontSize', 0);
  if Size <= 0 then
    Size := DefaultFontSizePt;
  if JInt(Obj, 'fontSizeHalfPoints', 0) > 0 then
    Run.FontSize := JInt(Obj, 'fontSizeHalfPoints', 0)
  else
    Run.FontSize := PtToHalfPoints(Size);

  Run.Bold := JBool(Obj, 'bold', False);
  Run.Italic := JBool(Obj, 'italic', False);
  Run.Underline := JBool(Obj, 'underline', False);

  S := FirstNonEmpty(JStr(Obj, 'link', ''), JStr(Obj, 'hyperlink', ''));
  if S <> '' then
    Run.Hyperlink := S;
end;

procedure TDocxTemplate.RenderTextBlock(Block: TJSONObject);
var
  Para: IWordParagraph;
  Run: IWordRun;
  S: string;
  Level: Integer;
begin
  Para := FDoc.AddParagraph;
  ApplyParagraphProps(Para, Block);

  S := JStr(Block, 'type', '');
  if (S = 'heading') or (S = 'title') then
  begin
    Level := JInt(Block, 'level', 1);
    if Level < 1 then
      Level := 1;
    if Level > 4 then
      Level := 4;
    Run := Para.AddRun(JStr(Block, 'text', ''));
    Run.Bold := True;
    case Level of
      1: Run.FontSize := PtToHalfPoints(20);
      2: Run.FontSize := PtToHalfPoints(16);
      3: Run.FontSize := PtToHalfPoints(13);
    else
      Run.FontSize := PtToHalfPoints(11);
    end;
    Run.FontName := DefaultFontName;
    Run.FontColor := NormalizeColor(JStr(Block, 'color', '1F3864'));
  end
  else
  begin
    Run := Para.AddRun(JStr(Block, 'text', JStr(Block, 'value', '')));
    ApplyRunProps(Run, Block);
  end;
end;

procedure TDocxTemplate.RenderRunsBlock(Block: TJSONObject);
var
  Para: IWordParagraph;
  Runs: TJSONArray;
  I: Integer;
  Item: TJSONObject;
  Run: IWordRun;
begin
  Para := FDoc.AddParagraph;
  ApplyParagraphProps(Para, Block);
  Runs := JArray(Block, 'runs');
  if Runs = nil then
    Exit;
  for I := 0 to Runs.Count - 1 do
  begin
    if not (Runs.Items[I] is TJSONObject) then
      Continue;
    Item := TJSONObject(Runs.Items[I]);
    if JBool(Item, 'image', False) then
    begin
      // 行内图片：与文字混排在当前段落
      var ImgData: TBytes;
      var ImgExt: string;
      var ImgW, ImgH: Integer;
      if PrepareImage(Item, ImgData, ImgExt, ImgW, ImgH) then
      begin
        Run := Para.AddRun('');
        Run.AddImage(ImgData, ImgExt, ImgW, ImgH);
      end;
      Continue;
    end;
    Run := Para.AddRun(JStr(Item, 'text', ''));
    ApplyRunProps(Run, Item);
  end;
end;

procedure TDocxTemplate.RenderListBlock(Block: TJSONObject);
var
  Items: TJSONArray;
  I: Integer;
  Para: IWordParagraph;
  Run: IWordRun;
  Style: string;
begin
  Style := JStr(Block, 'style', JStr(Block, 'listStyle', 'bullet'));
  Items := JArray(Block, 'items');
  if Items = nil then
    Exit;
  for I := 0 to Items.Count - 1 do
  begin
    Para := FDoc.AddParagraph;
    Para.ListStyle := ParseListStyle(Style);
    ApplyParagraphProps(Para, Block);
    if Items.Items[I] is TJSONObject then
      Run := Para.AddRun(JStr(TJSONObject(Items.Items[I]), 'text', ''))
    else
      Run := Para.AddRun(Items.Items[I].Value);
    ApplyRunProps(Run, JObj(Block, 'run'));
  end;
end;

procedure TDocxTemplate.RenderTableBlock(Block: TJSONObject);
var
  Rows, Widths: TJSONArray;
  RowArr: TJSONArray;
  Table: IWordTable;
  R, C, ColCount: Integer;
  CellObj: TJSONObject;
  HeaderShading, BorderColor: string;
  BorderWidth: Integer;
begin
  Rows := JArray(Block, 'rows');
  if Rows = nil then
    Rows := JArray(Block, 'cells');
  if (Rows = nil) or (Rows.Count = 0) then
    Exit;

  ColCount := 0;
  if Rows.Items[0] is TJSONArray then
    ColCount := TJSONArray(Rows.Items[0]).Count;
  if ColCount <= 0 then
    raise EDocxTemplateError.Create('表格配置 rows 的第一行必须是非空数组');

  Table := FDoc.AddTable(Rows.Count, ColCount);

  BorderColor := NormalizeColor(JStr(Block, 'borderColor', '808080'));
  BorderWidth := JInt(Block, 'borderWidth', 4);
  if JBool(Block, 'borders', True) then
    Table.Borders := TTableBorders.All(
      TTableBorder.Create(TBorderStyle.Single, BorderWidth, BorderColor))
  else
    Table.Borders := TTableBorders.None;

    Widths := JArray(Block, 'widths');
  if Widths = nil then
    Widths := JArray(Block, 'columnWidths');
  if Widths <> nil then
  begin
    var W: TArray<Integer>;
    SetLength(W, Widths.Count);
    for C := 0 to Widths.Count - 1 do
      if Widths.Items[C] is TJSONNumber then
        W[C] := TJSONNumber(Widths.Items[C]).AsInt
      else
        W[C] := StrToIntDef(Widths.Items[C].Value, 0);
    Table.SetColumnWidths(W);
  end;

  HeaderShading := NormalizeColor(JStr(Block, 'headerShading', 'D9E2F3'));

  for R := 0 to Rows.Count - 1 do
  begin
    if not (Rows.Items[R] is TJSONArray) then
      Continue;
    RowArr := TJSONArray(Rows.Items[R]);
    for C := 0 to ColCount - 1 do
    begin
      if C >= RowArr.Count then
        Break;
      if RowArr.Items[C] is TJSONObject then
      begin
        CellObj := TJSONObject(RowArr.Items[C]);
        Table.Cells[R, C].Text := JStr(CellObj, 'text', '');
        Table.Cells[R, C].Shading := NormalizeColor(JStr(CellObj, 'shading', ''));
      end
      else
        Table.Cells[R, C].Text := RowArr.Items[C].Value;

      if (R = 0) and JBool(Block, 'headerRow', False) and
         (Table.Cells[R, C].Shading = '') then
        Table.Cells[R, C].Shading := HeaderShading;
    end;
  end;
end;

class function TDocxTemplate.ReadImageSize(const Data: TBytes;
  out W, H: Integer): Boolean;
var
  Len: Integer;

  function Signature(const Bytes: array of Byte): Boolean;
  var
    I: Integer;
  begin
    Result := Length(Data) >= Length(Bytes);
    if not Result then
      Exit;
    for I := 0 to High(Bytes) do
      if Data[I] <> Bytes[I] then
        Exit(False);
  end;

var
  I, Marker, SegLen: Integer;
begin
  Result := False;
  W := 0;
  H := 0;
  Len := Length(Data);

  // PNG
  if Len >= 24 then
  begin
    if Signature([$89, $50, $4E, $47, $0D, $0A, $1A, $0A]) then
    begin
      W := ReadBE32(Data, 16);
      H := ReadBE32(Data, 20);
      Exit((W > 0) and (H > 0));
    end;
  end;

  // GIF
  if Len >= 10 then
  begin
    if Signature([$47, $49, $46, $38]) then
    begin
      W := Data[6] or (Data[7] shl 8);
      H := Data[8] or (Data[9] shl 8);
      Exit((W > 0) and (H > 0));
    end;
  end;

  // BMP
  if (Len >= 26) and (Data[0] = $42) and (Data[1] = $4D) then
  begin
    W := ReadLE32(Data, 18);
    H := ReadLE32(Data, 22);
    if H < 0 then
      H := -H;
    Exit((W > 0) and (H > 0));
  end;

  // JPEG / JFIF / EXIF
  if (Len >= 16) and (Data[0] = $FF) and (Data[1] = $D8) then
  begin
    I := 2;
    while I + 9 < Len do
    begin
      if Data[I] <> $FF then
      begin
        Inc(I);
        Continue;
      end;
      while (I < Len) and (Data[I] = $FF) do
        Inc(I);
      if I >= Len then
        Break;
      Marker := Data[I];
      Inc(I);

      if (Marker = $D8) or (Marker = $D9) or (Marker = $01) or
         ((Marker >= $D0) and (Marker <= $D7)) then
        Continue;
      if I + 1 >= Len then
        Break;

      SegLen := ReadBE16(Data, I);
      if SegLen < 2 then
        Break;

      if ((Marker >= $C0) and (Marker <= $CF)) and
         not (Marker in [$C4, $C8, $CC]) then
      begin
        if I + 8 < Len then
        begin
          H := ReadBE16(Data, I + 3);
          W := ReadBE16(Data, I + 5);
          Exit((W > 0) and (H > 0));
        end;
      end;
      Inc(I, SegLen);
    end;
  end;
end;

function TDocxTemplate.ResolveImage(const Block: TJSONObject;
  out Data: TBytes; out Ext: string): Boolean;
var
  Path, B64, S: string;
begin
  Result := False;
  Data := nil;

  S := LowerCase(Trim(JStr(Block, 'ext',
    JStr(Block, 'imageType', JStr(Block, 'format', '')))));
  if S = '' then
    S := 'png';
  Ext := S;

  Path := Trim(JStr(Block, 'path', JStr(Block, 'file', '')));
  if (Path <> '') and (FBaseDir <> '') and (not TPath.IsPathRooted(Path)) then
    Path := TPath.Combine(FBaseDir, Path);

  if (Path <> '') and TFile.Exists(Path) then
  begin
    Data := TFile.ReadAllBytes(Path);
    if Ext = 'png' then
    begin
      // 没有显式扩展名时按文件真实后缀推断
      S := LowerCase(TPath.GetExtension(Path));
      if (S = '.jpg') or (S = '.jpeg') or (S = '.gif') or (S = '.bmp') or (S = '.png') then
        Ext := Copy(S, 2, MaxInt);
    end;
    Result := Length(Data) > 0;
    Exit;
  end;

  B64 := Trim(JStr(Block, 'base64', ''));
  if B64 <> '' then
  begin
    var Enc: TBase64Encoding;
    Enc := TBase64Encoding.Create;
    try
      Data := Enc.Decode(TEncoding.ASCII.GetBytes(B64));
    finally
      Enc.Free;
    end;
    Result := Length(Data) > 0;
  end;
end;

function TDocxTemplate.PrepareImage(const Block: TJSONObject;
  out Data: TBytes; out Ext: string; out W, H: Integer): Boolean;
var
  ImgW, ImgH, MaxW: Integer;
  Scale: Double;
begin
  Result := ResolveImage(Block, Data, Ext);
  if not Result then
    Exit;

  W := JInt(Block, 'width', 0);
  H := JInt(Block, 'height', 0);
  MaxW := JInt(Block, 'maxWidth', DefaultMaxImageWidthPx);
  if MaxW <= 0 then
    MaxW := DefaultMaxImageWidthPx;
  Scale := JDouble(Block, 'scale', 0);

  if ReadImageSize(Data, ImgW, ImgH) then
  begin
    if (W > 0) and (H <= 0) then
      H := Max(1, Round(W * ImgH / ImgW))
    else if (H > 0) and (W <= 0) then
      W := Max(1, Round(H * ImgW / ImgH))
    else if (W <= 0) and (H <= 0) then
    begin
      W := ImgW;
      H := ImgH;
    end;
    if Scale > 0 then
    begin
      W := Max(1, Round(W * Scale));
      H := Max(1, Round(H * Scale));
    end;
    if (W > MaxW) then
    begin
      H := Max(1, Round(H * MaxW / W));
      W := MaxW;
    end;
  end
  else
  begin
    if W <= 0 then
      W := MaxW;
    if H <= 0 then
      H := W;
  end;
end;

procedure TDocxTemplate.RenderImageBlock(Block: TJSONObject);
var
  Data: TBytes;
  Ext: string;
  W, H: Integer;
  Para: IWordParagraph;
  Run: IWordRun;
begin
  if not PrepareImage(Block, Data, Ext, W, H) then
    raise EDocxTemplateError.Create('图片不存在或为空：' +
      JStr(Block, 'path', '[base64]'));

  Para := FDoc.AddParagraph;
  Para.Alignment := ParseAlign(JStr(Block, 'align', 'center'),
    TParagraphAlignment.Center);
  Para.LineSpacing := TLineSpacing.Single;
  Run := Para.AddRun('');
  Run.AddImage(Data, Ext, W, H);

  // 图注（可选）
  if JStr(Block, 'caption', '') <> '' then
  begin
    Para := FDoc.AddParagraph;
    Para.Alignment := TParagraphAlignment.Center;
    Para.LineSpacing := TLineSpacing.Create(240, TLineSpacingRule.Auto, 60, 120);
    Run := Para.AddRun(JStr(Block, 'caption', ''));
    Run.FontName := DefaultFontName;
    Run.FontSize := PtToHalfPoints(9);
    Run.Italic := True;
    Run.FontColor := '595959';
  end;
end;

procedure TDocxTemplate.RenderPageBreak;
var
  Para: IWordParagraph;
begin
  Para := FDoc.AddParagraph;
  Para.AddPageBreak;
end;

procedure TDocxTemplate.RenderBlock(Block: TJSONObject);
var
  Kind: string;
begin
  Kind := LowerCase(Trim(JStr(Block, 'type', 'paragraph')));
  if (Kind = 'heading') or (Kind = 'title') or (Kind = 'paragraph') or
     (Kind = 'text') then
    RenderTextBlock(Block)
  else if (Kind = 'runs') or (Kind = 'rich') then
    RenderRunsBlock(Block)
  else if (Kind = 'list') or (Kind = 'bullets') or (Kind = 'numbered') then
    RenderListBlock(Block)
  else if Kind = 'table' then
    RenderTableBlock(Block)
  else if (Kind = 'image') or (Kind = 'picture') then
    RenderImageBlock(Block)
  else if (Kind = 'pagebreak') or (Kind = 'newpage') then
    RenderPageBreak
  else if Kind = 'spacer' then
    FDoc.AddParagraph
  else
    raise EDocxTemplateError.Create('未知内容类型：' + Kind);
end;

procedure TDocxTemplate.LoadConfig(const FileName: string);
var
  Text: string;
  Root, Page: TJSONObject;
  Content: TJSONArray;
  I: Integer;
  JSON: TJSONValue;
begin
  if not TFile.Exists(FileName) then
    raise EDocxTemplateError.Create('配置文件不存在：' + FileName);

  Text := TFile.ReadAllText(FileName, TEncoding.UTF8);
  if (Text <> '') and (Text[1] = #$FEFF) then
    Delete(Text, 1, 1);

  JSON := TJSONObject.ParseJSONValue(Text);
  if JSON = nil then
    raise EDocxTemplateError.Create('JSON 解析失败：' + FileName);
  try
    if not (JSON is TJSONObject) then
      raise EDocxTemplateError.Create('配置根节点必须是 JSON 对象');
    Root := TJSONObject(JSON);

    FBaseDir := TPath.GetDirectoryName(TPath.GetFullPath(FileName));
    FDoc := TWordDocumentFactory.CreateDocument;

    Page := JObj(Root, 'page');
    if Page <> nil then
    begin
      if LowerCase(JStr(Page, 'orientation', 'portrait')) = 'landscape' then
        FDoc.PageOrientation := TPageOrientation.Landscape
      else
        FDoc.PageOrientation := TPageOrientation.Portrait;

      FDoc.PageMargins := TPageMargins.Create(
        PtToTwips(JInt(Page, 'marginTop', 72)),
        PtToTwips(JInt(Page, 'marginBottom', 72)),
        PtToTwips(JInt(Page, 'marginLeft', 72)),
        PtToTwips(JInt(Page, 'marginRight', 72)));
    end
    else
      FDoc.PageMargins := TPageMargins.Create(1440, 1440, 1440, 1440);

    FDoc.Header.Text := JStr(Root, 'header', '');
    FDoc.Footer.Text := JStr(Root, 'footer', '');

    Content := JArray(Root, 'content');
    if Content = nil then
      raise EDocxTemplateError.Create('配置缺少 content 数组');

    for I := 0 to Content.Count - 1 do
      if Content.Items[I] is TJSONObject then
        RenderBlock(TJSONObject(Content.Items[I]));
  finally
    JSON.Free;
  end;
end;

procedure TDocxTemplate.LoadConfig(const AJsonText: string;ABaseDir:string);
var
  Text: string;
  Root, Page: TJSONObject;
  Content: TJSONArray;
  I: Integer;
  JSON: TJSONValue;
begin
  if ''=AJsonText.Trim then  Exit;
  

  JSON := TJSONObject.ParseJSONValue(AJsonText);
  if JSON = nil then
    raise EDocxTemplateError.Create('JSON 解析失败：' + AJsonText);
  try
    if not (JSON is TJSONObject) then
      raise EDocxTemplateError.Create('配置根节点必须是 JSON 对象');
    Root := TJSONObject(JSON);

    ////C:\Users\yls\Desktop\pdfview\rar_creator
   // FBaseDir := TPath.GetDirectoryName(TPath.GetFullPath(FileName));
    FBaseDir :=ABaseDir;
    FDoc := TWordDocumentFactory.CreateDocument;

    Page := JObj(Root, 'page');
    if Page <> nil then
    begin
      if LowerCase(JStr(Page, 'orientation', 'portrait')) = 'landscape' then
        FDoc.PageOrientation := TPageOrientation.Landscape
      else
        FDoc.PageOrientation := TPageOrientation.Portrait;

      FDoc.PageMargins := TPageMargins.Create(
        PtToTwips(JInt(Page, 'marginTop', 72)),
        PtToTwips(JInt(Page, 'marginBottom', 72)),
        PtToTwips(JInt(Page, 'marginLeft', 72)),
        PtToTwips(JInt(Page, 'marginRight', 72)));
    end
    else
      FDoc.PageMargins := TPageMargins.Create(1440, 1440, 1440, 1440);

    FDoc.Header.Text := JStr(Root, 'header', '');
    FDoc.Footer.Text := JStr(Root, 'footer', '');

    Content := JArray(Root, 'content');
    if Content = nil then
      raise EDocxTemplateError.Create('配置缺少 content 数组');

    for I := 0 to Content.Count - 1 do
      if Content.Items[I] is TJSONObject then
        RenderBlock(TJSONObject(Content.Items[I]));
  finally
    JSON.Free;
  end;
end;

procedure TDocxTemplate.SaveTo(const FileName: string);
var
  Dir: string;
begin
  Dir := TPath.GetDirectoryName(TPath.GetFullPath(FileName));
  if (Dir <> '') and (not TDirectory.Exists(Dir)) then
    TDirectory.CreateDirectory(Dir);
  FDoc.SaveToFile(FileName);
end;

class function TDocxTemplate.Generate(const ConfigFile,
  OutputFile: string): IWordDocument;
var
  Tpl: TDocxTemplate;
begin
  Tpl := TDocxTemplate.Create;
  try
    Tpl.LoadConfig(ConfigFile);
    Tpl.SaveTo(OutputFile);
    Result := Tpl.Document;
  finally
    Tpl.Free;
  end;
end;

class function TDocxTemplate.Generate(const AJsonText,ABaseDir, OutputFile: string): IWordDocument;
begin
var
  Tpl: TDocxTemplate;
begin
  Tpl := TDocxTemplate.Create;
  try
    Tpl.LoadConfig(AJsonText,ABaseDir);
    Tpl.SaveTo(OutputFile);
    Result := Tpl.Document;
  finally
    Tpl.Free;
  end;
end;
end;

end.