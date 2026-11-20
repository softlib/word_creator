unit uDocxTemplate;

{
  uDocxTemplate.pas
  ---------------------------------------------------------------------------
  通用 Word 内容生成器（Delphi 11.3 + OfficeXML4D）。

  只依赖工作空间中的 OfficeXML4D 库和系统自带单元（System.JSON 等），
  通过 JSON 配置描述文档内容，生成包含文字、样式、列表、表格、
  分页、嵌入式图片、目录书签跳转以及首行字符缩进（中英文自适应）
  的 .docx 文件。

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
  注意：firstLineChars / firstLineIndentChars 的单位是“字符个数”，
  写出时换算为 OOXML 的百分之一字符；firstLine / firstLineTwips 为 twips。
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

  THeadingEntry = record
    Name: string;      // 书签内部名称，如 _Toc1
    Text: string;      // 标题文字
    Level: Integer;    // 标题层级
    ParaIndex: Integer;// 在正文段落中的序号（从 0 开始）
  end;

  TDocxTemplate = class
  private
    FDoc: IWordDocument;
    FBaseDir: string;
    FHeadings: TArray<THeadingEntry>;
    FTocEnabled: Boolean;
    FTocTitle: string;
    FTocMaxLevel: Integer;
    FTocPageNumbers: Boolean;
    FTocPageBreakAfter: Boolean;
    FDefaultFirstLineChars: Integer;
    FParaFirstLineChars: TArray<Integer>; // 与正文段落一一对应的首行缩进字符数

    function AddTrackedParagraph(FirstLineChars: Integer): IWordParagraph;
    function BlockFirstLineChars(Block: TJSONObject; DefaultChars: Integer): Integer;
    procedure ApplyParagraphProps(Para: IWordParagraph; Obj: TJSONObject;
      FirstLineChars: Integer);
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

    procedure LoadFromJsonText(const AJsonText, ABaseDir: string);
    procedure ApplyRootSettings(Root: TJSONObject);
    procedure RenderContent(Root: TJSONObject);
    procedure ParseTocSettings(Root: TJSONObject);
    function BuildTocXml: string;
    function InjectBookmarks(const Xml: string): string;
    function InjectFirstLineChars(const Xml: string): string;
    function SaveWithToc(const FileName: string): Boolean;
  public
    constructor Create;
    procedure LoadConfig(const FileName: string); overload;
    procedure LoadConfig(const AJsonText, ABaseDir: string); overload;
    procedure SaveTo(const FileName: string);

    class function Generate(const ConfigFile, OutputFile: string): IWordDocument; overload; static;
    class function Generate(const AJsonText, ABaseDir, OutputFile: string): IWordDocument; overload; static;

    property Document: IWordDocument read FDoc;
  end;

implementation

uses
  System.IOUtils,
  System.Math,
  System.StrUtils,
  System.Zip,
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

function XmlEscape(const S: string): string;
var
  I: Integer;
begin
  Result := '';
  for I := 1 to Length(S) do
    case S[I] of
      '&': Result := Result + '&amp;';
      '<': Result := Result + '&lt;';
      '>': Result := Result + '&gt;';
      '"': Result := Result + '&quot;';
      '''': Result := Result + '&apos;';
    else
      Result := Result + S[I];
    end;
end;

function NextRelationshipId(const Rels: string): Integer;
var
  P: Integer;
  S: string;
  N: Integer;
begin
  Result := 1;
  P := Pos('Id="rId', Rels);
  while P > 0 do
  begin
    S := Copy(Rels, P + 7, 24);
    if Pos('"', S) > 0 then
    begin
      N := StrToIntDef(Copy(S, 1, Pos('"', S) - 1), 0);
      if N >= Result then
        Result := N + 1;
    end;
    P := PosEx('Id="rId', Rels, P + 7);
  end;
end;

const
  WordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  SettingsPart = 'word/settings.xml';
  SettingsRelationshipType =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
  SettingsContentType =
    'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';

{ 首行缩进辅助：判断文本是否含中日韩全角字符 }

function HasWideChar(const S: string): Boolean;
var
  I: Integer;
  C: Word;
begin
  Result := False;
  for I := 1 to Length(S) do
  begin
    C := Ord(S[I]);
    if ((C >= $1100) and (C <= $115F)) or ((C >= $2E80) and (C <= $A4CF)) or
       ((C >= $AC00) and (C <= $D7A3)) or ((C >= $F900) and (C <= $FAFF)) or
       ((C >= $FE30) and (C <= $FE6F)) or ((C >= $FF00) and (C <= $FF60)) or
       ((C >= $FFE0) and (C <= $FFE6)) then
      Exit(True);
  end;
end;

{ 为不支持 firstLineChars 的阅读器给出 twips 兜底值：
  中文按两个全角字符宽，英文按两个半角字符宽。 }
function EstimateIndentTwips(Block: TJSONObject; Chars: Integer): Integer;
var
  Text: string;
  Runs: TJSONArray;
  I: Integer;
  FontPt: Double;
begin
  Text := JStr(Block, 'text', JStr(Block, 'value', ''));
  Runs := JArray(Block, 'runs');
  if Runs <> nil then
    for I := 0 to Runs.Count - 1 do
      if Runs.Items[I] is TJSONObject then
        Text := Text + JStr(TJSONObject(Runs.Items[I]), 'text', '');

  FontPt := JDouble(Block, 'fontSize', DefaultFontSizePt);
  if FontPt <= 0 then
    FontPt := DefaultFontSizePt;

  if HasWideChar(Text) then
    Result := Round(Chars * FontPt * 20)   // 全角：1 字符 = 1 磅字号
  else
    Result := Round(Chars * FontPt * 10);  // 半角：1 字符 = 半个字号
end;

{ ------------------------------- TDocxTemplate ---------------------------- }

constructor TDocxTemplate.Create;
begin
  inherited Create;
  FDoc := TWordDocumentFactory.CreateDocument;
  FBaseDir := '';
  FHeadings := nil;
  FParaFirstLineChars := nil;
  FDefaultFirstLineChars := 2;
  FTocEnabled := True;
  FTocTitle := '目录';
  FTocMaxLevel := 3;
  FTocPageNumbers := True;
  FTocPageBreakAfter := True;
end;

function TDocxTemplate.AddTrackedParagraph(
  FirstLineChars: Integer): IWordParagraph;
begin
  Result := FDoc.AddParagraph;
  SetLength(FParaFirstLineChars, Length(FParaFirstLineChars) + 1);
  FParaFirstLineChars[High(FParaFirstLineChars)] := FirstLineChars;
end;

function TDocxTemplate.BlockFirstLineChars(Block: TJSONObject;
  DefaultChars: Integer): Integer;
var
  HasExplicitTwips: Boolean;
begin
  // 无显式字符配置时：段落块用全局默认值，其他块默认不缩进。
  // 若块内同时给了 firstLine(twips)，则保留 twips 写法，不再按字符注入。
  Result := DefaultChars;
  if Block = nil then
    Exit;

  HasExplicitTwips := (JValue(Block, 'firstLine') <> nil) or
    (JValue(Block, 'firstLineTwips') <> nil);

  if JValue(Block, 'firstLineChars') <> nil then
    Result := JInt(Block, 'firstLineChars', DefaultChars)
  else if JValue(Block, 'indentChars') <> nil then
    Result := JInt(Block, 'indentChars', DefaultChars)
  else if HasExplicitTwips then
    Result := 0;

  if Result < 0 then
    Result := 0;
  if Result > 20 then
    Result := 20;
end;

procedure TDocxTemplate.ApplyParagraphProps(Para: IWordParagraph;
  Obj: TJSONObject; FirstLineChars: Integer);
var
  L: Double;
  S: string;
  FirstLineTwips: Integer;
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

  FirstLineTwips := PtToTwips(JInt(Obj, 'firstLine', 0));
  if JValue(Obj, 'firstLineTwips') <> nil then
    FirstLineTwips := JInt(Obj, 'firstLineTwips', 0);
  if (FirstLineTwips = 0) and (FirstLineChars > 0) then
    FirstLineTwips := EstimateIndentTwips(Obj, FirstLineChars);

  Para.Indent := TParagraphIndent.Create(
    PtToTwips(JInt(Obj, 'indentLeft', 0)),
    PtToTwips(JInt(Obj, 'indentRight', 0)),
    FirstLineTwips);
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
  Level, Chars: Integer;
begin
  S := LowerCase(Trim(JStr(Block, 'type', 'paragraph')));
  if (S = 'paragraph') or (S = 'text') then
    Chars := BlockFirstLineChars(Block, FDefaultFirstLineChars)
  else
    Chars := BlockFirstLineChars(Block, 0);

  Para := AddTrackedParagraph(Chars);
  ApplyParagraphProps(Para, Block, Chars);

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

    // 记录标题，供目录书签与内部跳转使用
    SetLength(FHeadings, Length(FHeadings) + 1);
    FHeadings[High(FHeadings)].Name := '_Toc' + IntToStr(Length(FHeadings));
    FHeadings[High(FHeadings)].Text := JStr(Block, 'text', '');
    FHeadings[High(FHeadings)].Level := Level;
    FHeadings[High(FHeadings)].ParaIndex := FDoc.ParagraphCount - 1;
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
  Chars: Integer;
begin
  // runs 块不默认缩进；如需缩进可在该块写 firstLineChars
  Chars := BlockFirstLineChars(Block, 0);
  Para := AddTrackedParagraph(Chars);
  ApplyParagraphProps(Para, Block, Chars);
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
  Chars: Integer;
begin
  Style := JStr(Block, 'style', JStr(Block, 'listStyle', 'bullet'));
  Chars := BlockFirstLineChars(Block, 0);
  Items := JArray(Block, 'items');
  if Items = nil then
    Exit;
  for I := 0 to Items.Count - 1 do
  begin
    Para := AddTrackedParagraph(Chars);
    Para.ListStyle := ParseListStyle(Style);
    ApplyParagraphProps(Para, Block, Chars);
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

  Para := AddTrackedParagraph(0);
  Para.Alignment := ParseAlign(JStr(Block, 'align', 'center'),
    TParagraphAlignment.Center);
  Para.LineSpacing := TLineSpacing.Single;
  Run := Para.AddRun('');
  Run.AddImage(Data, Ext, W, H);

  // 图注（可选）
  if JStr(Block, 'caption', '') <> '' then
  begin
    Para := AddTrackedParagraph(0);
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
  Para := AddTrackedParagraph(0);
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
    AddTrackedParagraph(0)
  else
    raise EDocxTemplateError.Create('未知内容类型：' + Kind);
end;

procedure TDocxTemplate.ParseTocSettings(Root: TJSONObject);
var
  Toc: TJSONObject;
begin
  FTocEnabled := True;
  FTocTitle := '目录';
  FTocMaxLevel := 3;
  FTocPageNumbers := True;
  FTocPageBreakAfter := True;

  Toc := JObj(Root, 'toc');
  if Toc <> nil then
  begin
    if JValue(Toc, 'enabled') <> nil then
      FTocEnabled := JBool(Toc, 'enabled', True);
    FTocTitle := JStr(Toc, 'title', FTocTitle);
    FTocMaxLevel := JInt(Toc, 'maxLevel', FTocMaxLevel);
    FTocPageNumbers := JBool(Toc, 'pageNumbers', FTocPageNumbers);
    FTocPageBreakAfter := JBool(Toc, 'pageBreakAfter', FTocPageBreakAfter);
  end
  else
  begin
    if JValue(Root, 'generateToc') <> nil then
      FTocEnabled := JBool(Root, 'generateToc', True);
    FTocTitle := JStr(Root, 'tocTitle', FTocTitle);
    FTocMaxLevel := JInt(Root, 'tocMaxLevel', FTocMaxLevel);
    FTocPageNumbers := JBool(Root, 'tocPageNumbers', FTocPageNumbers);
    FTocPageBreakAfter := JBool(Root, 'tocPageBreakAfter', FTocPageBreakAfter);
  end;

  if FTocMaxLevel < 1 then
    FTocMaxLevel := 1;
  if FTocMaxLevel > 4 then
    FTocMaxLevel := 4;
end;

function TDocxTemplate.BuildTocXml: string;
var
  I, Indent: Integer;
  H: THeadingEntry;
  RunProps: string;
begin
  Result :=
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:hAnsi="Microsoft YaHei"/>' +
    '<w:b/><w:sz w:val="36"/></w:rPr><w:t>' + XmlEscape(FTocTitle) + '</w:t></w:r></w:p>';

  RunProps := '<w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:hAnsi="Microsoft YaHei"/>' +
    '<w:sz w:val="20"/></w:rPr>';

  for I := 0 to High(FHeadings) do
  begin
    H := FHeadings[I];
    if H.Level > FTocMaxLevel then
      Continue;

    Indent := 360 * (H.Level - 1);
    Result := Result +
      '<w:p><w:pPr>' +
        '<w:ind w:left="' + IntToStr(Indent) + '"/>' +
        '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs>' +
        '<w:spacing w:after="80"/>' +
      '</w:pPr>' +
      '<w:hyperlink w:anchor="' + H.Name + '" w:history="1">' +
        '<w:r>' + RunProps + '<w:t xml:space="preserve">' +
          XmlEscape(H.Text) + '</w:t></w:r>';

    if FTocPageNumbers then
    begin
      Result := Result +
        '<w:r>' + RunProps + '<w:tab/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> PAGEREF ' + H.Name +
          ' \h </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r>' + RunProps + '<w:t>1</w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    end;

    Result := Result + '</w:hyperlink></w:p>';
  end;

  if FTocPageBreakAfter then
    Result := Result + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
end;

function TDocxTemplate.InjectBookmarks(const Xml: string): string;
var
  P, StartPos, EndPos, ParaIdx, HIdx, Id: Integer;
  BookmarkStart, BookmarkEnd: string;
begin
  Result := '';
  P := 1;
  ParaIdx := 0;
  HIdx := 0;
  Id := 1;

  while HIdx <= High(FHeadings) do
  begin
    StartPos := PosEx('<w:p>', Xml, P);
    if StartPos = 0 then
      Break;
    EndPos := PosEx('</w:p>', Xml, StartPos + 5);
    if EndPos = 0 then
      Break;

    if ParaIdx = FHeadings[HIdx].ParaIndex then
    begin
      BookmarkStart := '<w:bookmarkStart w:id="' + IntToStr(Id) +
        '" w:name="' + FHeadings[HIdx].Name + '"/>';
      BookmarkEnd := '<w:bookmarkEnd w:id="' + IntToStr(Id) + '"/>';
      Result := Result + Copy(Xml, P, StartPos - P) +
        '<w:p>' + BookmarkStart +
        Copy(Xml, StartPos + 5, EndPos - (StartPos + 5)) +
        BookmarkEnd + '</w:p>';
      P := EndPos + 6;
      Inc(HIdx);
      Inc(Id);
    end
    else
    begin
      Result := Result + Copy(Xml, P, EndPos + 6 - P);
      P := EndPos + 6;
    end;
    Inc(ParaIdx);
  end;

  Result := Result + Copy(Xml, P, Length(Xml) - P + 1);
end;

{ OfficeXML4D 生成的 numbering.xml 中项目符号文本为空，修补为实心圆点，
  并补上悬挂缩进与符号字体，保证 Word / WPS 正常显示列表标记。 }
function FixNumberingXml(const Xml: string): string;
const
  LevelProps = '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/></w:rPr>';
var
  P, LvlStart, LvlEnd: Integer;
  LvlXml: string;
begin
  Result := StringReplace(Xml, '<w:lvlText w:val=""/>',
    '<w:lvlText w:val="' + #$2022 + '"/>', [rfReplaceAll]);

  P := 1;
  while True do
  begin
    LvlStart := PosEx('<w:lvl w:ilvl="', Result, P);
    if LvlStart = 0 then
      Break;
    LvlEnd := PosEx('</w:lvl>', Result, LvlStart);
    if LvlEnd = 0 then
      Break;
    LvlXml := Copy(Result, LvlStart, LvlEnd + 8 - LvlStart);
    if Pos('<w:pPr>', LvlXml) = 0 then
    begin
      Insert(LevelProps, LvlXml, Length(LvlXml) - Length('</w:lvl>') + 1);
      Result := Copy(Result, 1, LvlStart - 1) + LvlXml +
        Copy(Result, LvlEnd + 8, MaxInt);
      P := LvlStart + Length(LvlXml) + 1;
    end
    else
      P := LvlEnd + 8;
  end;
end;

function TDocxTemplate.InjectFirstLineChars(const Xml: string): string;
var
  P, StartPos, EndPos, ParaIdx, IndPos, PPrEnd, InsertAt, Chars, Hundredths: Integer;
  ParaXml, IndXml: string;
begin
  Result := '';
  P := 1;
  ParaIdx := 0;

  while ParaIdx <= High(FParaFirstLineChars) do
  begin
    StartPos := PosEx('<w:p>', Xml, P);
    if StartPos = 0 then
      Break;
    EndPos := PosEx('</w:p>', Xml, StartPos + 5);
    if EndPos = 0 then
      Break;

    Chars := FParaFirstLineChars[ParaIdx];
    if Chars <= 0 then
    begin
      Result := Result + Copy(Xml, P, EndPos + 6 - P);
    end
    else
    begin
      ParaXml := Copy(Xml, StartPos, EndPos + 6 - StartPos);
      // OOXML 的 w:firstLineChars 单位是 1/100 字符：2 字符要写 200
      Hundredths := Chars * 100;
      IndXml := '<w:ind w:firstLineChars="' + IntToStr(Hundredths) + '"/>';
      PPrEnd := PosEx('</w:pPr>', ParaXml);
      IndPos := PosEx('<w:ind', ParaXml);

      if (IndPos > 0) and ((PPrEnd = 0) or (IndPos < PPrEnd)) then
      begin
        // 库已写入 w:firstLine（twips 兜底），补上字符数属性；Word 优先采用它
        InsertAt := IndPos + Length('<w:ind');
        ParaXml := Copy(ParaXml, 1, InsertAt - 1) +
          ' w:firstLineChars="' + IntToStr(Hundredths) + '"' +
          Copy(ParaXml, InsertAt, MaxInt);
      end
      else if PPrEnd > 0 then
        ParaXml := Copy(ParaXml, 1, PPrEnd - 1) + IndXml +
          Copy(ParaXml, PPrEnd, MaxInt)
      else
        ParaXml := '<w:p><w:pPr>' + IndXml + '</w:pPr>' +
          Copy(ParaXml, Length('<w:p>') + 1, MaxInt);

      Result := Result + Copy(Xml, P, StartPos - P) + ParaXml;
    end;

    P := EndPos + 6;
    Inc(ParaIdx);
  end;

  Result := Result + Copy(Xml, P, Length(Xml) - P + 1);
end;

function TDocxTemplate.SaveWithToc(const FileName: string): Boolean;
var
  Mem: TMemoryStream;
  ZipIn, ZipOut: TZipFile;
  FS: TFileStream;
  Names: TArray<string>;
  Contents: TArray<TBytes>;
  I: Integer;
  Name, DocXml, RelsXml, CTXml, NumberingXml: string;
  Data: TBytes;
  NewRel: string;
  HasSettings, NeedsToc, NeedsIndent, NeedsNumbering: Boolean;
begin
  NeedsToc := FTocEnabled and (Length(FHeadings) > 0);
  NeedsIndent := False;
  for I := 0 to High(FParaFirstLineChars) do
    if FParaFirstLineChars[I] > 0 then
    begin
      NeedsIndent := True;
      Break;
    end;

  Mem := TMemoryStream.Create;
  ZipIn := TZipFile.Create;
  try
    FDoc.SaveToStream(Mem);
    Mem.Position := 0;
    ZipIn.Open(Mem, zmRead);

    SetLength(Names, ZipIn.FileCount);
    SetLength(Contents, ZipIn.FileCount);
    DocXml := '';
    RelsXml := '';
    CTXml := '';
    NumberingXml := '';
    HasSettings := False;

    for I := 0 to ZipIn.FileCount - 1 do
    begin
      Names[I] := ZipIn.FileNames[I];
      ZipIn.Read(I, Data);
      Contents[I] := Data;
      Name := Names[I];
      if Name = SettingsPart then
        HasSettings := True
      else if Name = 'word/document.xml' then
        DocXml := TEncoding.UTF8.GetString(Data)
      else if Name = 'word/_rels/document.xml.rels' then
        RelsXml := TEncoding.UTF8.GetString(Data)
      else if Name = 'word/numbering.xml' then
        NumberingXml := TEncoding.UTF8.GetString(Data)
      else if Name = '[Content_Types].xml' then
        CTXml := TEncoding.UTF8.GetString(Data);
    end;
  finally
    ZipIn.Free;
    Mem.Free;
  end;

  // 1) 给需要首行缩进的段落写入 w:firstLineChars（按字符宽自适应中英文）
  if NeedsIndent then
    DocXml := InjectFirstLineChars(DocXml);

  // 2) 给标题段落加书签，并在正文开头插入目录页
  if NeedsToc then
  begin
    DocXml := InjectBookmarks(DocXml);
    DocXml := StringReplace(DocXml, '<w:body>',
      '<w:body>' + BuildTocXml, []);
  end;

  // 3) 修补库生成的空项目符号，让列表显示实心圆点/编号
  NeedsNumbering := (NumberingXml <> '') and (Pos('<w:numPr>', DocXml) > 0);
  if NeedsNumbering then
    NumberingXml := FixNumberingXml(NumberingXml);

  // 4) 页码字段需要 Word 打开时更新
  if NeedsToc and FTocPageNumbers then
  begin
    if Pos('PartName="/word/settings.xml"', CTXml) = 0 then
      CTXml := StringReplace(CTXml, '</Types>',
        '<Override PartName="/word/settings.xml" ContentType="' +
        SettingsContentType + '"/></Types>', []);

    if Pos(SettingsRelationshipType, RelsXml) = 0 then
    begin
      NewRel := '<Relationship Id="rId' + IntToStr(NextRelationshipId(RelsXml)) +
        '" Type="' + SettingsRelationshipType + '" Target="settings.xml"/>';
      RelsXml := StringReplace(RelsXml, '</Relationships>',
        NewRel + '</Relationships>', []);
    end;
  end;

  for I := 0 to High(Names) do
  begin
    if Names[I] = 'word/document.xml' then
      Contents[I] := TEncoding.UTF8.GetBytes(DocXml)
    else if Names[I] = 'word/_rels/document.xml.rels' then
      Contents[I] := TEncoding.UTF8.GetBytes(RelsXml)
    else if Names[I] = '[Content_Types].xml' then
      Contents[I] := TEncoding.UTF8.GetBytes(CTXml)
    else if Names[I] = 'word/numbering.xml' then
      Contents[I] := TEncoding.UTF8.GetBytes(NumberingXml);
  end;

  FS := TFileStream.Create(FileName, fmCreate);
  ZipOut := TZipFile.Create;
  try
    ZipOut.Open(FS, zmWrite);
    for I := 0 to High(Names) do
      ZipOut.Add(Contents[I], Names[I]);

    if NeedsToc and FTocPageNumbers and (not HasSettings) then
      ZipOut.Add(
        TEncoding.UTF8.GetBytes(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w:settings xmlns:w="' + WordNs + '">' +
          '<w:updateFields w:val="true"/>' +
          '</w:settings>'),
        SettingsPart);
  finally
    ZipOut.Free;
    FS.Free;
  end;
  Result := True;
end;

procedure TDocxTemplate.ApplyRootSettings(Root: TJSONObject);
var
  Page, ParaCfg: TJSONObject;
begin
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

  // 段落首行缩进：默认两个字符，可按全局配置覆盖或关闭
  FDefaultFirstLineChars := 2;
  ParaCfg := JObj(Root, 'paragraph');
  if ParaCfg <> nil then
  begin
    if JValue(ParaCfg, 'firstLineChars') <> nil then
      FDefaultFirstLineChars := JInt(ParaCfg, 'firstLineChars', 2)
    else if JValue(ParaCfg, 'firstLineIndentChars') <> nil then
      FDefaultFirstLineChars := JInt(ParaCfg, 'firstLineIndentChars', 2)
    else if JValue(ParaCfg, 'indentChars') <> nil then
      FDefaultFirstLineChars := JInt(ParaCfg, 'indentChars', 2);
    if (JValue(ParaCfg, 'enabled') <> nil) and
       (not JBool(ParaCfg, 'enabled', True)) then
      FDefaultFirstLineChars := 0;
  end
  else
  begin
    if JValue(Root, 'firstLineChars') <> nil then
      FDefaultFirstLineChars := JInt(Root, 'firstLineChars', 2)
    else if JValue(Root, 'firstLineIndentChars') <> nil then
      FDefaultFirstLineChars := JInt(Root, 'firstLineIndentChars', 2);
  end;
  if FDefaultFirstLineChars < 0 then
    FDefaultFirstLineChars := 0;
  if FDefaultFirstLineChars > 20 then
    FDefaultFirstLineChars := 20;

  ParseTocSettings(Root);
end;

procedure TDocxTemplate.RenderContent(Root: TJSONObject);
var
  Content: TJSONArray;
  I: Integer;
begin
  Content := JArray(Root, 'content');
  if Content = nil then
    raise EDocxTemplateError.Create('配置缺少 content 数组');

  for I := 0 to Content.Count - 1 do
    if Content.Items[I] is TJSONObject then
      RenderBlock(TJSONObject(Content.Items[I]));
end;

procedure TDocxTemplate.LoadFromJsonText(const AJsonText, ABaseDir: string);
var
  Text: string;
  Root: TJSONObject;
  JSON: TJSONValue;
begin
  Text := Trim(AJsonText);
  if Text = '' then
    Exit;
  if Text[1] = #$FEFF then
    Delete(Text, 1, 1);

  JSON := TJSONObject.ParseJSONValue(Text);
  if JSON = nil then
    raise EDocxTemplateError.Create('JSON 解析失败：' + Copy(Text, 1, 200));
  try
    if not (JSON is TJSONObject) then
      raise EDocxTemplateError.Create('配置根节点必须是 JSON 对象');
    Root := TJSONObject(JSON);

    FBaseDir := ABaseDir;
    FDoc := TWordDocumentFactory.CreateDocument;
    FHeadings := nil;
    FParaFirstLineChars := nil;
    FDefaultFirstLineChars := 2;

    ApplyRootSettings(Root);
    RenderContent(Root);
  finally
    JSON.Free;
  end;
end;

procedure TDocxTemplate.LoadConfig(const FileName: string);
var
  Text: string;
begin
  if not TFile.Exists(FileName) then
    raise EDocxTemplateError.Create('配置文件不存在：' + FileName);

  Text := TFile.ReadAllText(FileName, TEncoding.UTF8);
  LoadFromJsonText(Text, TPath.GetDirectoryName(TPath.GetFullPath(FileName)));
end;

procedure TDocxTemplate.LoadConfig(const AJsonText, ABaseDir: string);
begin
  LoadFromJsonText(AJsonText, ABaseDir);
end;

procedure TDocxTemplate.SaveTo(const FileName: string);
var
  Dir: string;
begin
  Dir := TPath.GetDirectoryName(TPath.GetFullPath(FileName));
  if (Dir <> '') and (not TDirectory.Exists(Dir)) then
    TDirectory.CreateDirectory(Dir);

  if not SaveWithToc(FileName) then
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

class function TDocxTemplate.Generate(const AJsonText, ABaseDir,
  OutputFile: string): IWordDocument;
var
  Tpl: TDocxTemplate;
begin
  Tpl := TDocxTemplate.Create;
  try
    Tpl.LoadConfig(AJsonText, ABaseDir);
    Tpl.SaveTo(OutputFile);
    Result := Tpl.Document;
  finally
    Tpl.Free;
  end;
end;

end.