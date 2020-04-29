unit Unit4;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters, dxCore, dxCoreClasses,
  dxHashUtils, dxSpreadSheetCore, dxSpreadSheetCoreHistory, dxSpreadSheetConditionalFormatting,
  dxSpreadSheetConditionalFormattingRules, dxSpreadSheetClasses, dxSpreadSheetContainers,
  dxSpreadSheetFormulas, dxSpreadSheetHyperlinks, dxSpreadSheetFunctions, dxSpreadSheetGraphics,
  dxSpreadSheetPrinting, dxSpreadSheetTypes, dxSpreadSheetUtils, dxBarBuiltInMenu, Menus,
  dxSpreadSheet, StdCtrls, ExtCtrls, Generics.Collections, unit1, ShellAPI, ComCtrls;

type
  TForm4 = class(TForm)
    dxSpreadSheet1: TdxSpreadSheet;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    Panel1: TPanel;
    Button1: TButton;
    OpenDialog1: TOpenDialog;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    ColorDialog1: TColorDialog;
    N5: TMenuItem;
    N6: TMenuItem;
    TreeView1: TTreeView;
    procedure N1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure TreeView1DblClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    ActiveSheet: TdxSpreadSheetTableView;
    CurItem: TList<TdxSpreadSheetTableRow>;
    NewItem: TList<TdxSpreadSheetTableRow>;
    WHos: TList<string>;
    SelectedItems: TList<TList<TdxSpreadSheetTableRow>>;
    NeedCheck: Boolean;
    SelectingColor, SelectedColor: TColor;
    AutoOpenPage: Boolean;
    procedure CMChildKey(var AMsg: TCMChildKey); message CM_CHILDKEY;
  end;

var
  Form4: TForm4;

implementation

{$R *.dfm}

procedure TForm4.Button1Click(Sender: TObject);
var
  I: Integer;
  J: Integer;
  S: string;
  Node: TTreeNode;
  Nameadd: Boolean;
begin
  Nameadd := False;
  if NeedCheck then
  begin
    with TForm1.Create(nil) do
    begin
      ShowModal;
      if not ModalResult then
        Exit;
      if WHos.Contains(Name) then
      begin
        ShowMessage(Name+'已经添加过');
        Exit
      end
      else
      begin
        Nameadd := True;
        WHos.Add(Name);
      end;
    end;
  end;

  for I := 0 to NewItem.Count - 1 do
  begin
    NewItem[I].Style.Brush.BackgroundColor := SelectedColor;
  end;
  if dxSpreadSheet1.ActiveSheet <> ActiveSheet then
  begin
    CurItem.Clear;
    if dxSpreadSheet1.ActiveSheet is TdxSpreadSheetTableView then
    begin
      ActiveSheet := dxSpreadSheet1.ActiveSheet as TdxSpreadSheetTableView;
      for I := 1 to ActiveSheet.Rows.Count - 1 do
      begin
        if ActiveSheet.Rows[I].Style.Brush.BackgroundColor <> clRed then
        begin
          if TryStrToInt(ActiveSheet.Cells[I, 0].DisplayText, J) then
          begin
            CurItem.Add(ActiveSheet.Rows[I]);
          end;
        end;
      end;
    end;
  end;
  if CurItem.Count < 1 then
  begin
    ShowMessage('已经抽完');
    Exit;
  end;
  NewItem := TList<TdxSpreadSheetTableRow>.Create;
  SelectedItems.Add(NewItem);
  for I := 0 to 3 - 1 do
  begin
    if CurItem.Count < 1 then
      Break;
    J := Random(CurItem.Count);
    NewItem.Add(CurItem[J]);
    CurItem.Delete(J);
  end;
  OutputDebugString('----------------------');
  for I := 0 to NewItem.Count - 1 do
  begin
    NewItem[I].Style.Brush.BackgroundColor := SelectingColor;
    if AutoOpenPage then
    begin
      S := 'http://pms.fsdev.cn/index.php?m=story&f=view&t=html&id=' + NewItem[I].Cells[0].DisplayText;
      ShellExecute(0, '', PChar(S), nil, nil, SW_SHOWNORMAL);
    end;
  end;

  //
  if Nameadd then
    Node := TreeView1.Items.AddObject(nil, WHos.Last, 0)
  else
    Node := TreeView1.Items.AddObject(nil, '__', 0);
  for I := 0 to NewItem.Count - 1 do
  begin
    S := 'http://pms.fsdev.cn/index.php?m=story&f=view&t=html&id=' + NewItem[I].Cells[0].DisplayText;
    TreeView1.Items.AddChildObject(Node, NewItem[I].Cells[0].DisplayText, PChar(S));
  end;
end;

procedure TForm4.CMChildKey(var AMsg: TCMChildKey);
var
  I: Integer;
begin
  I := GetKeyState(VK_F1);
  if I < 0 then
  begin
    Button1.Click;
  end;
end;

procedure TForm4.FormCreate(Sender: TObject);
begin
  ActiveSheet := nil;
  CurItem := TList<TdxSpreadSheetTableRow>.Create;
  NewItem := TList<TdxSpreadSheetTableRow>.Create;
  SelectedItems := TList<TList<TdxSpreadSheetTableRow>>.Create;
  WHos := TList<string>.Create;
  NeedCheck := True;
  SelectingColor := clYellow;
  SelectedColor := clRed;
  AutoOpenPage := N6.Checked;
end;

procedure TForm4.N1Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
  begin
    dxSpreadSheet1.LoadFromFile(OpenDialog1.FileName);
    ActiveSheet := nil;
    CurItem.Clear;
    NewItem.Clear;
    WHos.Clear;
    SelectedItems.Clear;
    TreeView1.Items.Clear;
  end;
//  dxSpreadSheet1.LoadFromFile('C:\Users\HuuPPP\Documents\WXWork\1688853499024982\Cache\File\2020-04\需求汇总表(1).xlsx');
end;

procedure TForm4.N3Click(Sender: TObject);
begin
  NeedCheck := N3.Checked;
end;

procedure TForm4.N4Click(Sender: TObject);
var
  I: Integer;
begin
  with ColorDialog1 do
  begin
    Caption := N4.Caption;
    if Execute then
    begin
      if SelectingColor <> Color then
      begin
        SelectingColor := Color;
      end;
    end;
  end;
  Caption := '抽签';
end;

procedure TForm4.N5Click(Sender: TObject);
var
  I: Integer;
begin
  with ColorDialog1 do
  begin
    Caption := N5.Caption;
    if Execute then
    begin
      if SelectingColor <> Color then
      begin
        SelectedColor := Color;
      end;
    end;
  end;
  Caption := '抽签';
end;

procedure TForm4.N6Click(Sender: TObject);
begin
  AutoOpenPage := N6.Checked;
end;

procedure TForm4.TreeView1DblClick(Sender: TObject);
begin
  if TreeView1.Selected <> nil then
  begin
    if TreeView1.Selected.Data <> nil then
    begin
      ShellExecute(0, '', PChar(TreeView1.Selected.Data), nil, nil, SW_SHOWNORMAL);
    end;
  end;
end;

end.
