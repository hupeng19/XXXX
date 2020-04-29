unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    Label1: TLabel;
    Edit1: TEdit;
    Button1: TButton;
    Button2: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
    XXX: Boolean;
    function GetName: string;
  public
    { Public declarations }
    property Name: string read GetName;
    property Modalresult: Boolean read XXX;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
  XXX := True;
  Self.Close;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  XXX := False;
  Self.Close;
end;

function TForm1.GetName: string;
begin
  Result := Edit1.Text;
end;

end.
