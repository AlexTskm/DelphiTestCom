unit UnitRegistration;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    BtnOpen: TButton;
    BtnReg: TButton;
    BtnExit: TButton;
    Panel2: TPanel;
    Result: TLabel;
    ListBox1: TListBox;
    OpenDialog1: TOpenDialog;
    procedure BtnOpenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BtnExitClick(Sender: TObject);
    procedure BtnRegClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses ShellAPI, CommonProcedure, MyInterface;

function IsUserAnAdmin(): BOOL; external shell32;

{$R *.dfm}

procedure TForm1.BtnExitClick(Sender: TObject);
begin
 self.Close;
end;

procedure TForm1.BtnOpenClick(Sender: TObject);
var
  i:integer;
begin
 if OpenDialog1.Execute then
 begin
  ListBox1.Items.Clear;
  for i:=0 to OpenDialog1.Files.Count-1 do
  begin
   ListBox1.Items.Add(OpenDialog1.Files.Strings[i]);
  end;
  if OpenDialog1.Files.Count>0 then
    BtnReg.Enabled:=true;
 end;
end;

procedure TForm1.BtnRegClick(Sender: TObject);
var
 i:integer;
begin
 for i:=0 to ListBox1.Items.Count-1 do
 begin
  if not RegFilePr(Self.Handle, ListBox1.Items[i]) then
  begin
    ShowMessage('Ошибка регистрации COM обьекта:'+ListBox1.Items[i]);
    exit;
  end;
 end;
 SaveFromRegistr(ThePathToTheApplication, Class_Of_TIncTheDate);
 SaveFromRegistr(ThePathToTheApplication, Class_Of_TSquareExtraction);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  if not IsUserAnAdmin then
  begin
   ShowMessage('Запустите программу от имени администратора');
   BtnOpen.Enabled:=false;
  end;
  OpenDialog1.InitialDir := GetCurrentDir;
  Position:=poDesktopCenter;
end;

end.
