program Test;

uses
  Vcl.Forms,
  Main in 'Main.pas' {FMain},
  MyInterface in 'MyInterface.pas',
  CommonProcedure in 'CommonProcedure.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFMain, FMain);
  Application.Run;
end.
