program TestRegistration;

uses
  Vcl.Forms,
  UnitRegistration in 'UnitRegistration.pas' {Form1},
  CommonProcedure in 'CommonProcedure.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
