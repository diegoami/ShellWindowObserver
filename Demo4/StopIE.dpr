program StopIE;

uses
  Forms,
  FormStopIE in 'FormStopIE.pas' {frmStopIE};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Stop Internet Explorer';
  Application.CreateForm(TfrmStopIE, frmStopIE);
  Application.Run;
end.
