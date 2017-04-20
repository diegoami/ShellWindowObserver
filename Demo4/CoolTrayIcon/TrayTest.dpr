program TrayTest;

uses
  Forms,
  WinProcs,
  TiMain in 'TiMain.pas' {MainForm},
  CoolTrayIcon in 'CoolTrayIcon.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'TrayIcon Test';
{
In stead of using StartMinimized you can insert these two lines.
This code will be executed immediately, while StartMinimized will
be executed when the TCoolTrayIcon component is loaded.
However, it makes no difference because the main form is not shown
before Application.Run is executed.

  Application.ShowMainForm := False;
  ShowWindow(Application.Handle, SW_HIDE);
}
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
