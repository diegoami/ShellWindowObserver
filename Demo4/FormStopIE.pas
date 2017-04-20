unit FormStopIE;

{*****************************************************************************
 *
 *  uShellWindowObserver.pas demo - Shell Window Observer Component - Version 1.2
 *
 *  Created by Randy Williams
 *****************************************************************************}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  uShellWindowObserver, ComCtrls, StdCtrls, Menus, CoolTrayIcon, ExtCtrls,
  Buttons;

type
  TfrmStopIE = class(TForm)
    ShellWindowObserver1: TShellWindowObserver;
    ListView1: TListView;
    StatusBar1: TStatusBar;
    CoolTrayIcon1: TCoolTrayIcon;
    PopupMenu1: TPopupMenu;
    Restore1: TMenuItem;
    Exit1: TMenuItem;
    Timer1: TTimer;
    Panel1: TPanel;
    Image1: TImage;
    bClose: TBitBtn;
    bMin: TBitBtn;
    procedure UpdateListView;
    procedure ShellWindowObserver1AddedEntry(Sender: TObject;
      ShellWindow: TShellWindow; Str: String);
    procedure ShellWindowObserver1ChangedNumber(Sender: TObject;
      Number: Integer);
    procedure FormCreate(Sender: TObject);
    procedure bCloseClick(Sender: TObject);
    procedure ListView1SelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure Restore1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmStopIE: TfrmStopIE;
  WindowCount: integer;

implementation

{$R *.DFM}

procedure TfrmStopIE.UpdateListView;
var i : integer;
  SHW : TShellWindow;
  LI : TListItem;
  TempWindowCount: integer;
  _vPlural: string;
begin
  TempWindowCount := WindowCount;
  ListView1.Items.Clear;
  with ShellWindowObserver1.WindowList do
  begin
    WindowCount := Count;
    If Count = 1 then _vPlural := '' else _vPlural := 's';
    StatusBar1.SimpleText := IntToStr(Count)+
    ' Internet Explorer window'+_vPlural+' running.';
    for i := 0 to Count-1 do
    begin
      SHW := TShellWindow(Items[i]);
      LI := ListView1.Items.Add;
      LI.Caption := IntToStr(SHW.Num);
      LI.SubItems.Add(SHW.LocationUrl);
      LI.SubItems.Add(SHW.LocationName);
      LI.Data := SHW;
    end;
    ListView1.Selected := ListView1.Items[Count-1];
    WindowCount := Count;
  end;
  bClose.Enabled := (ListView1.Selected <> nil);
  if WindowCount > TempWindowCount then
  begin
    Application.Restore;
    bClose.SetFocus;
    Timer1.Enabled := True;
  end;
end;

procedure TfrmStopIE.ShellWindowObserver1AddedEntry(Sender: TObject;
  ShellWindow: TShellWindow; Str: String);
begin
  UpdateListView;
end;

procedure TfrmStopIE.ShellWindowObserver1ChangedNumber(Sender: TObject;
  Number: Integer);
begin
  UpdateListView;
end;

procedure TfrmStopIE.FormCreate(Sender: TObject);
begin
  WindowCount := 0;
  UpdateListView;
end;

procedure TfrmStopIE.bCloseClick(Sender: TObject);
var
  LI : TListItem;
begin
  Timer1.Enabled := False;
  LI := ListView1.Selected;
  if LI <> nil then
   SendMessage(TShellWIndow(LI.Data).handle, WM_SYSCOMMAND, SC_CLOSE, 0);
  Application.Minimize;
end;

procedure TfrmStopIE.ListView1SelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
  bClose.Enabled := (ListView1.Selected <> nil);
end;

procedure TfrmStopIE.Restore1Click(Sender: TObject);
begin
  Application.Restore;
end;

procedure TfrmStopIE.Exit1Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmStopIE.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  Application.Minimize;
end;

end.
