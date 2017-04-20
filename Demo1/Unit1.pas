unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls,  uShellWindowObserver;

type
  TForm1 = class(TForm)
    ShellWindowObserver1: TShellWindowObserver;
    Memo1: TMemo;
    procedure ShellWindowObserver1AddedEntry(Sender: TObject;
      ShellWindow: TShellWindow; Str : String);
    procedure FormShow(Sender: TObject);
    procedure ShellWindowObserver1ChangedNumber(Sender: TObject;
      Number: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.ShellWindowObserver1AddedEntry(Sender: TObject;
  ShellWindow: TShellWindow; Str : String);
  var prova :string;
  begin
    Memo1.Lines.Add(Str)
end;

procedure TForm1.FormShow(Sender: TObject);
var prova, oriv : String;
begin
  Memo1.Lines.Clear;
  prova := 'A%20BB%20jgoer%20BBERF';
  oriv := ReplaceStr(prova,'%20',' ');
  //ShowMessage(oriv);
end;

procedure TForm1.ShellWindowObserver1ChangedNumber(Sender: TObject;
  Number: Integer);
begin
  Memo1.Lines.Add('Now '+IntToStr(Number)+ ' instances running. ');
end;

end.
