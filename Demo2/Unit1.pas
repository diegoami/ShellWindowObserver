unit Unit1;

{*****************************************************************************
 *
 *  Unit.pas - Shell Window Observer Component Demo
 *
 *  Copyright (c) 2000 Diego Amicabile
 *
 *  Author:     Diego Amicabile
 *  E-mail:     diegoami@yahoo.it
 *  Homepage:   http://www.geocities.com/diegoami
 *
 *  This component is free software; you can redistribute it and/or
 *  modify it under the terms of the GNU General Public License
 *  as published by the Free Software Foundation;
 *
 *  This component is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this component; if not, write to the Free Software
 *  Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA
 *
 *****************************************************************************}


interface
{$DEFINE COOBJS}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  uShellWindowObserver, ComCtrls, ExtCtrls, StdCtrls, ComObj, SHDocVW;

type
  TForm1 = class(TForm)
    ListView1: TListView;
    ShellWindowObserver1: TShellWindowObserver;
    Splitter1: TSplitter;
    URLEdit: TEdit;
    GoButton: TButton;
    NewButton: TButton;
    Memo1: TMemo;
    Panel1: TPanel;
    procedure ShellWindowObserver1AddedEntry(Sender: TObject;
      ShellWindow: TShellWindow; Str: String);
    procedure ShellWindowObserver1ChangedNumber(Sender: TObject;
      Number: Integer);
    procedure GoButtonClick(Sender: TObject);
    procedure NewButtonClick(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
  private
    procedure UpdateListView;
{$IFDEF COOBJS}
    function GetIEAtListItem(LI : TListItem) : IWebBrowser2;
{$ELSE }
    function GetIEAtListItem(LI : TListItem) : Variant;
{$ENDIF }
  public


  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}
var
{$IFDEF COOBJS}
  IET : IWebBrowser2;
{$ELSE}
  IET : Variant;
{$ENDIF}


{$IFDEF COOBJS}
function TForm1.GetIEAtListItem(LI : TListItem) : IWebBrowser2;
var IEC : IWebBrowser2;
{$ELSE }
function TForm1.GetIEAtListItem(LI : TListItem) : Variant;
var IEC : Variant;
{$ENDIF }
  SHW : TShellWindow;
begin
  SHW := TShellWindow(LI.Data );
  IEC := SHW.IEHandle;
  result := IEC
end;

procedure TForm1.ShellWindowObserver1AddedEntry(Sender: TObject;
  ShellWindow: TShellWindow; Str: String);
begin
  UpdateListView;
    Memo1.Lines.Add(Str)

end;

procedure TForm1.UpdateListView;
var i : integer;
  SHW : TShellWindow;
  LI : TListItem;
begin
  ListView1.Items.Clear;
  with ShellWindowObserver1.WindowList do begin
    for i := 0 to Count-1 do begin
      SHW := TShellWindow(Items[i]);
      LI := ListView1.Items.Add;
      LI.Caption := IntToStr(SHW.Num);
      LI.SubItems.Add(SHW.LocationUrl);
      LI.SubItems.Add(SHW.LocationName);
      LI.Data := SHW
    end
  end;
end;

procedure TForm1.ShellWindowObserver1ChangedNumber(Sender: TObject;
  Number: Integer);
begin
  UpdateListView;
  Memo1.Lines.Add('Now '+IntToStr(Number)+ ' instances running. ');

end;

procedure TForm1.GoButtonClick(Sender: TObject);
var LI : TListItem;
  ET : OleVariant;
begin
  LI := ListView1.Selected;
  if LI<>nil then begin
    IET := GetIEAtListItem(LI );
    IET.Navigate(URLEdit.Text{$IFDEF COOBJS},ET,ET,ET,ET{$ENDIF})
  end;
end;






procedure TForm1.NewButtonClick(Sender: TObject);
var ET :OleVariant;
begin

 {$IFDEF COOBJS}    IET := CoInternetExplorer.Create;
{$ELSE }
    IET := CreateComObject(Class_InternetExplorer) as IWebBrowser2;

{$ENDIF }
  IET.Visible := True;
  if URLEdit.Text <> '' then
    IET.Navigate(URLEDit.Text{$IFDEF COOBJS},ET,ET,ET,ET{$ENDIF})
  else if ListView1.Selected <> nil then
    IET.Navigate(TShellWindow(ListView1.Selected.Data).LocationURL{$IFDEF COOBJS},ET,ET,ET,ET{$ENDIF});
end;

procedure TForm1.CloseButtonClick(Sender: TObject);
var LI : TListItem;
  SHW : TShellWindow;
begin
  LI := ListView1.Selected;
  if LI<>nil then begin

 //   DestroyWindow(TShellWIndow(LI.Data).handle);
  end;
end;

end.
