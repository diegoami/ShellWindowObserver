unit uShellWindowObserver;

{*****************************************************************************
 *
 *  uShellWindowObserver.pas - Shell Window Observer Component - Version 1.1
 *
 *****************************************************************************
 *  Created by Diego Amicabile
 *  Updated by Michel Hibon - 08/27/2001
 *****************************************************************************
 *
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

{*****************************************************************************
 *
 * requirements : The Internet Explorer Active X Component must have been
 * installed
 *****************************************************************************}

 interface
//remove the following line if you are using Delphi 3
{$DEFINE COOBJS}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  extctrls, SHDocVW
  // replace this with SHDocVW_TLB if you installed it as an activeX component
  , Comobj, comctrls;

type
  TShellMessage = (smClose,smMaximize,smMinimize,smRestore);
  EIdNotFound = class(Exception);
  TShellWindow = class
    Num : integer;
    handle : HWND;
    LocationURL : String;
    LocationName : String;
    {$IFDEF COOBJS}
      IEHandle : IWebBrowser2;
    {$ELSE }
      IEHandle : Variant;
    {$ENDIF }
  end;

  TOnAddedEntry = procedure (Sender : TObject; ShellWindow : TShellWindow; Str : String) of object;
  TOnChangedNumber = procedure (Sender : TObject; Number : Integer) of object;

  TShellWindowObserver = class(TComponent)
  private
      FOnAddedEntry : TOnAddedEntry;
      FLoggingDirs : Boolean;
      FLoggedChanges : TStrings;
      FWindowList : TList;
      FOnChangedNUmber : TOnChangedNumber;
      FDelimiter : Char;
      FActive : Boolean;
      FShellMessage : TShellMessage;
      FNumMessage: WORD;
      procedure OnTimerObserver(Sender : TObject);
      procedure SetMessage(Value: TShellMessage);
   protected
      FTimer : TTimer;
      function GetMaxId : integer;
      procedure RemoveZombies;
      procedure UpdateList;
      procedure AddEntry(SHW : TSHellWindow);
      procedure AddBrowserId(URL, Name : String; curriehandle : HWND);
      procedure UpdateWins;
   public
     Count: Integer;
     function    SendMessageToWindowByNumber(Num: Integer): LRESULT;
     function    GetShellWindowOnNumber(Num : integer) : TShellWindow;
     constructor Create(AOwner : TComponent); override;
     destructor  Destroy; override;
     property    WindowList : TList read FWindowList write FWindowList;
     property    LoggedChanges : TStrings read FLoggedChanges;
   published
      property LoggingDirs : boolean read FLoggingDirs write FLoggingDirs;
      property OnAddedEntry : TOnAddedEntry read FOnAddedEntry write FOnAddedEntry;
      property OnChangedNumber : TOnChangedNumber read FOnChangedNumber write FOnChangedNumber;
      property Delimiter : Char read FDelimiter write FDelimiter;
      property Active : boolean read FActive write FActive;
      property SendAMessage : TShellMessage read FShellMessage write SetMessage;
  end;

function ReplaceStr(str : String;old,new:string):string;
procedure Register;

implementation



{$R SHELLWINDOWOBSERVER.DCR}
const
      TIMERINTERVAL = 75;

var
{$IFDEF COOBJS}
   SH : IShellWindows;
   IE : IWebBrowser2;
{$ELSE }
   SH : Variant;
   IE : Variant;
{$ENDIF}
   initializ : boolean = false;

procedure Init;
begin
  if not initializ then begin
 {$IFDEF COOBJS}
    SH := CoShellWindows.Create;
{$ELSE }
    SH := CreateComObject(Class_ShellWindows) as IShellWindows;
{$ENDIF }
  end;
  initializ := true;
end;

//----------------------------------------------------------------------------//
//                                     max                                    //
//----------------------------------------------------------------------------//
function max(a,b :integer) : integer;
begin
  if a > b then result := a else result := b;
end;

//----------------------------------------------------------------------------//
//                                 ReplaceStr                                 //
//----------------------------------------------------------------------------//
function ReplaceStr(str : STring;old,new:string):string;
var
  P : integer;
  S : String;
  first, last : String;

begin
  P := Pos(Old,str);
  if P > 0 then
    begin
      first:=Copy(str,1,P-1);
      last:=copy(str,P+Length(old),Length(str)-P-Length(Old)+1);
      S := first + new + Replacestr(last,old,new);
      result := S;
    end
  else
    result := str;
end;

//----------------------------------------------------------------------------//
//                                  SetMessage                                //
//----------------------------------------------------------------------------//
procedure TShellWindowObserver.SetMessage(Value: TShellMessage);
begin
  if Value <> FShellMessage then
    begin
      Case Value of
        smClose       : FNumMessage := SC_CLOSE;
        smMaximize    : FNumMessage := SC_MAXIMIZE;
        smMinimize    : FNumMessage := SC_MINIMIZE;
        smRestore     : FNumMessage := SC_RESTORE;
      end;
      FShellMessage := Value;
    end;
end;

//----------------------------------------------------------------------------//
//                          SendMessageToWindowByNumber                       //
//----------------------------------------------------------------------------//
function TShellWindowObserver.SendMessageToWindowByNumber(Num: Integer): LRESULT;
begin
  Result := SendMessage(TShellWindow(FWindowList.Items[Pred(Num)]).handle,
                        WM_SYSCOMMAND,FNumMessage,0);
end;

//----------------------------------------------------------------------------//
//                            GetShellWindowOnNumber                          //
//----------------------------------------------------------------------------//
function TShellWindowObserver.GetShellWindowOnNumber(Num : Integer) : TShellWindow;
var
  i : integer;
  CurrWin : TShellWindow;

begin
  result := nil;
  for i := 0 to Pred(FWindowList.Count) do
    begin
      CurrWin := TShellWindow(FWindowList.Items[i]);
      if CurrWin.Num = Num then
        result := CurrWin
    end;
end;

//----------------------------------------------------------------------------//
//                                  GetMaxId                                  //
//----------------------------------------------------------------------------//
function TShellWindowObserver.GetMaxId : integer;
var
  i : integer;
  maxfound : integer;

begin
  maxfound := 0;
  for i := 0 to FWindowList.Count - 1 do
    maxfound := max(Maxfound, TShellWindow(FWIndowList.Items[i]).Num);
  result := maxfound + 1
end;

//----------------------------------------------------------------------------//
//                                  AddEntry                                  //
//----------------------------------------------------------------------------//
procedure TShellWindowObserver.AddEntry(SHW : TShellWindow);
var
  NewStr : String;

begin
  NewStr := IntTostr(SHW.Num) + FDELIMITER + '-' + FDELIMITER +
            'Hwnd' + FDELIMITER + ':' + FDELIMITER +
            IntToHex(SHW.handle,8) + FDELIMITER +
            SHW.LocationURL + FDELIMITER +
            SHW.LocationName + FDELIMITER +
            DateTimeToStr(Now);
  FLoggedChanges.Add(NewStr);
  Count := FWindowList.Count;
  if Assigned(FOnAddedEntry) then
    FOnAddedEntry(Self, SHW, NewStr);
end;


//----------------------------------------------------------------------------//
//                                AddBrowserId                                //
//----------------------------------------------------------------------------//
procedure TShellWindowObserver.AddBrowserId(URL, Name : String;
                                            curriehandle : HWND);
var
  NewId : integer;
  ShellWindow : TShellWindow;

begin
  NewId := GetMaxId;
  ShellWindow := TShellWindow.Create;
  with ShellWindow do
    begin
      handle := CurrIeHandle;
      Num := NewId;
      LocationURL := URL;
      LocationName := Name;
      IEHandle := IE
    end;
  FWindowList.ADD(ShellWindow);
  Count := FWindowList.Count;
  if Assigned(FOnChangedNumber) then
    FOnChangedNumber(Self, FWindowList.Count);
  AddEntry(ShellWindow);
end;

//----------------------------------------------------------------------------//
//                                 UpdateWins                                 //
//----------------------------------------------------------------------------//
procedure TShellWindowObserver.UpdateWins;

  function IsValidUrl(URL : String) : boolean;
  begin
    result := false;
    if Pos('http',URL) > 0 then result := true;
    if Pos('ftp',URL) > 0 then result := true;
    if Pos('gopher',URL) > 0 then result := true;

  end;

var
  i : Integer;
  IEURL : String;
  IEHWND : HWND;
  IETitle : String;
  CurrWin : TShellWindow;
  found : Boolean;

begin
{$IFDEF COOBJS}
  if IE = nil then exit;
{$ENDIF}
  try
    IEURL := IE.LocationURL;
    IEURL := ReplaceSTr(IEURL,'%20',' ');
    IETITLE := IE.LocationName;
    if (FLoggingDirs) or (IsValidURL(IEURL)) then
      begin
        IEHWND := IE.HWND;
        found := false;
        for i := 0 to FWindowList.Count-1 do
          begin
            CurrWin := TShellWindow(FWindowList.Items[i]);
            if (CurrWin.handle = IEHWND) then
              begin
                found := true;
                if (IEURL <> CurrWin.LocationURL) then
                  begin
                    CurrWin.LocationURL := {ReplaceSTr(IEURL,'%20',' ')}IEURL;
                    CurrWin.LocationName := IETitle;
                    AddEntry(CurrWin);
                  end;
              end;
          end;
        if not found then
          AddBrowserId(IEURL, IETITLE, IE.HWND);
      end;
  except on EOleException do end;
end;

//----------------------------------------------------------------------------//
//                                RemoveZombies                               //
//----------------------------------------------------------------------------//
procedure TShellWindowObserver.RemoveZombies;
var
  i, j : integer;
  ih : HWND;
  found, changed  : boolean;
  Num: Integer;

begin
  i := 0;
  changed := False;
  Num := 1;
  while i < FWindowList.Count  do
    begin
      ih := TShellWindow(FWindowList.Items[i]).handle;
      found := false;
      try
        for j := 0 to Pred(SH.Count) do
          begin
            IE := SH.Item(j) {$IFDEF COOBJS} as IWebBrowser2 {$ENDIF};
            if IE = nil then continue;
            if HWND(IE.HWND) = ih then
              begin
                found := True;
                break
              end;
          end;
      except on EOleSysError do
        found := false
      end;
      if not found then
        begin
          changed := true;
          FWindowList.Delete(i);
        end
      else
        begin
          // Correct a bug
          TShellWindow(FWindowList.Items[i]).Num := Num;
          Inc(Num);
          inc(i);
        end;
    end;
  Count := FWindowList.Count;
  if (changed) and (Assigned(FOnChangedNumber)) then
    FOnChangedNumber(Self, FWindowList.Count);
end;

//----------------------------------------------------------------------------//
//                                 UpdateList                                 //
//----------------------------------------------------------------------------//
procedure TShellWindowObserver.UpdateList;
var
  i: integer;

begin
  try
    for i := 0 to Pred(SH.Count) do
      begin
        IE := SH.Item(i) {$IFDEF COOBJS} as IWebBrowser2 {$ENDIF};
        UpdateWins;
      end;
    RemoveZombies;
  except on Exception do end;
end;

//----------------------------------------------------------------------------//
//                             OnTimerObserver                                //
//----------------------------------------------------------------------------//
procedure TShellWindowObserver.OnTimerObserver(Sender : TObject);
begin
  if FActive then
    UpdateList;
end;

//----------------------------------------------------------------------------//
//                                  Create                                    //
//----------------------------------------------------------------------------//
constructor TShellWindowObserver.Create(AOwner : TComponent);
begin
  Init;
  inherited;
  FLoggingDirs := False;
  FDelimiter := ',';
  FNumMessage := SC_CLOSE;
  Count := 0;
  FLoggedChanges := TStringList.Create;
  FWindowList := TList.Create;
  FTimer := TTimer.Create(Self);
  with FTimer do
    begin
      Interval := TIMERINTERVAL;
      Enabled := True;
      OnTimer := OnTimerObserver;
    end;
end;

//----------------------------------------------------------------------------//
//                                 Destroy                                    //
//----------------------------------------------------------------------------//
destructor TShellWindowObserver.Destroy;
var
  i: Integer;

begin
  FLoggedChanges.Free;
  for i := 0 to Pred(FWindowList.Count) do
    TShellWindow(FWindowList.Items[i]).Free;
  FWindowList.Free;
  FTimer.Free;
  inherited;
end;

//----------------------------------------------------------------------------//
//                                 Register                                   //
//----------------------------------------------------------------------------//
procedure Register;
begin
  RegisterComponents('VCL Internet', [TShellWindowObserver]);
end;

end.
