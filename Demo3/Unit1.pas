unit Unit1;

{*****************************************************************************
 *
 *  uShellWindowObserver.pas demo - Shell Window Observer Component - Version 1.2
 *
 *  Created by Michel Hibon
 *****************************************************************************}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls,  uShellWindowObserver, Spin, Buttons, ExtCtrls;

type
  TForm1 = class(TForm)
    ShellWindowObserver1: TShellWindowObserver;
    Memo1: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Panel1: TPanel;
    Label6: TLabel;
    SpinEdit1: TSpinEdit;
    Label7: TLabel;
    ComboBox1: TComboBox;
    Button2: TButton;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    CheckBox1: TCheckBox;
    procedure ShellWindowObserver1AddedEntry(Sender: TObject;
      ShellWindow: TShellWindow; Str : String);
    procedure FormShow(Sender: TObject);
    procedure ShellWindowObserver1ChangedNumber(Sender: TObject;
      Number: Integer);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

const
  M1    = ' *  uShellWindowObserver.pas - Shell Window Observer Component - Version 1.1';
  M2    = ' *';
  M3    = ' *  Copyright (c) 2000 Diego Amicabile';
  M4    = ' *  Author:     Diego Amicabile';
  M5    = ' *  E-mail:     diegoami@yahoo.it';
  M6    = ' *  Homepage:   http://www.geocities.com/diegoami';
  M7    = ' *  This component is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License';
  M8    = ' *  as published by the Free Software Foundation. This component is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY';
  M9    = ' *  without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.';
  M10   = ' *  See the GNU General Public License for more details.';
  M11   = ' *  You should have received a copy of the GNU General Public License along with this component; if not, write to the :';
  M12   = ' *  Free Software Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA';
  M13   = ' *  TSHELLWINDOWOBSERVER';
  M14   = ' *  This component was born out of frustration, since I wasn''t able to connect to an existing instance of Internet Explorer AND catch the events it fires.';
  M15   = ' *  Using this component you can, and you can perform actions when a new instance is started or an existing instance moves to a new location.';
  M16   = ' *  There must be a better way to do it, let me know if you have figured it out. I use timers.';
  M17   = ' *  See help.htm for documentation.';
  M18   = ' *  See license.txt for license agreements.';
  M19   = ' *  REQUIREMENTS';
  M20   = ' *  THe Internet Explorer Active X Component must have been imported (it is by default in Delphi 5)';
  M21   = ' *  You may need to replace SHDOCVW with SHDOCVW_TLB';
  M22   = ' *  INSTALLING';
  M23   = ' *  Put the uShellWindowObserver.pas file in a package and compile it.';
  M24   = ' *  IMPORTANT : You may have to remove the line {$DEFINE COOBJS} if you are using Delphi 3.';
  M25   = ' *  NOTE';
  M26   = ' *  The component will be installed in any version of Delphi (3 or higher). The demo was created using Delphi 5,';
  M27   = ' *  so the .DFM file may have to be converted.';
  M28   = ' *  Created by Diego Amicabile';
  M29   = ' *  Updated by Michel Hibon - 08/27/2001';
  M30   = ' **********************************************************************************************************************************************************************************';
  M31   = ' *  Like Diego, i was researched a method to get the handle of an explorer''s instance. When you create a new explorere''s instance';
  M32   = ' *  with CreateProcess API Function, the PROCESS_INFORMATION returns false data. Why ? I don''t know ! I''m found this excellent ';
  M33   = ' *  component, and i have decided to integrate it in my program. But, a little bug appears. I correct it, and i add some improvements :';
  M34   = ' *  Count : The number of active instance';
  M35   = ' *  function SendMessageToWindowByNumber(Num: Integer): LRESULT; : Allow you to send a message to choosen instance';
  M36   = ' *  property SendAMessage : Allow you to choose the type of message you want to send at the choosen instance';
  M37   = ' *  I design the demo program for test this component with this improvements. Component and demo program was updated with Delphi 6.';
  M38   = ' *  Name:     Michel Hibon (alias : Mickey)';
  M39   = ' *  E-mail:     mhibon@ifrance.com';
  M40   = ' *  Homepage:   http://www.guetali.fr/home/mhibon';
var
  Form1: TForm1;

implementation

{$R *.DFM}

//----------------------------------------------------------------------------//
//                            GetShellWindowOnNumber                          //
//----------------------------------------------------------------------------//
procedure TForm1.ShellWindowObserver1AddedEntry(Sender: TObject;
  ShellWindow: TShellWindow; Str : String);
begin
  Memo1.Lines.Add(Str)
end;

//----------------------------------------------------------------------------//
//                                    FormShow                                //
//----------------------------------------------------------------------------//
procedure TForm1.FormShow(Sender: TObject);
var
  prova, oriv : String;

begin
  Memo1.Lines.Clear;
  prova := 'A%20BB%20jgoer%20BBERF';    // ???
  oriv := ReplaceStr(prova,'%20',' ');  // Test function ReplaceStr ???
//  ShowMessage(oriv);                    // and show the result ???
end;

//----------------------------------------------------------------------------//
//                       ShellWindowObserver1ChangedNumber                    //
//----------------------------------------------------------------------------//
procedure TForm1.ShellWindowObserver1ChangedNumber(Sender: TObject;
  Number: Integer);
begin
  Memo1.Lines.Add('Now '+IntToStr(Number)+ ' instances running. ');
  SpinEdit1.MaxValue := Number;
  if SpinEdit1.Value > Number then SpinEdit1.Value := Number;
  if ShellwindowObserver1.Count = 0 then
    begin
      ComboBox1.Enabled := False;
      SpinEdit1.Enabled := False;
      Button2.Enabled := False;
    end
  else
    begin
      ComboBox1.Enabled := True;
      SpinEdit1.Enabled := True;
      Button2.Enabled := True;
    end;
  Button1Click(Sender);
end;

//----------------------------------------------------------------------------//
//                                   Button1Click                             //
//----------------------------------------------------------------------------//
procedure TForm1.Button1Click(Sender: TObject);
var
  CW : TShellWindow;

begin
  if SpinEdit1.Value = 0 then
    begin
      Label1.Caption := '';
      Label2.Caption := '';
      Label3.Caption := '';
      Label4.Caption := '';
      Label5.Caption := '';
      Exit;
    end;
  CW := ShellwindowObserver1.GetShellWindowOnNumber(SpinEdit1.Value);
  Label1.Caption := 'Num = ' + IntToStr(CW.Num);
  Label2.Caption := 'Handle = ' + IntToHex(CW.handle,8);
  Label3.Caption := 'Name = ' + CW.LocationName;
  Label4.Caption := 'URL = ' + CW.LocationURL;
  Label5.Caption := 'Class = ' + CW.ClassName;
end;

//----------------------------------------------------------------------------//
//                                  Button2Click                              //
//----------------------------------------------------------------------------//
procedure TForm1.Button2Click(Sender: TObject);
begin
  if SpinEdit1.Value = 0 then
    begin
      Label1.Caption := '';
      Label2.Caption := '';
      Label3.Caption := '';
      Label4.Caption := '';
      Label5.Caption := '';
      Exit;
    end;
  case ComboBox1.ItemIndex of
    0 : ShellwindowObserver1.SendAMessage := smClose;
    1 : ShellwindowObserver1.SendAMessage := smMaximize;
    2 : ShellwindowObserver1.SendAMessage := smMinimize;
    3 : ShellwindowObserver1.SendAMessage := smRestore;
  end;
  ShellWindowObserver1.SendMessageToWindowByNumber(SpinEdit1.Value)
end;

//----------------------------------------------------------------------------//
//                                SpinEdit1Change                             //
//----------------------------------------------------------------------------//
procedure TForm1.SpinEdit1Change(Sender: TObject);
begin
  Button1Click(Sender);
end;

//----------------------------------------------------------------------------//
//                               SpeedButton2Click                            //
//----------------------------------------------------------------------------//
procedure TForm1.SpeedButton2Click(Sender: TObject);
begin
  Close;
end;

//----------------------------------------------------------------------------//
//                                  FormCreate                                //
//----------------------------------------------------------------------------//
procedure TForm1.FormCreate(Sender: TObject);
begin
  if ShellwindowObserver1.Count = 0 then
    begin
      ComboBox1.Enabled := False;
      SpinEdit1.Enabled := False;
      Button2.Enabled := False;
    end
  else
    begin
      ComboBox1.Enabled := True;
      SpinEdit1.Enabled := True;
      Button2.Enabled := True;
    end;
end;

//----------------------------------------------------------------------------//
//                              SpeedButton1Click                             //
//----------------------------------------------------------------------------//
procedure TForm1.SpeedButton1Click(Sender: TObject);
begin
  ShellWindowObserver1.Active := False;
  CheckBox1.Checked := False;
  Memo1.Height := 184;
  Memo1.Clear;
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M1);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M30);
  Memo1.Lines.Add(M28);
  Memo1.Lines.Add(M29);
  Memo1.Lines.Add(M30);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M3);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M3);
  Memo1.Lines.Add(M4);
  Memo1.Lines.Add(M5);
  Memo1.Lines.Add(M6);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M7);
  Memo1.Lines.Add(M8);
  Memo1.Lines.Add(M9);
  Memo1.Lines.Add(M10);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M11);
  Memo1.Lines.Add(M12);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M13);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M14);
  Memo1.Lines.Add(M15);
  Memo1.Lines.Add(M16);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M17);
  Memo1.Lines.Add(M18);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M19);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M20);
  Memo1.Lines.Add(M21);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M22);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M23);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M24);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M25);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M26);
  Memo1.Lines.Add(M27);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M30);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M31);
  Memo1.Lines.Add(M32);
  Memo1.Lines.Add(M33);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M34);
  Memo1.Lines.Add(M35);
  Memo1.Lines.Add(M36);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M37);
  Memo1.Lines.Add(M2);//
  Memo1.Lines.Add(M38);
  Memo1.Lines.Add(M39);
  Memo1.Lines.Add(M40);
  Memo1.SelStart := 0;
  Memo1.SelLength := 0;
end;

//----------------------------------------------------------------------------//
//                               CheckBox1Click                               //
//----------------------------------------------------------------------------//
procedure TForm1.CheckBox1Click(Sender: TObject);
begin
  ShellWindowObserver1.Active := CheckBox1.Checked;
  SpinEdit1.Enabled := CheckBox1.Checked;
  ComboBox1.Enabled := CheckBox1.Checked;
  Button2.Enabled := CheckBox1.Checked;
  if Memo1.Height = 184 then
    begin
      Memo1.Height := 84;
      Memo1.Clear;
      ShellWindowObserver1ChangedNumber(Sender,ShellWindowObserver1.Count);
    end;
  if CheckBox1.Checked then
    begin
      if ShellwindowObserver1.Count = 0 then
        begin
          ComboBox1.Enabled := False;
          SpinEdit1.Enabled := False;
          Button2.Enabled := False;
        end
      else
        begin
          ComboBox1.Enabled := True;
          SpinEdit1.Enabled := True;
          Button2.Enabled := True;
        end;
    end;
end;

end.
