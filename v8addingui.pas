unit V8AddInGUI;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Interfaces, InterfaceBase, Forms;

type

  { TGUIThread }

  TGUIThread = class (TThread)
    procedure Execute; override;
    procedure Synchronize(AMethod: TThreadMethod);
  end;

function GUIThread: TGUIThread;

implementation

type

  { TFormController }

  TFormController = class
    Form: TForm;
    FormClass: TFormClass;
    procedure CreateForm;
    procedure FreeForm;
    procedure ShowForm;
    procedure HideForm;
  end;

var
  GUICriticalSection: TRTLCriticalSection;
  FGUIThread: TGUIThread;
  FakeMainForm: TForm;

function GUIThread: TGUIThread;
begin
  if FGUIThread = nil then
    begin
      EnterCriticalsection(GUICriticalSection);
      try
        Application.Initialize;
        if FGUIThread = nil then
          FGUIThread := TGUIThread.Create(False);
      finally
        LeaveCriticalSection(GUICriticalSection);
      end;
      Result := FGUIThread;
    end
  else
    Result := FGUIThread;
end;

{ TFormController }

procedure TFormController.CreateForm;
begin
  Form := FormClass.Create(nil);
end;

procedure TFormController.FreeForm;
begin
  FreeAndNil(Form);
end;

procedure TFormController.ShowForm;
begin
  Form.Show;
end;

procedure TFormController.HideForm;
begin
  Form.Hide;
end;

{ TGUIThread }

procedure TGUIThread.Execute;
begin
  Application.ShowMainForm := False;
  Application.CreateForm(TForm, FakeMainForm);
  repeat
    WidgetSet.AppProcessMessages;
    if not Terminated then
      begin
        Application.Idle(False);
        Sleep(100);
      end;
  until Terminated;
  FakeMainForm.Free;
end;

procedure TGUIThread.Synchronize(AMethod: TThreadMethod);
begin
  TThread.Synchronize(Self, AMethod);
end;

initialization

  InitCriticalSection(GUICriticalSection);
  FGUIThread := nil;

finalization

  if FGUIThread <> nil then
    FGUIThread.Free;
  DoneCriticalSection(GUICriticalSection);

end.

