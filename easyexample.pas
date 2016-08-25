unit EasyExample;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, V8AddIn, fpTimer, FileUtil;

type

  { TTimerThread }

  TTimerThread = class (TThread)
  public
    Interval: Integer;
    OnTimer: TNotifyEvent;
    procedure Execute; override;
  end;

  { TAddInEasyExample }

  TAddInEasyExample = class (TAddIn)
  private
    FIsEnabled: Boolean;
    FTimer: TTimerThread;
    function GetIsEnabled: Variant;
    function GetIsTimerPresent: Variant;
    procedure MyTimerProc(Sender: TObject);
    procedure SetIsEnabled(AValue: Variant);
  public
    property IsEnabled: Variant read GetIsEnabled write SetIsEnabled;
    property IsTimerPresent: Variant read GetIsTimerPresent;
    procedure Enable;
    procedure Disable;
    procedure ShowInStatusLine(Text: Variant);
    procedure StartTimer;
    procedure StopTimer;
    destructor Destroy; override;
    function LoadPicture(FileName: Variant): Variant;
    procedure ShowMessageBox;
  end;

implementation

{ TTimerThread }

procedure TTimerThread.Execute;
begin
  repeat
    Sleep(Interval);
    if Terminated then
      Break
    else
      OnTimer(Self);
  until False;
end;

{ TAddInEasyExample }

function TAddInEasyExample.GetIsEnabled: Variant;
begin
  Result := FIsEnabled;
end;

function TAddInEasyExample.GetIsTimerPresent: Variant;
begin
  Result := True;
end;

procedure TAddInEasyExample.MyTimerProc(Sender: TObject);
begin
  ExternalEvent('EasyComponentNative', 'Timer', DateTimeToStr(Now));
end;

procedure TAddInEasyExample.SetIsEnabled(AValue: Variant);
begin
  FIsEnabled := AValue;
end;

procedure TAddInEasyExample.Enable;
begin
  FIsEnabled := True;
end;

procedure TAddInEasyExample.Disable;
begin
  FIsEnabled := False;
end;

procedure TAddInEasyExample.ShowInStatusLine(Text: Variant);
begin
  SetStatusLine(Text);
end;

procedure TAddInEasyExample.StartTimer;
begin
  if FTimer = nil then
    begin
      FTimer := TTimerThread.Create(True);
      FTimer.Interval := 1000;
      FTimer.OnTimer := @MyTimerProc;
      FTimer.Start;
    end;
end;

procedure TAddInEasyExample.StopTimer;
begin
  if FTimer <> nil then
    FreeAndNil(FTimer);
end;

destructor TAddInEasyExample.Destroy;
begin
  if FTimer <> nil then
    FreeAndNil(FTimer);
  inherited Destroy;
end;

function TAddInEasyExample.LoadPicture(FileName: Variant): Variant;
var
  SFileName: String;
  FileStream: TFileStream;
begin
  SFileName := UTF8ToSys(String(FileName));
  FileStream := TFileStream.Create(SFileName, fmOpenRead);
  try
    Result := StreamTo1CBinaryData(FileStream);
  finally
    FileStream.Free;
  end;
end;

procedure TAddInEasyExample.ShowMessageBox;
begin
  if Confirm(AppVersion) then
    Alert('OK')
  else
    Alert('Cancel');
end;

initialization

  with TAddInEasyExample do
    begin
      RegisterClass('AddInNativeEasyExtension', 2000);
      AddProp('IsEnabled', 'Включен', @GetIsEnabled, @SetIsEnabled);
      AddProp('IsTimerPresent', 'ЕстьТаймер', @GetIsTimerPresent);
      AddProc('Enable', 'Включить', @Enable);
      AddProc('Disable', 'Выключить', @Disable);
      AddProc('ShowInStatusLine', 'ПоказатьВСтрокеСтатуса', @ShowInStatusLine);
      AddProc('StartTimer', 'СтартТаймер', @StartTimer);
      AddProc('StopTimer', 'СтопТаймер', @StopTimer);
      AddFunc('LoadPicture', 'ЗагрузитьКартинку', @LoadPicture);
      AddProc('ShowMessageBox', 'ПоказатьСообщение', @ShowMessageBox);
    end;

end.

