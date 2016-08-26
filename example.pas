unit Example;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, V8AddIn, fpTimer, FileUtil;

type

  TProps = (
    ePropIsEnabled = 0,
    ePropIsTimerPresent
  );

  TMethods = (
    eMethEnable = 0,
    eMethDisable,
    eMethShowInStatusLine,
    eMethStartTimer,
    eMethStopTimer,
    eMethLoadPicture,
    eMethShowMsgBox
  );

const
  PropNames: array [TProps] of string = (
    'IsEnabled', 'IsTimerPresent');
  MethodNames: array [TMethods] of string = (
    'Enable', 'Disable', 'ShowInStatusLine', 'StartTimer',
    'StopTimer', 'LoadPicture', 'ShowMessageBox');

  PropNamesRu: array [TProps] of string = (
    'Включен', 'ЕстьТаймер');
  MethodNamesRu: array [TMethods] of string = (
    'Включить', 'Выключить', 'ПоказатьВСтрокеСтатуса', 'СтартТаймер',
    'СтопТаймер', 'ЗагрузитьКартинку', 'ПоказатьСообщение');

type

  { TTimerThread }

  TTimerThread = class (TThread)
  public
    Interval: Integer;
    OnTimer: TNotifyEvent;
    procedure Execute; override;
  end;

  { TAddInExample }

  TAddInExample = class (TAddIn)
  private
    FEnabled: Boolean;
    FTimer: TTimerThread;
    function FindName(const Names: array of String; const Name: String): Integer;
    function GetName(const Names: array of String; const Index: Integer): String;
    procedure MyTimerProc(Sender: TObject);
  public
    function GetInfo: Integer; override;
    procedure Done; override;
    function RegisterExtensionAs(var ExtensionName: String): Boolean; override;
    function GetNProps: Integer; override;
    function FindProp(const PropName: String): Integer; override;
    function GetPropName(const PropNum, PropAlias: Integer): String; override;
    function GetPropVal(const PropNum: Integer; var Value: Variant): Boolean; override;
    function SetPropVal(const PropNum: Integer; const Value: Variant): Boolean; override;
    function IsPropReadable(const PropNum: Integer): Boolean; override;
    function IsPropWritable(const PropNum: Integer): Boolean; override;
    function GetNMethods: Integer; override;
    function FindMethod(const AMethodName: String): Integer; override;
    function GetMethodName(const MethodNum, MethodAlias: Integer): String; override;
    function GetNParams(const MethodNum: Integer): Integer; override;
    function HasRetVal(const MethodNum: Integer): Boolean; override;
    function CallAsProc(const MethodNum: Integer; const Params: array of Variant): Boolean; override;
    function CallAsFunc(const MethodNum: Integer; var RetValue: Variant; const Params: array of Variant): Boolean; override;
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

{ TAddInExample }

function TAddInExample.FindName(const Names: array of String; const Name: String
  ): Integer;
begin
  for Result := 0 to Length(Names) - 1 do
    if AnsiCompareText(Names[Result], Name) = 0 then
      Exit;
  Result := -1;
end;

function TAddInExample.GetName(const Names: array of String;
  const Index: Integer): String;
begin
  if (Index < 0) or (Index >= Length(Names)) then
    Result := ''
  else
    Result := Names[Index];
end;

procedure TAddInExample.MyTimerProc(Sender: TObject);
begin
  ExternalEvent('ComponentNative', 'Timer', DateTimeToStr(Now));
end;

function TAddInExample.GetInfo: Integer;
begin
  Result := 2000;
end;

procedure TAddInExample.Done;
begin
  if FTimer <> nil then
    FreeAndNil(FTimer);
end;

function TAddInExample.RegisterExtensionAs(var ExtensionName: String): Boolean;
begin
  ExtensionName := 'AddInNativeExtension';
  Result := True;
end;

function TAddInExample.GetNProps: Integer;
begin
  Result := Ord(High(TProps)) + 1;
end;

function TAddInExample.FindProp(const PropName: String): Integer;
begin
  Result := FindName(PropNames, PropName);
  if Result < 0 then
    Result := FindName(PropNamesRu, PropName);
end;

function TAddInExample.GetPropName(const PropNum, PropAlias: Integer): String;
begin
  case PropAlias of
    0:
      Result := GetName(PropNames, PropNum);
    1:
      Result := GetName(PropNamesRu, PropNum);
  else
    Result := '';
  end;
end;

function TAddInExample.GetPropVal(const PropNum: Integer; var Value: Variant
  ): Boolean;
begin
  Result := True;
  case TProps(PropNum) of
    ePropIsEnabled:
      Value := FEnabled;
    ePropIsTimerPresent:
      Value := True;
  else
    Result := False;
  end;
end;

function TAddInExample.SetPropVal(const PropNum: Integer; const Value: Variant
  ): Boolean;
begin
  Result := True;
  case TProps(PropNum) of
    ePropIsEnabled:
      FEnabled := Value;
    ePropIsTimerPresent:;
  else
    Result := False;
  end;
end;

function TAddInExample.IsPropReadable(const PropNum: Integer): Boolean;
begin
  case TProps(PropNum) of
    ePropIsEnabled,
    ePropIsTimerPresent:
      Result := True;
  else
    Result := False;
  end;
end;

function TAddInExample.IsPropWritable(const PropNum: Integer): Boolean;
begin
  case TProps(PropNum) of
    ePropIsEnabled:
      Result := True;
    ePropIsTimerPresent:
      Result := False;
  else
    Result := False;
  end;
end;

function TAddInExample.GetNMethods: Integer;
begin
  Result := Ord(High(TMethods)) + 1;
end;

function TAddInExample.FindMethod(const AMethodName: String): Integer;
begin
  Result := FindName(MethodNames, AMethodName);
  if Result < 0 then
    Result := FindName(MethodNamesRu, AMethodName);
end;

function TAddInExample.GetMethodName(const MethodNum, MethodAlias: Integer
  ): String;
begin
  case MethodAlias of
    0:
      Result := GetName(MethodNames, MethodNum);
    1:
      Result := GetName(MethodNamesRu, MethodNum);
  else
    Result := '';
  end;
end;

function TAddInExample.GetNParams(const MethodNum: Integer): Integer;
begin
  case TMethods(MethodNum) of
    eMethShowInStatusLine,
    eMethLoadPicture:
      Result := 1;
  else
    Result := 0;
  end;
end;

function TAddInExample.HasRetVal(const MethodNum: Integer): Boolean;
begin
  case TMethods(MethodNum) of
    eMethLoadPicture:
      Result := True;
  else
    Result := False;
  end;
end;

function TAddInExample.CallAsProc(const MethodNum: Integer;
  const Params: array of Variant): Boolean;
var
  RetValue: Variant;
begin
  Result := True;
  case TMethods(MethodNum) of
    eMethEnable:
      FEnabled := True;
    eMethDisable:
      FEnabled := False;
    eMethShowInStatusLine:
      SetStatusLine(Params[0]);
    eMethStartTimer:
      begin
        if FTimer = nil then
          begin
            FTimer := TTimerThread.Create(True);
            FTimer.Interval := 1000;
            FTimer.OnTimer := @MyTimerProc;
            FTimer.Start;
          end;
      end;
    eMethStopTimer:
      begin
        if FTimer <> nil then
          FreeAndNil(FTimer);
      end;
    eMethShowMsgBox:
      begin
        RetValue := Unassigned;
        if Confirm(AppVersion, RetValue) then
          begin
            if RetValue then
              Alert('OK')
            else
              Alert('Cancel');
          end;
      end;
  else
    Result := False;
  end;
end;

function TAddInExample.CallAsFunc(const MethodNum: Integer;
  var RetValue: Variant; const Params: array of Variant): Boolean;
var
  FileName: String;
  FileStream: TFileStream;
begin
  Result := True;
  case TMethods(MethodNum) of
    eMethLoadPicture:
      begin
        FileName := UTF8ToSys(String(Params[0]));
        try
          FileStream := TFileStream.Create(FileName, fmOpenRead);
          try
            RetValue := StreamTo1CBinaryData(FileStream);
          finally
            FileStream.Free;
          end;
        except
          on E: Exception do
            begin
              Result := False;
              AddError(ADDIN_E_VERY_IMPORTANT, 'AddInNative', E.Message, 0);
            end;
        end;
      end;
  else
    Result := False;
  end;
end;

end.

